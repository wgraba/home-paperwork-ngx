#!/usr/bin/env python3

from ctypes.wintypes import tagMSG
import datetime
import json
import logging
import os
import pathlib
import requests
import subprocess
import typer

from requests.auth import AuthBase
from requests import models
from rich.progress import track
from rich import print
from rich.logging import RichHandler

logging.basicConfig(handlers=[RichHandler()])

logger = logging.getLogger()


class PaperlessAuth(AuthBase):
    """
    Attaches `Authentication: Token <token>` to a Request object
    """

    def __init__(self, token: str) -> None:
        super().__init__()

        self._token = token

    def __call__(self, r: models.PreparedRequest) -> models.PreparedRequest:
        r.headers["Authorization"] = f"Token {self._token}"
        return r


paperwork_json_cmd = [
    "flatpak",
    "run",
    "--command=paperwork-json",
    "work.openpaper.Paperwork",
]

paperwork_cli_cmd = [
    "flatpak",
    "run",
    "--command=paperwork-cli",
    "work.openpaper.Paperwork",
]


def main(
    paperwork_path: str = typer.Argument(..., help="Path to paperwork archive"),
    paperless_url: str = typer.Argument(..., help="URL to paperless-ngx instance"),
    paperless_token: str = typer.Argument(
        ..., help="paperless-ngx Token to use for authentication"
    ),
    tmp_path: str = typer.Argument(..., help="Path to store exported PDFs"),
    dryrun: bool = False,
):
    """
    Migrate paperwork documents to paperless-ngx.

    paperwork-json is used to get all of the documents, export them, and get
    the labels.

    The paperless-ngx REST API is used to consume these docs.

    NOTE: Only a flatpak install of Paperwork is currently supported.
    NOTE: The labels of docs in Paperwork are assumed to be the same as tags in
          paperless-ngx.
    """

    # Check paths exist
    pwork_path = pathlib.Path(paperwork_path)
    if not pwork_path.exists():
        raise ValueError(f"'{pwork_path}' does not exist!")

    tpath = pathlib.Path(tmp_path)
    if not tpath.exists():
        raise ValueError(f"{tpath}' does not exist!")

    # Check paperless-ngx instance exists
    req_session = requests.Session()
    req_session.auth = PaperlessAuth(paperless_token)
    # req_session.verify = False

    rsp = req_session.get(f"{paperless_url}/api/")
    if not rsp.ok:
        raise RuntimeError(
            f"Unable to communication with paperless-ngx instance at '{rsp.url}'"
        )

    print(
        f"Found paperless-ngx v{rsp.headers['X-Version']} instance with API v{rsp.headers['X-Api-Version']}"
    )

    # Get mapping of tag names to ids
    tag_ids = {}
    rsp = req_session.get(f"{paperless_url}/api/tags/")
    rsp.raise_for_status()

    for tag in rsp.json()["results"]:
        tag_ids[tag["slug"]] = tag["id"]

    print(f"Getting docs from paperwork archive at '{pwork_path}'...")
    for f in track(os.scandir(pwork_path), description="Processing docs..."):
        if f.is_dir():
            print(f"Processing {f.name}...")

            export_path = tpath.joinpath(f"./{f.name}.pdf")
            if export_path.exists():
                print(f"{export_path} already exists; assume already migrated")
                continue

            doc_date = datetime.datetime.strptime(f.name.split("_")[0], "%Y%m%d")

            # Get Labels
            cp = subprocess.run(
                paperwork_json_cmd + ["show", f.name], check=True, capture_output=True
            )
            info = json.loads(cp.stdout)

            labels = []
            try:
                for label_el in info["document"]["labels"]:
                    labels.append(label_el["label"])

            except KeyError:
                logger.warning(f"No labels for {f.name}")

            print(f"Date: {doc_date}, Labels: {labels}")

            # Export PDF
            print(f"Exporting to {export_path}...")

            # Determine available filters
            cp = subprocess.run(
                paperwork_json_cmd
                + [
                    "export",
                    f.name,
                ],
                check=True,
                capture_output=True,
            )
            filters = json.loads(cp.stdout)

            # Prefere unmodified PDF
            if "unmodified_pdf" in filters:
                subprocess.run(
                    paperwork_json_cmd
                    + [
                        "export",
                        f.name,
                        "--filter",
                        "unmodified_pdf",
                        "--out",
                        export_path.absolute(),
                    ],
                    check=True,
                    capture_output=True,
                )

            elif "doc_to_pages" in filters:
                subprocess.run(
                    paperwork_json_cmd
                    + [
                        "export",
                        f.name,
                        "--filter",
                        "doc_to_pages",
                        "--filter",
                        "img_boxes",
                        "--filter",
                        "generated_pdf",
                        "--out",
                        export_path.absolute(),
                    ],
                    check=True,
                    capture_output=True,
                )

            else:
                logger.error(f"Unknown filters {filters}!")
                # raise RuntimeError(f"Unknown filters {filters}!")

            if not dryrun:
                # Push to paperless-ngx instance
                print(f"Uploading {export_path} to paperless-ngx server...")
                payload_tags = []
                for label in labels:
                    if label.lower() in tag_ids:
                        payload_tags.append(tag_ids[label.lower()])

                payload = {
                    "created": doc_date.strftime("%Y-%m-%d"),
                    "tags": payload_tags,
                }

                documents = {
                    "document": open(export_path, "rb"),
                }
                rsp = req_session.post(
                    f"{paperless_url}/api/documents/post_document/",
                    data=payload,
                    files=documents,
                )
                if not rsp.ok:
                    raise RuntimeError(
                        f"Unable to submit '{export_path}' to paperless-ngx!"
                    )


if __name__ == "__main__":
    typer.run(main)
