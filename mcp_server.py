"""Model Context Protocol server exposing SharePoint content."""
from __future__ import annotations

import argparse
import base64
import logging
from functools import lru_cache
from typing import Annotated, Any, Literal

from mcp.server.fastmcp import FastMCP

from mcp_sharepoint import SharePointDownload, SharePointService, SharePointSettings

logger = logging.getLogger("mcp_sharepoint.server")


def configure_logging(verbose: bool) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(level=level, format="%(asctime)s %(levelname)s %(name)s: %(message)s")


@lru_cache(maxsize=1)
def get_service() -> SharePointService:
    try:
        settings = SharePointSettings.from_env()
    except RuntimeError as exc:
        logger.error("SharePoint settings are missing: %s", exc)
        raise
    logger.debug("Loaded SharePoint settings for site %s", settings.site_url)
    return SharePointService(settings)


server = FastMCP("sharepoint", version="0.1.0")


def _serialize_file(file_obj: SharePointDownload, *, include_content: bool = False) -> dict[str, Any]:
    base = {
        "name": file_obj.file.name,
        "path": file_obj.file.server_relative_path,
        "size": file_obj.file.size,
        "download_url": file_obj.file.download_url,
    }
    if include_content:
        content = file_obj.content
        if isinstance(content, str):
            base["content_type"] = "text"
            base["text"] = content
        else:
            base["content_type"] = "base64"
            base["base64"] = base64.b64encode(content).decode("ascii")
    return base


@server.tool()
async def list_folder(
    path: Annotated[str | None, "Server-relative folder path (defaults to configured library)"] = None,
) -> dict[str, Any]:
    """Return folders and files inside a SharePoint document library folder."""

    svc = get_service()
    items = svc.list_folder(path)
    return {
        "path": path or svc.default_library,
        "items": [
            {
                "name": item.name,
                "path": item.server_relative_path,
                "type": "folder" if item.is_folder else "file",
                "size": item.size,
                "download_url": item.download_url,
            }
            for item in items
        ],
    }


@server.tool()
async def download_file(
    path: Annotated[str, "Server-relative file path inside the SharePoint site"],
    mode: Annotated[
        Literal["auto", "text", "binary"],
        "Set to 'text' to force utf-8 decode, 'binary' to force base64 response",
    ] = "auto",
) -> dict[str, Any]:
    """Download a file from SharePoint and return its content."""

    svc = get_service()
    if mode == "text":
        download = svc.download_file(path, as_text=True)
    elif mode == "binary":
        download = svc.download_file(path, as_text=False)
    else:
        download = svc.download_file(path)
    return _serialize_file(download, include_content=True)


@server.tool()
async def upload_file(
    folder: Annotated[str | None, "Destination folder; defaults to configured library"],
    name: Annotated[str, "Filename to create inside the folder"],
    payload: Annotated[str, "Either raw text or base64-encoded binary payload"],
    is_base64: Annotated[bool, "Pass true when payload is base64 encoded"] = False,
) -> dict[str, Any]:
    """Upload a file to SharePoint."""

    svc = get_service()
    raw_folder = folder or svc.default_library
    data = base64.b64decode(payload) if is_base64 else payload.encode("utf-8")
    file_obj = svc.upload_file(raw_folder, name, data)
    return {
        "name": file_obj.name,
        "path": file_obj.server_relative_path,
        "size": file_obj.size,
        "download_url": file_obj.download_url,
    }


@server.tool()
async def resolve_download_url(
    path: Annotated[str, "Server-relative file or folder path"],
) -> dict[str, str]:
    """Return the absolute SharePoint URL for a server-relative path."""

    svc = get_service()
    return {"path": path, "download_url": svc.build_absolute_url(path)}


def main() -> None:
    parser = argparse.ArgumentParser(description="Run the SharePoint MCP server")
    parser.add_argument("--verbose", action="store_true", help="Enable debug logging")
    args = parser.parse_args()
    configure_logging(args.verbose)
    server.run()


if __name__ == "__main__":
    main()
