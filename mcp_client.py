"""Example ChatGPT client that connects to the SharePoint MCP server."""
from __future__ import annotations

import argparse
import os
from typing import Sequence

from openai import OpenAI


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Run a ChatGPT prompt against the SharePoint MCP server")
    parser.add_argument(
        "prompt",
        nargs="*",
        help="Prompt to send to ChatGPT. If omitted, a default SharePoint query is used.",
    )
    parser.add_argument(
        "--model",
        default=os.getenv("OPENAI_MODEL", "gpt-4.1-mini"),
        help="ChatGPT model to use (default: gpt-4.1-mini)",
    )
    parser.add_argument(
        "--temperature",
        type=float,
        default=0.1,
        help="Sampling temperature passed to ChatGPT",
    )
    parser.add_argument(
        "--server-command",
        default=os.getenv("MCP_SERVER_COMMAND", "python"),
        help="Command used to launch the MCP server",
    )
    parser.add_argument(
        "--server-args",
        nargs=argparse.REMAINDER,
        default=os.getenv("MCP_SERVER_ARGS", "mcp_server.py").split(),
        help="Arguments passed to the MCP server command (default: mcp_server.py)",
    )
    return parser


def ensure_prompt(words: Sequence[str]) -> str:
    if words:
        return " ".join(words)
    return "List the files available in our SharePoint document library."


def run_chatgpt_request(
    *,
    model: str,
    prompt: str,
    temperature: float,
    server_command: str,
    server_args: Sequence[str],
) -> str:
    client = OpenAI()
    response = client.responses.create(
        model=model,
        temperature=temperature,
        input=[
            {
                "role": "user",
                "content": [
                    {"type": "input_text", "text": prompt},
                ],
            }
        ],
        attachments=[
            {
                "kind": "mcp",
                "name": "sharepoint",
                "metadata": {
                    "description": "SharePoint document library bridge",
                },
                "server": {
                    "command": server_command,
                    "args": list(server_args),
                },
            }
        ],
    )
    return response.output_text or ""


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    prompt = ensure_prompt(args.prompt)

    if not os.getenv("OPENAI_API_KEY"):
        parser.error("OPENAI_API_KEY environment variable must be set")

    output_text = run_chatgpt_request(
        model=args.model,
        prompt=prompt,
        temperature=args.temperature,
        server_command=args.server_command,
        server_args=args.server_args,
    )
    print(output_text)


if __name__ == "__main__":
    main()
