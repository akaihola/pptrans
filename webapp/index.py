"""Python scripts for the example web app."""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING, cast

import click
from pyodide.ffi.wrappers import add_event_listener

from js import Uint8Array, document
from pptrans.__main__ import main

if TYPE_CHECKING:
    from typing import Callable

    from js import InputEvent


def _write_input_file(data: Uint8Array) -> None:
    """Write data to the input file."""
    with Path("input.pptx").open("wb") as input_file:
        data.to_file(input_file)


async def handle_file_upload(event: InputEvent) -> None:
    """Handle file upload event."""
    file = event.target.files.item(0)
    file_bytes = await file.arrayBuffer()
    data = Uint8Array.new(file_bytes)
    _write_input_file(data)
    ctx = click.Context(main)
    ctx.forward(main, input_path="input.pptx", output_path="output.pptx")
    output = document.getElementById("output")
    output.innerHTML = "<a href='output.pptx'>Download Translated Presentation</a>"


add_event_listener(
    document.getElementById("file-upload"),
    "change",
    cast("Callable[[InputEvent], None]", handle_file_upload),
)
