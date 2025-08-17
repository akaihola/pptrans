"""Type stubs for the Pyodide ``js`` module."""

from __future__ import annotations

from typing import IO, TYPE_CHECKING, Self

from pyodide.ffi import JsDomElement

if TYPE_CHECKING:
    from collections.abc import Awaitable


class ArrayBuffer: ...


class Blob:
    def arrayBuffer(self) -> Awaitable[ArrayBuffer]: ...


class HTMLAllCollection:
    def item(self, index: int) -> Blob: ...


class FileList(HTMLAllCollection): ...


class HTMLInputElement:
    files: FileList


class Event:
    target: HTMLInputElement


class UIEvent(Event): ...


class InputEvent(UIEvent): ...


class Uint8Array:
    @classmethod
    def new(cls, data: ArrayBuffer) -> Self: ...
    def to_file(self, f: IO[bytes] | IO[str]) -> None: ...


class Element(JsDomElement):
    innerHTML: str


class Document:
    def getElementById(self, elementId: str) -> Element: ...


document = Document()
