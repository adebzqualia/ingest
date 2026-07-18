"""Atomic, deterministic output bundle helpers."""

from __future__ import annotations

import csv
from dataclasses import dataclass, field
from datetime import datetime, timezone
import json
from pathlib import Path
import shutil
import tempfile
from typing import Any, Iterable, TextIO

from .utils import json_default


class OutputExistsError(FileExistsError):
    """Raised instead of silently replacing an earlier extraction."""


def default_output_path(source: Path, root: Path | None = None) -> Path:
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    base = root or Path.cwd() / "output"
    return base / f"{source.stem}_{timestamp}"


@dataclass
class AtomicOutput:
    """Write into a sibling temporary directory and rename on success."""

    final_path: Path
    _temp_path: Path | None = None
    _open_files: list[TextIO] = field(default_factory=list)
    _committed: bool = False

    def __enter__(self) -> "AtomicOutput":
        self.final_path = self.final_path.expanduser().resolve()
        self.final_path.parent.mkdir(parents=True, exist_ok=True)
        if self.final_path.exists():
            raise OutputExistsError(
                f"Output directory already exists: {self.final_path}. Choose another --output path."
            )
        self._temp_path = Path(
            tempfile.mkdtemp(
                prefix=f".{self.final_path.name}.tmp-", dir=self.final_path.parent
            )
        ).resolve()
        return self

    @property
    def root(self) -> Path:
        if self._temp_path is None:
            raise RuntimeError("AtomicOutput has not been entered")
        return self._temp_path

    def path(self, relative: str | Path) -> Path:
        target = (self.root / relative).resolve()
        if self.root != target and self.root not in target.parents:
            raise ValueError(f"Output path escapes bundle: {relative}")
        target.parent.mkdir(parents=True, exist_ok=True)
        return target

    def write_json(self, relative: str | Path, value: Any, *, compact: bool = False) -> Path:
        destination = self.path(relative)
        kwargs: dict[str, Any] = {
            "ensure_ascii": False,
            "sort_keys": True,
            "default": json_default,
        }
        if compact:
            kwargs["separators"] = (",", ":")
        else:
            kwargs["indent"] = 2
        destination.write_text(json.dumps(value, **kwargs) + "\n", encoding="utf-8")
        return destination

    def open_text(self, relative: str | Path, *, newline: str | None = "") -> TextIO:
        handle = self.path(relative).open("w", encoding="utf-8", newline=newline)
        self._open_files.append(handle)
        return handle

    def write_jsonl(self, relative: str | Path, rows: Iterable[dict[str, Any]]) -> Path:
        destination = self.path(relative)
        with destination.open("w", encoding="utf-8", newline="\n") as handle:
            for row in rows:
                handle.write(
                    json.dumps(
                        row,
                        ensure_ascii=False,
                        sort_keys=True,
                        separators=(",", ":"),
                        default=json_default,
                    )
                    + "\n"
                )
        return destination

    def csv_writer(self, relative: str | Path, fieldnames: list[str]) -> csv.DictWriter:
        handle = self.open_text(relative, newline="")
        writer = csv.DictWriter(handle, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        return writer

    def close_files(self) -> None:
        while self._open_files:
            handle = self._open_files.pop()
            if not handle.closed:
                handle.flush()
                handle.close()

    def commit(self) -> Path:
        if self._committed:
            return self.final_path
        self.close_files()
        self.root.rename(self.final_path)
        self._committed = True
        self._temp_path = None
        return self.final_path

    def __exit__(self, exc_type: object, exc: object, traceback: object) -> None:
        self.close_files()
        if not self._committed and self._temp_path is not None:
            temp = self._temp_path.resolve()
            parent = self.final_path.parent.resolve()
            if temp.parent == parent and temp.name.startswith(f".{self.final_path.name}.tmp-"):
                shutil.rmtree(temp, ignore_errors=True)
