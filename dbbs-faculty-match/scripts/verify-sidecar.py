"""Reject sidecars containing stale Python modules, even after a successful freeze."""

import marshal
from pathlib import Path
import sys
from types import CodeType

from PyInstaller.archive.readers import CArchiveReader


def semantics(value):
    if isinstance(value, CodeType):
        return (
            value.co_code, semantics(value.co_consts), value.co_names,
            value.co_varnames, value.co_freevars, value.co_cellvars,
            value.co_flags, value.co_argcount, value.co_posonlyargcount,
            value.co_kwonlyargcount, value.co_exceptiontable,
        )
    if isinstance(value, tuple):
        return tuple(semantics(item) for item in value)
    return value


def verify(binary, python_directory):
    archive = CArchiveReader(str(binary))
    modules = archive.open_embedded_archive("PYZ.pyz")
    for source in sorted(python_directory.glob("*.py")):
        expected = compile(source.read_bytes(), source.name, "exec", optimize=0)
        actual = (
            marshal.loads(archive.extract(source.stem))
            if source.stem == "embedding_helper" else modules.extract(source.stem)
        )
        if semantics(expected) != semantics(actual):
            raise RuntimeError(f"The packaged helper contains stale code for {source.name}.")
    print("Verified packaged Python modules match the current source.")


if __name__ == "__main__":
    verify(Path(sys.argv[1]), Path(sys.argv[2]))
