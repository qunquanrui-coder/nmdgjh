"""Helpers for writing results without replacing existing user files."""

from contextlib import contextmanager
from pathlib import Path
import errno
import os
import shutil


def new_output_directory(path: Path) -> Path:
    """Create a separate directory for each export without touching previous ones."""
    path = Path(path)
    index = 0
    while True:
        candidate = path if index == 0 else path.with_name(f"{path.name} ({index})")
        try:
            candidate.mkdir()
            return candidate
        except FileExistsError:
            index += 1


@contextmanager
def new_output_file(path: Path):
    """Exclusively create a unique filename and remove incomplete output on error."""
    path = Path(path)
    index = 0
    while True:
        candidate = path if index == 0 else path.with_name(f"{path.stem} ({index}){path.suffix}")
        try:
            stream = candidate.open("xb")
            break
        except FileExistsError:
            index += 1
    try:
        with stream:
            yield stream, candidate
    except BaseException:
        try:
            candidate.unlink()
        except OSError:
            pass
        raise


def publish_output(tmp_path: Path, desired_path: Path) -> Path:
    """Publish a complete result exclusively, preserving existing files and the temp."""
    tmp_path, desired_path = Path(tmp_path), Path(desired_path)
    index = 0
    while True:
        candidate = desired_path if index == 0 else desired_path.with_name(
            f"{desired_path.stem} ({index}){desired_path.suffix}"
        )
        try:
            # A hard link publishes the fully written file atomically and fails
            # rather than replacing a file created by another process.
            os.link(tmp_path, candidate)
            return candidate
        except FileExistsError:
            index += 1
        except OSError as exc:
            # Filesystems without hard-link support use exclusive copying.
            if exc.errno not in (errno.EPERM, errno.EACCES, errno.EXDEV, errno.ENOTSUP, errno.EOPNOTSUPP, errno.EINVAL):
                raise
            break
    with new_output_file(desired_path) as (output, actual_path):
        with tmp_path.open("rb") as source:
            shutil.copyfileobj(source, output)
    return actual_path
