"""Save generated files to disk and bundle images."""

import zipfile
from io import BytesIO
from pathlib import Path


def zip_images(images: dict[str, bytes]) -> bytes:
    """Bundle PNG images into a zip archive.

    Args:
        images: Mapping of section name to PNG bytes.

    Returns:
        Bytes of the zip file.
    """
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in images.items():
            archive.writestr(f"{name}.png", data)
    return buffer.getvalue()


def save_outputs(
    directory: Path, workbook: bytes, images: dict[str, bytes], stem: str
) -> list[Path]:
    """Write the workbook and PNG images to a directory.

    Args:
        directory: Target directory, created if missing.
        workbook: Bytes of the xlsx file.
        images: Mapping of section name to PNG bytes.
        stem: File name prefix, without extension.

    Returns:
        Paths of all written files.
    """
    directory.mkdir(parents=True, exist_ok=True)
    written = [directory / f"{stem}.xlsx"]
    written[0].write_bytes(workbook)
    for name, data in images.items():
        path = directory / f"{stem}_{name}.png"
        path.write_bytes(data)
        written.append(path)
    return written
