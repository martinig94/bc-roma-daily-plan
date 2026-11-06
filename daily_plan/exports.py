"""Save generated files to disk and bundle images."""

import zipfile
from io import BytesIO
from pathlib import Path


def zip_images(images: dict[str, bytes]) -> bytes:
    """Bundle PNG images into a zip archive.

    Args:
        images: Mapping of image name (no extension) to PNG bytes.

    Returns:
        Bytes of the zip file.
    """
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in images.items():
            archive.writestr(f"{name}.png", data)
    return buffer.getvalue()


def save_outputs(
    directory: Path, workbook: bytes, images: dict[str, bytes], workbook_stem: str
) -> list[Path]:
    """Write the workbook and PNG images to a directory.

    Args:
        directory: Target directory, created if missing.
        workbook: Bytes of the xlsx file.
        images: Mapping of image name (no extension) to PNG bytes.
        workbook_stem: Workbook file name without extension.

    Returns:
        Paths of all written files.
    """
    directory.mkdir(parents=True, exist_ok=True)
    files = {f"{workbook_stem}.xlsx": workbook}
    files.update({f"{name}.png": data for name, data in images.items()})
    written = []
    for filename, data in files.items():
        path = directory / filename
        path.write_bytes(data)
        written.append(path)
    return written
