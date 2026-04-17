import os
import shutil
import subprocess
from typing import Any, Dict


def convert_with_libreoffice(
    input_path: str,
    target_format: str,
    output_path: str,
    **kwargs,
) -> Dict[str, Any]:
    """Convert documents using LibreOffice headless mode."""
    libreoffice_path = shutil.which("libreoffice") or shutil.which("soffice")
    if not libreoffice_path:
        raise RuntimeError(
            "LibreOffice not found. Install it first. Example (Ubuntu): sudo apt install libreoffice"
        )

    output_dir = os.path.dirname(output_path) or "."
    os.makedirs(output_dir, exist_ok=True)

    cmd = [
        libreoffice_path,
        "--headless",
        "--convert-to",
        target_format,
        "--outdir",
        output_dir,
        input_path,
    ]

    try:
        subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            check=True,
            timeout=kwargs.get("timeout", 300),
        )

        expected_generated_path = os.path.join(
            output_dir,
            f"{os.path.splitext(os.path.basename(input_path))[0]}.{target_format}",
        )

        if os.path.exists(expected_generated_path) and expected_generated_path != output_path:
            if os.path.exists(output_path):
                os.remove(output_path)
            os.rename(expected_generated_path, output_path)

        if not os.path.exists(output_path):
            raise RuntimeError(
                f"Conversion command succeeded but output file was not found: {output_path}"
            )

        return {"success": True, "output_path": output_path}

    except subprocess.CalledProcessError as e:
        stderr = (e.stderr or "").strip()
        stdout = (e.stdout or "").strip()
        raise RuntimeError(f"LibreOffice conversion failed. stderr={stderr} stdout={stdout}") from e


def get_file_type(file_path: str) -> str:
    """Return lower-case extension without leading dot."""
    if not os.path.exists(file_path):
        return ""
    return os.path.splitext(file_path)[1].lower().lstrip(".")


def is_office_file(file_path: str) -> bool:
    """Check if file extension is supported by this package."""
    supported_extensions = {"docx", "doc", "xlsx", "xls", "pptx", "ppt", "pdf"}
    ext = get_file_type(file_path)
    return ext in supported_extensions


def get_file_size(file_path: str, unit: str = "bytes") -> float:
    """Get file size in bytes/KB/MB/GB."""
    if not os.path.exists(file_path):
        return 0.0

    size_bytes = os.path.getsize(file_path)

    units = {
        "bytes": 1,
        "kb": 1024,
        "mb": 1024 * 1024,
        "gb": 1024 * 1024 * 1024,
    }

    factor = units.get(unit.lower(), 1)
    return round(size_bytes / factor, 2)


def ensure_dir(path: str) -> None:
    """Ensure directory exists."""
    os.makedirs(path, exist_ok=True)


def clean_temp_files(temp_dir: str, older_than_hours: int = 24) -> int:
    """Delete files older than threshold from temp directory."""
    import time

    now = time.time()
    deleted_count = 0

    if not os.path.exists(temp_dir):
        return 0

    for filename in os.listdir(temp_dir):
        file_path = os.path.join(temp_dir, filename)
        if os.path.isfile(file_path):
            if (now - os.path.getmtime(file_path)) > older_than_hours * 3600:
                os.remove(file_path)
                deleted_count += 1

    return deleted_count
