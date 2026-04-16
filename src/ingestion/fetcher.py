"""Ingestion module: discover suppliers, fetch latest Excel files."""
import os
import re
from pathlib import Path


def get_suppliers(base_path: str) -> list[str]:
    """Return sorted list of supplier folder names under base_path."""
    if not os.path.isdir(base_path):
        return []
    return sorted(
        [d for d in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, d))]
    )


def _get_excel_files_in_folder(folder: str) -> list[Path]:
    """Return all .xlsx/.xls files in folder sorted by modification time (newest first)."""
    p = Path(folder)
    files = [f for f in p.iterdir() if f.is_file() and f.suffix.lower() in (".xlsx", ".xls") and not f.name.startswith("~$")]
    files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
    return files


def fetch_supplier_files(base_path: str, suppliers: list[str]) -> dict[str, list[Path]]:
    """
    For each supplier, find the "Spending and rebate" subfolder and return
    the list of Excel files sorted newest-first.
    Returns dict: {supplier: [Path, ...]}
    """
    result = {}
    for supplier in suppliers:
        supplier_dir = os.path.join(base_path, supplier)
        spending_dir = None
        for entry in os.scandir(supplier_dir):
            if entry.is_dir() and "spending and rebate" in entry.name.lower():
                spending_dir = entry.path
                break
        if spending_dir is None:
            result[supplier] = []
            continue
        result[supplier] = _get_excel_files_in_folder(spending_dir)
    return result
