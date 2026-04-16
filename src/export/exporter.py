"""Export module: save history files and merge for PowerBI."""
import os
import re

import pandas as pd


def save_history(df: pd.DataFrame, sheet_name: str, history_dir: str) -> str:
    """
    Save consolidated DataFrame to history folder.
    File: FYXX_Rebate & Spending Shipment Report.xlsx
    Sheet: FYXX Data
    Returns saved file path.
    """
    os.makedirs(history_dir, exist_ok=True)
    fy_label = sheet_name.upper()
    file_name = f"{fy_label}_Rebate & Spending Shipment Report.xlsx"
    file_path = os.path.join(history_dir, file_name)
    sheet_tab = f"{fy_label} Data"
    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_tab, index=False)
    return file_path


def get_available_fy(history_dir: str) -> list[str]:
    """Return sorted list of FY labels available in history dir."""
    if not os.path.isdir(history_dir):
        return []
    fy_list = []
    for f in os.listdir(history_dir):
        m = re.match(r"^(FY\d+)_Rebate & Spending Shipment Report\.xlsx$", f, re.I)
        if m:
            fy_list.append(m.group(1).upper())
    return sorted(fy_list)


def merge_for_powerbi(fy_labels: list[str], history_dir: str, output_dir: str) -> str:
    """
    Merge selected FY files from history into InputDevice_Shipment.xlsx.
    Returns output path.
    """
    os.makedirs(output_dir, exist_ok=True)
    frames = []
    for fy in fy_labels:
        file_name = f"{fy.upper()}_Rebate & Spending Shipment Report.xlsx"
        file_path = os.path.join(history_dir, file_name)
        if not os.path.exists(file_path):
            continue
        sheet_tab = f"{fy.upper()} Data"
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_tab, engine="openpyxl")
            frames.append(df)
        except Exception:
            pass
    if not frames:
        return ""
    merged = pd.concat(frames, ignore_index=True)
    out_path = os.path.join(output_dir, "InputDevice_Shipment.xlsx")
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        merged.to_excel(writer, sheet_name="InputDevice_Shipment", index=False)
    return out_path
