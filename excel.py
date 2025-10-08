import re
from typing import Optional, Dict, Any
import pandas as pd

BANNED_TOKENS = {
    "MONITOR","DISPLAY","WINDOWS","SERVER","STANDARD","MICROSOFT","ASUS","LENOVO","HPE","HP",
    "INTEL","NVIDIA","RYZEN","GEFORCE","CORE","DDR","DDR4","DDR5","UHD","FHD","IPS","VA","TN",
    "DOS","ENG","ENGLISH","NEW","AIO","PCS","WORKSTATION","NOTEBOOK","LAPTOP","WUE","G5","G9",
    "BACKLIT","KEYBOARD","MOUSE","WIFI","BT","BLUETOOTH","PCIe","NVME","SSD","HDD","GB","TB",
    "INCH","WARRANTY","YRS","3YRS","1YW","AX","TNR","ID","FLEX","SLOT","HOT","PLUG","LOW",
    "HALOGEN","KIT","POWER","SUPPLY","TITANIUM","PRO","HOME","BUSINESS","STANDARD","LOQ"
}

def _clean_caps(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip().upper()

def _find_parenthetical_id(text: str) -> Optional[str]:
    s = _clean_caps(text)
    m = re.search(r"\(([^\)]+)\)", s)
    if not m:
        return None
    cand = re.sub(r"[^A-Z0-9#\-]", "", m.group(1))
    if len(re.sub(r"[^A-Z0-9]", "", cand)) < 5:
        return None
    return cand

def _find_hp_monitor_id(text: str) -> Optional[str]:
    s = _clean_caps(text)
    m = re.search(r"\b[0-9A-Z]{5,}(?:#[0-9A-Z]{2,})\b", s)
    return m.group(0) if m else None

def _find_dell_model_like(text: str) -> Optional[str]:
    s = _clean_caps(text)
    s = re.sub(r"[^A-Z0-9\-\s]", " ", s)
    tokens = re.findall(r"\b[A-Z0-9]{5,}\b", s)
    filtered = []
    for t in tokens:
        if t.isdigit():
            continue
        if t in BANNED_TOKENS:
            continue
        if re.match(r"^\d{3,}$", t):
            continue
        if re.match(r"^\d{1,2}YRS$", t):
            continue
        has_letter = bool(re.search(r"[A-Z]", t))
        has_digit = bool(re.search(r"\d", t))
        if not (has_letter and has_digit):
            continue
        filtered.append(t)
    if not filtered:
        return None
    preferred = [t for t in filtered if t[0] in ("P","U","S","E","A")]
    return preferred[-1] if preferred else filtered[-1]

def _phrase_before_first_comma(text: str) -> Optional[str]:
    if not text:
        return None
    head = text.split(",")[0].strip()
    return head or None

def _generic_from_description(text: str) -> Optional[str]:
    return (
        _find_parenthetical_id(text)
        or _find_hp_monitor_id(text)
        or _find_dell_model_like(text)
    )

def extract_product_id_for_row(sheet_name: str, row: Dict[str, Any]) -> Optional[str]:
    sname = _clean_caps(sheet_name)
    # Best-effort to find common columns or positional cells
    desc = None
    part = None
    ref = None

    # Try common label variants
    for k in row.keys():
        ku = _clean_caps(str(k))
        if ku in ("DESCRIPTION",):
            desc = row[k]
        elif ku in ("PART NUMBER","PART","PART NO","P/N"):
            part = row[k]
        elif ku in ("REF","REFERENCE"):
            ref = row[k]
        elif ku in ("UNIT PRICE /JAFZ","UNIT PRICE/JAFZ","UNIT PRICE","PRICE"):
            # cost handled outside; keep here if needed
            pass

    # Fallbacks for headerless rows: take first/second cells by index
    if desc is None:
        if 0 in row:
            desc = row[0]
        else:
            # Try the first key by order if it looks like description text
            try:
                first_key = list(row.keys())[0]
                desc = row[first_key]
            except Exception:
                desc = None
    if part is None and 1 in row:
        part = row[1]
    if ref is None and "REF" in row:
        ref = row["REF"]

    desc = str(desc or "").strip()
    part = str(part or "").strip()
    ref = str(ref or "").strip()

    # UPS: product id from ref
    if "UPS" in sname:
        return ref or None

    # Microsoft & ASUS: prefer description, else part number
    if "MICROSOFT" in sname or "ASUS" in sname:
        return _generic_from_description(desc) or (part if part and part.upper() != "NEW" else None)

    # Lenovo notebook & option: prefer description, else part number
    if "LENOVO" in sname and ("NOTEBOOK" in sname or "OPTION" in sname):
        return _generic_from_description(desc) or (part or None)

    # Lenovo PCs/AIO/Workstation/Monitor: prefer description, else part number
    if "LENOVO" in sname:
        return _generic_from_description(desc) or (part or None)

    # HP monitor: ID in description (often with '#')
    if "HP" in sname and "MONITOR" in sname:
        return _find_hp_monitor_id(desc) or _find_parenthetical_id(desc) or _find_dell_model_like(desc)

    # HP servers & parts: parentheses in description
    if "HP" in sname and ("SERVER" in sname or "PART" in sname):
        return _find_parenthetical_id(desc) or _generic_from_description(desc)

    # HP notebooks/workstation/option (no headers): take phrase before first comma
    if "HP" in sname and any(x in sname for x in ("NOTEBOOK","WORKSTATION","OPTION")):
        return _phrase_before_first_comma(desc) or _generic_from_description(desc)

    # HP PCs/AIO/Workstation (no headers), parentheses style like (9M9D7AT)
    if "HP" in sname and any(x in sname for x in ("PCS","AIO","WORKSTATION")):
        return _find_parenthetical_id(desc) or _generic_from_description(desc)

    # Dell monitors & accessories: from description
    if "DELL" in sname and (("MONITOR" in sname) or ("ACCESSOR" in sname)):
        return _generic_from_description(desc) or (part or None)

    # Consumer + AIO + Gaming: parentheses model
    if any(x in sname for x in ("CONSUMER","AIO","GAMING")):
        return _find_parenthetical_id(desc) or _generic_from_description(desc)

    # Default
    return _generic_from_description(desc) or (part or ref or None)
def run_on_excel_file(excel_file_path: str, output_file: str = "output_results.xlsx"):
    # Load all sheets
    xls = pd.read_excel(excel_file_path, sheet_name=None, dtype=str, engine="openpyxl")
    writer = pd.ExcelWriter(output_file, engine='openpyxl')

    for sheet_name, df in xls.items():
        print(f"Processing sheet: {sheet_name}")
        df.fillna("", inplace=True)

        # Apply the extraction function row by row
        def process_row(row: pd.Series) -> Optional[str]:
            return extract_product_id_for_row(sheet_name, row.to_dict())

        df["EXTRACTED_PRODUCT_ID"] = df.apply(process_row, axis=1)

        # Write to output Excel file
        df.to_excel(writer, sheet_name=sheet_name, index=False)

    writer._save()
    print(f"\n✅ Done! Results saved to {"final_merged.csv"}")

# === Entry Point ===

if __name__ == "__main__":
    # 🔁 Change this to the path of your Excel file
    excel_file_path = "Solid Solution October PRICE LIST 2025.xlsx"
    run_on_excel_file(excel_file_path)
