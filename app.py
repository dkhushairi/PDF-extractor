import re
from io import BytesIO
from typing import Dict, List, Optional, Tuple
from datetime import datetime
from zoneinfo import ZoneInfo

import pdfplumber
import streamlit as st
import fitz  # PyMuPDF


# -----------------------------
# 1) EXTRACTION CONFIG
# -----------------------------
HEADERS = [
    "Stellvertreter",
    "Anschlussnehmer",
    "Anschlussort",
    "Angaben zur Kundenanlage",
    "Angaben zu den PV-Modulen",
    "Angaben zur Speichereinheit",
]

TARGETS = {
    "Stellvertreter": [
        "Firma",
        "Erreichbarkeit E-Mail",
    ],
    "Anschlussnehmer": [
        "Anschrift",
        "Erreichbarkeit Telefon",
    ],
    "Anschlussort": [
        "Straße",
    ],
    "Angaben zur Kundenanlage": [
        "Mess- und Betriebskonzept",
    ],
    "Angaben zu den PV-Modulen": [
        "Gesamtleistung aller PV-Module in kWp",
    ],
    "Angaben zur Speichereinheit": [
        "Bruttokapazität des Speichereinheit",
    ],
}


def extract_full_text(pdf_bytes: bytes) -> str:
    """Extract text from all pages of a PDF."""
    out_lines: List[str] = []
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            txt = page.extract_text(layout=True) or ""
            out_lines.append(txt)
    text = "\n".join(out_lines)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    return text


def split_lines(text: str) -> List[str]:
    return [ln.strip() for ln in text.split("\n")]


def find_header_indices(lines: List[str], headers: List[str]) -> Dict[str, int]:
    idx_map = {}
    for i, ln in enumerate(lines):
        if ln in headers and ln not in idx_map:
            idx_map[ln] = i
    return idx_map


def get_section_slice(lines: List[str], header: str, header_indices: Dict[str, int]) -> Tuple[int, int]:
    if header not in header_indices:
        return (-1, -1)
    start = header_indices[header] + 1
    current_idx = header_indices[header]
    end = len(lines)

    starts_sorted = sorted((h, idx) for h, idx in header_indices.items())
    for _, idx in starts_sorted:
        if idx > current_idx:
            end = idx
            break
    return start, end


def extract_value_after_colon(section_lines: List[str], label: str) -> Optional[str]:
    pattern = rf"^{re.escape(label)}\s*:\s*(.+?)\s*$"
    for ln in section_lines:
        m = re.match(pattern, ln, flags=re.IGNORECASE)
        if m:
            return m.group(1).strip()
    return None


def extract_first_person_name(section_lines: List[str]) -> Optional[str]:
    """
    Under 'Anschlussnehmer' the name is typically a standalone line like:
    'Herr XY' (not 'Name: ...').
    """
    for ln in section_lines:
        if not ln:
            continue
        if ":" in ln:
            continue
        if ln in HEADERS:
            continue
        # common honorifics
        if re.match(r"^(Herr|Frau)\b", ln):
            return ln.strip()
        # fallback: first plain non-empty line
        return ln.strip()
    return None


def extract_requested_fields(pdf_bytes: bytes) -> Dict[str, List[str]]:
    """
    Returns: { section_header: ["Field: Value", ...] }
    """
    text = extract_full_text(pdf_bytes)
    lines = split_lines(text)
    header_indices = find_header_indices(lines, HEADERS)

    results: Dict[str, List[str]] = {}

    for header in HEADERS:
        start, end = get_section_slice(lines, header, header_indices)
        if start == -1:
            continue

        section_lines = lines[start:end]
        section_out: List[str] = []

        # Special: Anschlussnehmer name line
        if header == "Anschlussnehmer":
            name = extract_first_person_name(section_lines)
            if name:
                section_out.append(f"Name: {name}")

        # Standard labels
        for label in TARGETS.get(header, []):
            val = extract_value_after_colon(section_lines, label)
            if val is not None:
                section_out.append(f"{label}: {val}")

        if section_out:
            results[header] = section_out

    return results


def format_as_txt(results: Dict[str, List[str]]) -> str:
    parts: List[str] = []
    for header in HEADERS:
        if header not in results:
            continue
        parts.append(header)
        for line in results[header]:
            parts.append(f"- {line}")
        parts.append("")
    return "\n".join(parts).strip() + "\n"


def _get_value(results_by_section: Dict[str, List[str]], section: str, label: str) -> str:
    for ln in results_by_section.get(section, []):
        ln = ln.lstrip("- ").strip()
        if ln.lower().startswith(label.lower() + ":"):
            return ln.split(":", 1)[1].strip()
    return ""


def build_extracted_dict(results_by_section: Dict[str, List[str]]) -> Dict[str, str]:
    name_line = ""
    for ln in results_by_section.get("Anschlussnehmer", []):
        ln = ln.lstrip("- ").strip()
        if ln.lower().startswith("name:"):
            name_line = ln.split(":", 1)[1].strip()
            break

    return {
        "name": name_line,
        "anschrift": _get_value(results_by_section, "Anschlussnehmer", "Anschrift"),
        "telefon": _get_value(results_by_section, "Anschlussnehmer", "Erreichbarkeit Telefon"),
        "anschlussort_strasse": _get_value(results_by_section, "Anschlussort", "Straße"),
        "messkonzept": _get_value(results_by_section, "Angaben zur Kundenanlage", "Mess- und Betriebskonzept"),
        "pv_kwp": _get_value(results_by_section, "Angaben zu den PV-Modulen", "Gesamtleistung aller PV-Module in kWp"),
        "speicher_kwh": _get_value(results_by_section, "Angaben zur Speichereinheit", "Bruttokapazität des Speichereinheit"),
    }


# -----------------------------
# 2) FILL FILLABLE PDF (ACROFORM)
# -----------------------------
def _set_text_field(doc: fitz.Document, field_name: str, value: str) -> None:
    for page in doc:
        widgets = page.widgets() or []
        for w in widgets:
            if w.field_name == field_name:
                w.field_value = value or ""
                w.update()
                return


def _set_checkbox(doc: fitz.Document, field_name: str, checked: bool) -> None:
    for page in doc:
        widgets = page.widgets() or []
        for w in widgets:
            if w.field_name == field_name:
                # Use actual "on" state from PDF (could be "Ja", "On", etc.)
                w.field_value = w.on_state() if checked else "Off"
                w.update()
                return


def fill_template_pdf(template_pdf_bytes: bytes, extracted: Dict[str, str]) -> bytes:
    """
    Field mapping based on YOUR fillable PDF field names:
      - Name Vorname
      - Straße  Nr  Ort
      - Telefonnummer
      - 2 Standort der Photovoltaikanlage
      - Check 1, Check 5, Check 7, Check 9
      - Bemerkung 5, Bemerkung 9
      - kWp, kWh
      - ja_2
    """
    doc = fitz.open(stream=template_pdf_bytes, filetype="pdf")

    # 1) Angaben zum Anlagenbetreiber
    _set_text_field(doc, "Name Vorname", extracted.get("name", ""))
    _set_text_field(doc, "Straße  Nr  Ort", extracted.get("anschrift", ""))
    _set_text_field(doc, "Telefonnummer", extracted.get("telefon", ""))

    # 2) Standort der Photovoltaikanlage
    _set_text_field(doc, "2 Standort der Photovoltaikanlage", extracted.get("anschlussort_strasse", ""))

    # Unterlagen checkboxes
    _set_checkbox(doc, "Check 1", True)
    _set_checkbox(doc, "Check 5", True)
    _set_checkbox(doc, "Check 7", True)
    _set_checkbox(doc, "Check 9", True)

    # Bemerkung 5 = current date DD.MM.YYYY (Berlin time)
    today = datetime.now(ZoneInfo("Europe/Berlin")).strftime("%d.%m.%Y")
    _set_text_field(doc, "Bemerkung 5", today)

    # Bemerkung 9 = Mess- und Betriebskonzept
    _set_text_field(doc, "Bemerkung 9", extracted.get("messkonzept", ""))

    # 3) Technische Daten
    _set_text_field(doc, "kWp", extracted.get("pv_kwp", ""))

    speicher_kwh = (extracted.get("speicher_kwh") or "").strip()
    _set_text_field(doc, "kWh", speicher_kwh)
    _set_checkbox(doc, "ja_2", bool(speicher_kwh))

    out = doc.tobytes(deflate=True)
    doc.close()
    return out


# -----------------------------
# 3) STREAMLIT UI
# -----------------------------
st.set_page_config(page_title="PDF Extract + Fill", layout="centered")
st.title("PDF Daten extrahieren + Fillable PDF automatisch ausfüllen")

st.write(
    "1) **Source PDF** hochladen (mit Stellvertreter/Anschlussnehmer/...)\n"
    "2) **Fillable Template PDF** hochladen (AcroForm)\n"
    "→ Dann bekommst du **TXT** + **gefülltes PDF** zum Download."
)

source_pdf = st.file_uploader("1) Source PDF hochladen", type=["pdf"])
template_pdf = st.file_uploader("2) Fillable Template PDF hochladen", type=["pdf"])

if source_pdf:
    source_bytes = source_pdf.read()

    with st.spinner("Extrahiere Daten aus Source PDF..."):
        results_by_section = extract_requested_fields(source_bytes)
        txt_out = format_as_txt(results_by_section)

    st.subheader("Extrahierte Daten (TXT)")
    st.text_area("TXT Output", value=txt_out, height=320)

    st.download_button(
        "TXT herunterladen",
        data=txt_out.encode("utf-8"),
        file_name=f"{source_pdf.name.rsplit('.', 1)[0]}_extracted.txt",
        mime="text/plain",
    )

    st.subheader("Mapping (was ins Template geschrieben wird)")
    extracted_dict = build_extracted_dict(results_by_section)
    st.json(extracted_dict)

    if template_pdf:
        template_bytes = template_pdf.read()

        with st.spinner("Fülle Template PDF..."):
            filled_pdf_bytes = fill_template_pdf(template_bytes, extracted_dict)

        st.download_button(
            "Gefülltes PDF herunterladen",
            data=filled_pdf_bytes,
            file_name="deckblatt_filled.pdf",
            mime="application/pdf",
        )
    else:
        st.info("Jetzt noch die **Fillable Template PDF** hochladen, dann kann ich das Deckblatt automatisch füllen.")
else:
    st.info("Bitte zuerst die **Source PDF** hochladen.")
