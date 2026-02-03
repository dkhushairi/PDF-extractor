import re
from io import BytesIO
from typing import Dict, List, Optional, Tuple

import pdfplumber
import streamlit as st


# --- Template headers we rely on (from your sample PDF format) ---
HEADERS = [
    "Stellvertreter",
    "Anschlussnehmer",
    "Anschlussort",
    "Angaben zur Kundenanlage",
    "Angaben zu den PV-Modulen",
    "Angaben zur Speichereinheit",
]

# Labels inside sections (we extract value AFTER ":")
TARGETS = {
    "Stellvertreter": [
        "Firma",
        "Erreichbarkeit E-Mail",
    ],
    "Anschlussnehmer": [
        "Anschrift",
        "Erreichbarkeit Telefon",
        # name is special (not after ":"), handled separately
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
    """Extract text from all pages."""
    out_lines: List[str] = []
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            txt = page.extract_text(layout=True) or ""
            out_lines.append(txt)
    # Normalize newlines
    text = "\n".join(out_lines)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    return text


def split_lines(text: str) -> List[str]:
    """Split and lightly normalize lines."""
    lines = [ln.strip() for ln in text.split("\n")]
    # Keep empty lines (sometimes helps), but we often skip them later.
    return lines


def find_header_indices(lines: List[str], headers: List[str]) -> Dict[str, int]:
    """Return the first occurrence index of each header in the lines list."""
    idx_map = {}
    for i, ln in enumerate(lines):
        if ln in headers and ln not in idx_map:
            idx_map[ln] = i
    return idx_map


def get_section_slice(lines: List[str], header: str, header_indices: Dict[str, int]) -> Tuple[int, int]:
    """Return [start, end) line indices for a section."""
    if header not in header_indices:
        return (-1, -1)
    start = header_indices[header] + 1

    # end is the next header occurrence after this header
    starts_sorted = sorted((h, idx) for h, idx in header_indices.items())
    current_idx = header_indices[header]
    end = len(lines)
    for h, idx in starts_sorted:
        if idx > current_idx:
            end = idx
            break
    return start, end


def extract_value_after_colon(section_lines: List[str], label: str) -> Optional[str]:
    """
    Extracts the value after 'label:' in the section.
    Works even if there are extra spaces.
    """
    # Example: "Firma: XY"
    pattern = rf"^{re.escape(label)}\s*:\s*(.+?)\s*$"
    for ln in section_lines:
        m = re.match(pattern, ln)
        if m:
            return m.group(1).strip()
    return None


def extract_first_person_name(section_lines: List[str]) -> Optional[str]:
    """
    In your sample, under 'Anschlussnehmer' the name is a standalone line like:
    'XY' (not 'Name: ...').
    We'll return the first non-empty line that does NOT contain ':' and is not a sub-heading.
    """
    for ln in section_lines:
        if not ln:
            continue
        if ":" in ln:
            continue
        # Skip obvious non-name lines
        if ln in HEADERS:
            continue
        if ln.lower().startswith("geburtsdatum"):
            continue
        # Heuristic: common German honorifics or general "person-like" line
        if re.match(r"^(Herr|Frau|Firma)\b", ln):
            return ln.strip()
        # If honorific not present, still accept first plain line (fallback)
        return ln.strip()
    return None


def extract_requested_fields(pdf_bytes: bytes) -> Dict[str, List[str]]:
    """
    Returns dict: section -> list of 'Field: Value' lines (including field title).
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

        # Standard label: value after colon
        for label in TARGETS.get(header, []):
            val = extract_value_after_colon(section_lines, label)
            if val is not None:
                section_out.append(f"{label}: {val}")

        if section_out:
            results[header] = section_out

    return results


def format_as_txt(results: Dict[str, List[str]]) -> str:
    """
    Build a clean TXT with section titles + extracted lines.
    """
    parts: List[str] = []
    for header in HEADERS:
        if header not in results:
            continue
        parts.append(header)
        for line in results[header]:
            parts.append(f"- {line}")
        parts.append("")  # blank line between sections
    return "\n".join(parts).strip() + "\n"


# ------------------ Streamlit UI ------------------

st.set_page_config(page_title="PDF Daten-Extractor", layout="centered")
st.title("PDF Daten-Extractor (Template-basiert)")
st.write("Upload dein PDF, dann bekommst du die gewünschten Felder als TXT zum Kopieren/Download.")

uploaded = st.file_uploader("PDF hochladen", type=["pdf"])

if uploaded:
    pdf_bytes = uploaded.read()

    with st.spinner("Extrahiere Daten..."):
        results = extract_requested_fields(pdf_bytes)
        txt = format_as_txt(results)

    st.subheader("Ergebnis (zum Kopieren)")
    st.text_area("TXT Output", value=txt, height=320)

    st.download_button(
        label="TXT herunterladen",
        data=txt.encode("utf-8"),
        file_name=f"{uploaded.name.rsplit('.', 1)[0]}_extracted.txt",
        mime="text/plain",
    )

    with st.expander("Debug: Gefundene Sections/Keys anzeigen"):
        st.write(results)

else:
    st.info("Bitte ein PDF hochladen.")
