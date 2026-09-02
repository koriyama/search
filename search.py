import streamlit as st
import pandas as pd
import requests
import time
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
import re

# ============================================================
# 1. YOUR ORIGINAL FUNCTIONS (unchanged except export adapted)
# ============================================================

def reconstruct_abstract(inverted_index):
    if not inverted_index:
        return ""
    positions = {}
    for word, idxs in inverted_index.items():
        for i in idxs:
            positions[i] = word
    if not positions:
        return ""
    max_pos = max(positions.keys())
    return " ".join(positions.get(i, "") for i in range(max_pos + 1))

def extract_authors(work):
    names = []
    for a in work.get("authorships", []):
        author = a.get("author", {}) or {}
        name = author.get("display_name")
        if name:
            names.append(name)
    return "; ".join(names)

def extract_journal(work):
    primary = work.get("primary_location") or {}
    source = primary.get("source") or {}
    return source.get("display_name", "") or ""

def extract_pdf_url(work):
    for loc_key in ("best_oa_location", "primary_location"):
        loc = work.get(loc_key) or {}
        pdf = loc.get("pdf_url")
        if pdf:
            return pdf
    open_access = work.get("open_access") or {}
    return open_access.get("oa_url", "") or ""

def normalize_work(work):
    doi = work.get("doi", "") or ""
    if doi.startswith("https://doi.org/"):
        doi_url = doi
    elif doi:
        doi_url = f"https://doi.org/{doi}"
    else:
        doi_url = ""
    return {
        "id": work.get("id", ""),
        "title": work.get("display_name", "") or work.get("title", "") or "",
        "authors": extract_authors(work),
        "journal": extract_journal(work),
        "abstract": reconstruct_abstract(work.get("abstract_inverted_index")),
        "doi_url": doi_url,
        "pdf_url": extract_pdf_url(work),
        "year": work.get("publication_year"),
    }

OPENALEX_URL = "https://api.openalex.org/works"

def search_openalex_phrase_year(phrase, year, email=None, per_page=200, session=None, sleep_between=0.1):
    session = session or requests.Session()
    results = []
    cursor = "*"
    params_base = {
        "filter": f'title_and_abstract.search:"{phrase}",publication_year:{year}',
        "per-page": per_page,
    }
    if email:
        params_base["mailto"] = email

    while cursor:
        params = dict(params_base)
        params["cursor"] = cursor
        resp = session.get(OPENALEX_URL, params=params, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        for work in data.get("results", []):
            results.append(normalize_work(work))
        cursor = (data.get("meta", {}) or {}).get("next_cursor")
        if not data.get("results"):
            break
        time.sleep(sleep_between)
    return results

def collect(phrases, years, email=None, sleep_between=0.1, fetch_fn=search_openalex_phrase_year):
    by_year = {y: {} for y in years}
    for year in years:
        for phrase in phrases:
            # We'll use st.progress if called from Streamlit, but this function can also run standalone
            rows = fetch_fn(phrase, year, email=email, sleep_between=sleep_between)
            for row in rows:
                key = row["id"] or row["doi_url"] or f"{row['title'].strip().lower()}|{row['year']}"
                if key in by_year[year]:
                    existing = by_year[year][key]
                    if phrase not in existing["matched_phrases"]:
                        existing["matched_phrases"].append(phrase)
                else:
                    row = dict(row)
                    row["matched_phrases"] = [phrase]
                    by_year[year][key] = row
    return {year: list(rows.values()) for year, rows in by_year.items()}

# ============================================================
# 2. EXPORT ADAPTED TO BYTESIO (instead of writing to disk)
# ============================================================

ILLEGAL_CHARACTERS_RE = re.compile(r'[\000-\010]|[\013-\014]|[\016-\037]')

def sanitize_for_excel(value):
    if isinstance(value, str):
        return ILLEGAL_CHARACTERS_RE.sub('', value)
    return value

def export_to_excel_bytes(results_by_year, phrases, years, out_bytes, email=None):
    wb = Workbook()
    wb.remove(wb.active)

    header = ["Title", "Authors", "Journal", "Abstract", "DOI", "PDF URL", "Matched Phrase(s)"]
    bold = Font(bold=True)
    missing_pdfs_rows = []

    for year in years:
        ws = wb.create_sheet(title=str(year))
        for col_idx, h in enumerate(header, start=1):
            ws.cell(row=1, column=col_idx, value=h).font = bold

        rows = sorted(results_by_year.get(year, []), key=lambda r: r["title"].lower())
        for r_idx, row in enumerate(rows, start=2):
            ws.cell(row=r_idx, column=1, value=sanitize_for_excel(row["title"]))
            ws.cell(row=r_idx, column=2, value=sanitize_for_excel(row["authors"]))
            ws.cell(row=r_idx, column=3, value=sanitize_for_excel(row["journal"]))
            ws.cell(row=r_idx, column=4, value=sanitize_for_excel(row["abstract"]))

            doi_cell = ws.cell(row=r_idx, column=5)
            if row["doi_url"]:
                doi_cell.value = row["doi_url"]
                doi_cell.hyperlink = row["doi_url"]
                doi_cell.font = Font(color="0563C1", underline="single")

            pdf_cell = ws.cell(row=r_idx, column=6)
            if row["pdf_url"]:
                pdf_cell.value = row["pdf_url"]
                pdf_cell.hyperlink = row["pdf_url"]
                pdf_cell.font = Font(color="0563C1", underline="single")
            else:
                missing_pdfs_rows.append(row)

            ws.cell(row=r_idx, column=7, value="; ".join(row.get("matched_phrases", [])))

        widths = [50, 30, 30, 70, 35, 45, 25]
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"

    # Settings sheet
    settings_ws = wb.create_sheet(title="Search Settings")
    settings_ws.cell(row=1, column=1, value="Setting").font = bold
    settings_ws.cell(row=1, column=2, value="Value").font = bold
    settings_ws.cell(row=2, column=1, value="Exact phrases searched")
    settings_ws.cell(row=2, column=2, value="; ".join(phrases))
    settings_ws.cell(row=3, column=1, value="Publication years")
    settings_ws.cell(row=3, column=2, value=", ".join(str(y) for y in years))
    settings_ws.cell(row=4, column=1, value="OpenAlex polite-pool email")
    settings_ws.cell(row=4, column=2, value=email or "")
    settings_ws.cell(row=5, column=1, value="Total unique records")
    settings_ws.cell(row=5, column=2, value=sum(len(v) for v in results_by_year.values()))
    settings_ws.column_dimensions["A"].width = 30
    settings_ws.column_dimensions["B"].width = 60

    # Missing PDFs sheet
    missing_ws = wb.create_sheet(title="Missing PDFs")
    for col_idx, h in enumerate(["Title", "Authors", "Year", "DOI"], start=1):
        missing_ws.cell(row=1, column=col_idx, value=h).font = bold
    for r_idx, row in enumerate(sorted(missing_pdfs_rows, key=lambda r: (r["year"] or 0, r["title"].lower())), start=2):
        missing_ws.cell(row=r_idx, column=1, value=sanitize_for_excel(row["title"]))
        missing_ws.cell(row=r_idx, column=2, value=sanitize_for_excel(row["authors"]))
        missing_ws.cell(row=r_idx, column=3, value=row["year"])
        doi_cell = missing_ws.cell(row=r_idx, column=4, value=row["doi_url"])
        if row["doi_url"]:
            doi_cell.hyperlink = row["doi_url"]
            doi_cell.font = Font(color="0563C1", underline="single")
    for i, w in enumerate([50, 30, 10, 35], start=1):
        missing_ws.column_dimensions[get_column_letter(i)].width = w

    wb.save(out_bytes)

# ============================================================
# 3. STREAMLIT USER INTERFACE
# ============================================================

st.set_page_config(page_title="OpenAlex Literature Collector", layout="wide")
st.title("📚 OpenAlex Literature Collector (generic)")

st.markdown("Enter any search phrases (one per line) to find scholarly works on OpenAlex. "
            "The tool fetches title, authors, journal, abstract, DOI links, and PDF URLs when available.")

with st.form("search_form"):
    col1, col2 = st.columns(2)
    with col1:
        phrases_input = st.text_area(
            "🔍 Search Phrases (one per line)",
            value="epistemic cognition\npersonal epistemology\nepistemological beliefs",  # example – users can replace
            help="Each phrase is searched as an exact match in title and abstract."
        )
        start_year = st.number_input("Start Year", min_value=1900, max_value=2030, value=2000, step=1)
        end_year = st.number_input("End Year", min_value=1900, max_value=2030, value=2026, step=1)
    with col2:
        email = st.text_input(
            "📧 Your Email (for OpenAlex polite pool)",
            value="your_email@example.com",
            help="Providing an email gets you faster API access. You can leave it as is, but real email is better."
        )
        st.caption("OpenAlex gives higher priority to requests that include a contact email.")

    submitted = st.form_submit_button("🚀 Run Search", use_container_width=True)

# Cache results to avoid re‑running on every UI interaction
@st.cache_data(show_spinner=False)
def run_collection(phrases_tuple, years_tuple, email):
    return collect(list(phrases_tuple), list(years_tuple), email=email)

if submitted:
    phrases = [p.strip() for p in phrases_input.splitlines() if p.strip()]
    years = list(range(int(start_year), int(end_year) + 1))

    if not phrases:
        st.error("Please enter at least one search phrase.")
    else:
        with st.status("⏳ Searching OpenAlex...", expanded=True) as status:
            results_by_year = run_collection(tuple(phrases), tuple(years), email)
            total = sum(len(v) for v in results_by_year.values())
            status.update(label=f"✅ Done! Found {total} unique records.", state="complete")

        # Preview
        st.subheader("📄 Preview of Results")
        all_rows = []
        for year, rows in results_by_year.items():
            for r in rows:
                all_rows.append({
                    "Year": year,
                    "Title": r["title"],
                    "Authors": r["authors"],
                    "Journal": r["journal"],
                    "Has PDF": "✅" if r["pdf_url"] else "❌",
                    "Phrases": "; ".join(r.get("matched_phrases", []))
                })
        df = pd.DataFrame(all_rows)
        st.dataframe(df, use_container_width=True, height=400)

        missing_count = sum(1 for r in all_rows if r["Has PDF"] == "❌")
        st.caption(f"📌 {missing_count} records are missing a PDF link (they'll appear in the 'Missing PDFs' sheet).")

        # Export button
        if st.button("⬇️ Download Excel File", use_container_width=True):
            with st.spinner("Generating Excel..."):
                output = BytesIO()
                export_to_excel_bytes(results_by_year, phrases, years, output, email=email)
                output.seek(0)
                st.download_button(
                    label="📥 Click here to save the file",
                    data=output,
                    file_name=f"OpenAlex_Search_{start_year}-{end_year}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )