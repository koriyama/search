import streamlit as st
import pandas as pd
import requests
import time
import urllib.parse
from io import BytesIO, StringIO
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
import re
import random
from datetime import datetime

# ============================================================
# 1. CORE FUNCTIONS – unchanged (reconstruct_abstract, extract_authors, etc.)
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

def normalize_work(work, source="OpenAlex"):
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
        "source": source,
    }

def normalize_eric(record):
    """Map ERIC record to common schema."""
    pubdate = record.get("publicationdate", "")
    year = None
    if pubdate:
        try:
            year = int(pubdate[:4])
        except ValueError:
            pass
    return {
        "id": record.get("id", ""),
        "title": record.get("title", ""),
        "authors": record.get("author", ""),
        "journal": record.get("source", ""),
        "abstract": record.get("abstract", ""),
        "doi_url": record.get("doi", "") or "",
        "pdf_url": record.get("fulltext", "") or "",
        "year": year,
        "source": "ERIC",
    }

# ============================================================
# 2. QUERY TRANSLATION
# ============================================================

def translate_for_openalex(query):
    """Map common field names to OpenAlex search fields."""
    # OpenAlex supports: title, author, primary_location.source.display_name, abstract
    pattern = re.compile(r'\bjournal\s*:', re.IGNORECASE)
    return pattern.sub("primary_location.source.display_name:", query)

def translate_for_eric(query):
    """Map common field names to ERIC field codes."""
    mapping = {
        r'\btitle\s*:': "ti:",
        r'\bauthor\s*:': "au:",
        r'\bjournal\s*:': "so:",
        r'\babstract\s*:': "ab:",
    }
    for pat, repl in mapping.items():
        query = re.sub(pat, repl, query, flags=re.IGNORECASE)
    return query

# ============================================================
# 3. SEARCH FUNCTIONS – OpenAlex
# ============================================================

OPENALEX_URL = "https://api.openalex.org/works"

def search_openalex_query_year(query, year, email=None, api_key=None, per_page=200,
                               session=None, sleep_between=0.1, work_types=None,
                               just_count=False):
    session = session or requests.Session()
    cursor = "*"
    count = 0

    filter_parts = [f'publication_year:{year}']
    if work_types and "All types" not in work_types:
        type_filter = "|".join(work_types)
        filter_parts.append(f'type:{type_filter}')
    filter_string = ",".join(filter_parts)

    translated = translate_for_openalex(query)

    params = {
        "search": translated,
        "filter": filter_string,
        "per-page": 1 if just_count else per_page,
    }
    if email:
        params["mailto"] = email
    if api_key:
        params["api_key"] = api_key

    temp_params = dict(params)
    temp_params["cursor"] = "*"
    debug_url = f"{OPENALEX_URL}?{urllib.parse.urlencode(temp_params)}"

    if just_count:
        try:
            resp = session.get(OPENALEX_URL, params=params, timeout=30)
            resp.raise_for_status()
            data = resp.json()
            count = data.get("meta", {}).get("count", 0)
        except Exception as e:
            st.error(f"OpenAlex count error for '{query}' year {year}: {e}")
            raise
        return count, debug_url
    else:
        results = []
        first_page_meta_count = None
        while cursor:
            params["cursor"] = cursor
            try:
                resp = session.get(OPENALEX_URL, params=params, timeout=30)
                resp.raise_for_status()
            except Exception as e:
                st.error(f"OpenAlex fetch error for '{query}' year {year}: {e}")
                raise
            data = resp.json()
            if first_page_meta_count is None:
                first_page_meta_count = data.get("meta", {}).get("count", 0)
            for work in data.get("results", []):
                results.append(normalize_work(work, source="OpenAlex"))
            cursor = data.get("meta", {}).get("next_cursor")
            if not data.get("results"):
                break
            time.sleep(sleep_between)
        return results, first_page_meta_count, debug_url

# ============================================================
# 4. SEARCH FUNCTIONS – ERIC
# ============================================================

ERIC_URL = "https://api.ies.ed.gov/eric/"

def search_eric_query_year(query, year, per_page=200, session=None, sleep_between=0.1,
                           just_count=False):
    session = session or requests.Session()
    translated = translate_for_eric(query)

    start = 0
    results = []
    total_found = 0
    debug_url = ""

    # For count, just fetch one record to get numFound
    if just_count:
        params = {
            "query": translated,
            "rows": 0,
            "start": 0,
            "format": "json",
            "publicationdatestart": f"{year}-01-01",
            "publicationdateend": f"{year}-12-31",
        }
        resp = session.get(ERIC_URL, params=params, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        count = data.get("response", {}).get("numFound", 0)
        return count, f"{ERIC_URL}?{urllib.parse.urlencode(params)}"

    while True:
        params = {
            "query": translated,
            "rows": per_page,
            "start": start,
            "format": "json",
            "publicationdatestart": f"{year}-01-01",
            "publicationdateend": f"{year}-12-31",
        }
        if start == 0:
            debug_url = f"{ERIC_URL}?{urllib.parse.urlencode(params)}"
        try:
            resp = session.get(ERIC_URL, params=params, timeout=30)
            resp.raise_for_status()
        except Exception as e:
            st.error(f"ERIC fetch error for '{query}' year {year}: {e}")
            raise
        data = resp.json()
        response = data.get("response", {})
        total_found = response.get("numFound", 0)
        docs = response.get("docs", [])
        if not docs:
            break
        for doc in docs:
            results.append(normalize_eric(doc))
        start += per_page
        if start >= total_found:
            break
        time.sleep(sleep_between)
    return results, total_found, debug_url

# ============================================================
# 5. COLLECT FUNCTION (supports multiple sources)
# ============================================================

def collect(queries, years, fetch_fn, source_label, email=None, api_key=None,
            sleep_between=0.1, work_types=None):
    """
    Fetch results for each (query, year) using fetch_fn.
    Returns dict year -> list of rows (with source_label added).
    """
    by_year = {y: {} for y in years}
    debug_info = {}
    for year in years:
        for query in queries:
            rows, count, debug_url = fetch_fn(
                query, year, email=email, api_key=api_key,
                sleep_between=sleep_between, work_types=work_types,
                just_count=False
            )
            debug_info[(query, year)] = {"count": count, "url": debug_url}
            for row in rows:
                # Ensure source is set
                row["source"] = source_label
                key = row.get("id") or row.get("doi_url") or f"{row['title'].strip().lower()}|{row['year']}"
                if key in by_year[year]:
                    existing = by_year[year][key]
                    if query not in existing["matched_phrases"]:
                        existing["matched_phrases"].append(query)
                    # If sources differ, combine
                    if row["source"] not in existing.get("sources", []):
                        existing.setdefault("sources", []).append(row["source"])
                else:
                    row["matched_phrases"] = [query]
                    row["sources"] = [row["source"]]
                    by_year[year][key] = row
    # Convert dicts to lists
    return {year: list(rows.values()) for year, rows in by_year.items()}, debug_info

# ============================================================
# 6. MERGE FUNCTION
# ============================================================

def merge_by_year(dicts_by_year):
    """
    Merge multiple by_year dicts (from different sources).
    Deduplicate by key (id/doi/title+year) and combine matched_phrases & sources.
    """
    merged = {}
    for by_year in dicts_by_year:
        for year, rows in by_year.items():
            if year not in merged:
                merged[year] = {}
            for row in rows:
                key = row.get("id") or row.get("doi_url") or f"{row['title'].strip().lower()}|{row['year']}"
                if key in merged[year]:
                    existing = merged[year][key]
                    # Combine matched_phrases
                    for phrase in row.get("matched_phrases", []):
                        if phrase not in existing["matched_phrases"]:
                            existing["matched_phrases"].append(phrase)
                    # Combine sources
                    for src in row.get("sources", []):
                        if src not in existing["sources"]:
                            existing["sources"].append(src)
                else:
                    merged[year][key] = row
    # Convert back to list of rows per year
    return {year: list(rows.values()) for year, rows in merged.items()}

# ============================================================
# 7. EXPORT FUNCTIONS (updated with Source column)
# ============================================================

def sanitize_for_excel(value):
    if isinstance(value, str):
        return re.sub(r'[\000-\010]|[\013-\014]|[\016-\037]', '', value)
    return value

def write_excel_sheet(ws, rows, header, bold):
    """Write data to a worksheet and format as an Excel Table."""
    for col_idx, h in enumerate(header, start=1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = bold
    for r_idx, row in enumerate(rows, start=2):
        ws.cell(row=r_idx, column=1, value=sanitize_for_excel(row["title"]))
        ws.cell(row=r_idx, column=2, value=sanitize_for_excel(row["authors"]))
        ws.cell(row=r_idx, column=3, value=sanitize_for_excel(row["journal"]))
        ws.cell(row=r_idx, column=4, value=sanitize_for_excel(row["abstract"]))
        # DOI
        doi_cell = ws.cell(row=r_idx, column=5, value=row["doi_url"])
        if row["doi_url"]:
            doi_cell.hyperlink = row["doi_url"]
            doi_cell.font = Font(color="0563C1", underline="single")
        # PDF
        pdf_cell = ws.cell(row=r_idx, column=6, value=row["pdf_url"])
        if row["pdf_url"]:
            pdf_cell.hyperlink = row["pdf_url"]
            pdf_cell.font = Font(color="0563C1", underline="single")
        # Matched Phrases
        ws.cell(row=r_idx, column=7, value="; ".join(row.get("matched_phrases", [])))
        # Source(s)
        sources = row.get("sources", [row.get("source", "Unknown")])
        ws.cell(row=r_idx, column=8, value=", ".join(sources))
    # Set column widths
    widths = [50, 30, 30, 70, 35, 45, 25, 15]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    max_row = ws.max_row
    max_col = ws.max_column
    if max_row > 1:
        table = Table(displayName=f"Table_{ws.title.replace(' ', '_')}", ref=f"A1:{get_column_letter(max_col)}{max_row}")
        style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                               showLastColumn=False, showRowStripes=True, showColumnStripes=False)
        table.tableStyleInfo = style
        ws.add_table(table)

def export_to_excel_bytes(merged_rows, all_queries, all_years, email=None, separate_tabs=False):
    wb = Workbook()
    if wb.active:
        wb.remove(wb.active)

    header = ["Title", "Authors", "Journal", "Abstract", "DOI", "PDF URL", "Matched Phrase(s)", "Source(s)"]
    bold = Font(bold=True)
    missing_pdfs_rows = []

    if separate_tabs:
        by_year = {}
        for year, row in merged_rows:
            by_year.setdefault(year, []).append(row)
        for year in sorted(by_year.keys()):
            ws = wb.create_sheet(title=str(year))
            rows = sorted(by_year[year], key=lambda r: r["title"].lower())
            write_excel_sheet(ws, rows, header, bold)
            for r in rows:
                if not r.get("pdf_url"):
                    missing_pdfs_rows.append(r)
    else:
        ws = wb.create_sheet(title="All Records")
        rows = sorted([r[1] for r in merged_rows], key=lambda r: (r["year"], r["title"].lower()))
        write_excel_sheet(ws, rows, header, bold)
        for r in rows:
            if not r.get("pdf_url"):
                missing_pdfs_rows.append(r)

    # Settings sheet
    settings_ws = wb.create_sheet(title="Search Settings")
    settings_ws.cell(row=1, column=1, value="Setting").font = bold
    settings_ws.cell(row=1, column=2, value="Value").font = bold
    settings_ws.cell(row=2, column=1, value="Queries searched")
    settings_ws.cell(row=2, column=2, value="; ".join(all_queries))
    settings_ws.cell(row=3, column=1, value="Publication years")
    settings_ws.cell(row=3, column=2, value=", ".join(str(y) for y in all_years))
    settings_ws.cell(row=4, column=1, value="OpenAlex email (polite pool)")
    settings_ws.cell(row=4, column=2, value=email or "Not provided")
    settings_ws.cell(row=5, column=1, value="API Key used")
    settings_ws.cell(row=5, column=2, value="Yes (server-side)" if st.secrets.get("OPENALEX_API_KEY") else "No")
    settings_ws.cell(row=6, column=1, value="Total included records")
    settings_ws.cell(row=6, column=2, value=len(merged_rows))
    settings_ws.column_dimensions["A"].width = 30
    settings_ws.column_dimensions["B"].width = 60

    # Missing PDFs sheet
    missing_ws = wb.create_sheet(title="Missing PDFs")
    missing_headers = ["Title", "Authors", "Year", "DOI"]
    for col_idx, h in enumerate(missing_headers, start=1):
        missing_ws.cell(row=1, column=col_idx, value=h).font = bold
    if missing_pdfs_rows:
        for r_idx, r in enumerate(sorted(missing_pdfs_rows, key=lambda r: (r["year"], r["title"].lower())), start=2):
            missing_ws.cell(row=r_idx, column=1, value=sanitize_for_excel(r["title"]))
            missing_ws.cell(row=r_idx, column=2, value=sanitize_for_excel(r["authors"]))
            missing_ws.cell(row=r_idx, column=3, value=r["year"])
            doi_cell = missing_ws.cell(row=r_idx, column=4, value=r["doi_url"])
            if r["doi_url"]:
                doi_cell.hyperlink = r["doi_url"]
                doi_cell.font = Font(color="0563C1", underline="single")
    else:
        missing_ws.cell(row=2, column=1, value="No missing PDFs in the download list.")
    for i, w in enumerate([50, 30, 10, 35], start=1):
        missing_ws.column_dimensions[get_column_letter(i)].width = w

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out

def export_to_csv_bytes(merged_rows):
    """Export all records as a single CSV file."""
    data = []
    for year, r in merged_rows:
        sources = r.get("sources", [r.get("source", "Unknown")])
        data.append({
            "Year": year,
            "Title": r["title"],
            "Authors": r["authors"],
            "Journal": r["journal"],
            "Abstract": r["abstract"],
            "DOI": r["doi_url"],
            "PDF URL": r["pdf_url"],
            "Matched Phrases": "; ".join(r.get("matched_phrases", [])),
            "Source(s)": ", ".join(sources),
        })
    df = pd.DataFrame(data)
    df = df.sort_values(["Year", "Title"]).reset_index(drop=True)
    out = StringIO()
    df.to_csv(out, index=False)
    return out.getvalue().encode('utf-8')

# ============================================================
# 8. PRE‑FLIGHT COUNT (always both sources)
# ============================================================

def get_total_count(queries, years, email, api_key, work_types):
    total = 0
    # OpenAlex
    for year in years:
        for query in queries:
            count, _ = search_openalex_query_year(query, year, email=email, api_key=api_key,
                                                  work_types=work_types, just_count=True)
            total += count
    # ERIC
    for year in years:
        for query in queries:
            count, _ = search_eric_query_year(query, year, just_count=True)
            total += count
    return total

# ============================================================
# 9. STREAMLIT UI
# ============================================================

st.set_page_config(page_title="LitFind", layout="wide")

st.markdown(
    """
    <style>
    .block-container {
        padding-top: 3.5rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

api_key = st.secrets.get("OPENALEX_API_KEY")
if not api_key:
    st.error("🚨 **API Key Missing!** Please add OPENALEX_API_KEY to your secrets.")
    st.stop()

st.markdown(
    """
    <div style="display: flex; align-items: center; gap: 0.7rem; margin-bottom: 0.2rem;">
        <svg width="38" height="38" viewBox="0 0 40 40" style="flex-shrink: 0;">
            <mask id="searchMask">
                <rect width="40" height="40" fill="white"/>
                <circle cx="17" cy="17" r="8" fill="none" stroke="black" stroke-width="3.5"/>
                <line x1="23" y1="23" x2="31" y2="31" stroke="black" stroke-width="3.5" stroke-linecap="round"/>
            </mask>
            <rect width="40" height="40" rx="9" fill="#2563eb" mask="url(#searchMask)"/>
        </svg>
        <span style="font-size: 1.9rem; font-weight: 700; letter-spacing: -0.01em; line-height: 1.3; color: #1a1a2e;">LitFind</span>
    </div>
    """,
    unsafe_allow_html=True,
)
st.markdown("Search scholarly works from **OpenAlex** and **ERIC** simultaneously. Enter your queries with field‑specific syntax (e.g., `title:\"blended learning\" AND author:\"Smith\"`).")

with st.expander("ℹ️ About this search tool", expanded=False):
    st.markdown("""
    This tool searches **OpenAlex** (broad scholarly index) and **ERIC** (US education database) and merges results automatically.

    **Search syntax:**
    - Use `title:`, `author:`, `journal:`, `abstract:` to limit fields.
    - Combine with `AND`, `OR`, `NOT` and parentheses.
    - Use double quotes for exact phrases.
    - Example: `title:"blended learning" AND author:"Smith" AND year:2022`

    **Features:**
    - Multiple search sessions, each with its own set of queries.
    - Include/exclude checkboxes per row; apply to build a download list.
    - Export as Excel (single sheet or tabs by year) or CSV.
    - Optional email for OpenAlex polite pool.
    - Document type filtering (OpenAlex only).

    **Pre‑flight check:** The app first counts matches; if > 2000, you'll be warned.
    """)

# ---- Session state ----
if "email" not in st.session_state:
    st.session_state.email = ""
if "search_sessions" not in st.session_state:
    st.session_state.search_sessions = []
if "do_full_fetch" not in st.session_state:
    st.session_state.do_full_fetch = False
if "force_refresh" not in st.session_state:
    st.session_state.force_refresh = False
if "download_rows" not in st.session_state:
    st.session_state.download_rows = []

# ---- FORM ----
with st.form("search_form"):
    col1, col2 = st.columns([1, 1])
    with col1:
        queries_input = st.text_area(
            "🔍 Search Queries",
            value='title:"epistemic cognition" AND author:"EFL"',
            help="Each line is a separate query. Use field prefixes (title:, author:, journal:, abstract:) and boolean operators."
        )
        start_year = st.number_input("Start Year", min_value=1900, max_value=2030, value=2020, step=1)
        end_year = st.number_input("End Year", min_value=1900, max_value=2030, value=2026, step=1)
        force_refresh = st.checkbox("🔄 Force Refresh (Ignore Cache)", value=False)
    with col2:
        email = st.text_input(
            "📧 Recommended: Your Email (for OpenAlex polite pool)",
            value=st.session_state.email,
            help="Optional, but using a real email gives you 10x more daily searches."
        )
        st.caption("**Providing an email is optional** – without it, you'll have a lower daily quota.")
        work_type_options = ["All types", "article", "book", "book-chapter", "dataset",
                             "dissertation", "preprint", "conference-paper", "conference-abstract",
                             "book-review", "report", "editorial", "letter", "erratum"]
        work_types = st.multiselect(
            "📚 Work Types (optional – OpenAlex only)",
            options=work_type_options,
            default=["All types"],
            help="Filter by document type. 'All types' means no filter. Only applies to OpenAlex."
        )
        if "All types" in work_types:
            work_types = ["All types"]
    submitted = st.form_submit_button("🚀 Run Search", use_container_width=True)

if submitted:
    queries = [q.strip() for q in queries_input.splitlines() if q.strip()]
    if not queries:
        st.error("Please enter at least one search query.")
        st.stop()
    years = list(range(int(start_year), int(end_year) + 1))
    st.session_state.email = email
    st.session_state.queries = queries
    st.session_state.years = years
    st.session_state.work_types = work_types
    st.session_state.force_refresh = force_refresh

    with st.status("🔎 Checking search scope...", expanded=True) as status:
        try:
            total_count = get_total_count(queries, years, email, api_key, work_types)
            st.session_state.preflight_count = total_count
            status.update(label=f"🔎 Found approximately {total_count} matching works (OpenAlex + ERIC).", state="running")
        except Exception as e:
            st.error(f"Pre‑flight count failed: {e}")
            st.stop()

    THRESHOLD = 2000
    if total_count > THRESHOLD:
        st.warning(f"⚠️ **Search too broad!** ~{total_count} works. Please narrow your search.")
        force_check = st.checkbox("⚠️ **Force full fetch anyway** (I understand the risks)")
        if force_check:
            if st.button("📥 Fetch all records"):
                st.session_state.do_full_fetch = True
                st.rerun()
        else:
            st.stop()
    else:
        st.session_state.do_full_fetch = True
        st.rerun()

if st.session_state.do_full_fetch:
    queries = st.session_state.queries
    years = st.session_state.years
    work_types = st.session_state.work_types
    email = st.session_state.email
    force_refresh = st.session_state.force_refresh

    @st.cache_data(show_spinner=False)
    def run_collection_cached(queries_tuple, years_tuple, email, api_key, work_types_tuple, refresh_seed):
        # Always search both OpenAlex and ERIC
        by_year_openalex, _ = collect(
            list(queries_tuple), list(years_tuple),
            search_openalex_query_year, "OpenAlex",
            email=email, api_key=api_key, work_types=list(work_types_tuple)
        )
        by_year_eric, _ = collect(
            list(queries_tuple), list(years_tuple),
            search_eric_query_year, "ERIC",
            email=email, api_key=api_key, work_types=None
        )
        merged = merge_by_year([by_year_openalex, by_year_eric])
        flat = []
        for year, rows in merged.items():
            for row in rows:
                flat.append((year, row))
        return flat

    with st.status("⏳ Fetching full records from OpenAlex
