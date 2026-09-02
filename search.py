import streamlit as st
import pandas as pd
import requests
import time
import urllib.parse
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
import re
import random

# ============================================================
# 1. CORE FUNCTIONS – unchanged
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

# ============================================================
# 2. SEARCH FUNCTION – correctly encoded, no double‑encoding
# ============================================================

def search_openalex_phrase_year(phrase, year, email=None, api_key=None, per_page=200, 
                                session=None, sleep_between=0.1, work_type=None):
    session = session or requests.Session()
    results = []
    cursor = "*"
    first_page_meta_count = None

    filter_parts = [
        f'title_and_abstract.search:"{phrase}"',
        f'publication_year:{year}'
    ]
    if work_type and work_type != "All types":
        filter_parts.append(f'type:{work_type}')
    
    filter_string = ",".join(filter_parts)
    
    params = {
        "filter": filter_string,
        "per-page": per_page
    }
    
    if email:
        params["mailto"] = email
    
    if api_key:
        params["api_key"] = api_key

    temp_params = dict(params)
    temp_params["cursor"] = "*"
    debug_url = f"{OPENALEX_URL}?{urllib.parse.urlencode(temp_params)}"

    while cursor:
        params["cursor"] = cursor
        
        try:
            resp = session.get(OPENALEX_URL, params=params, timeout=30)
            resp.raise_for_status()
        except requests.exceptions.HTTPError as e:
            st.error(f"❌ OpenAlex API error for phrase '{phrase}' (year {year}):")
            st.error(f"Status code: {resp.status_code}")
            st.error(f"Response body (first 500 chars):\n{resp.text[:500]}")
            raise

        data = resp.json()
        
        if first_page_meta_count is None:
            first_page_meta_count = data.get("meta", {}).get("count", 0)

        for work in data.get("results", []):
            results.append(normalize_work(work))
        
        cursor = (data.get("meta", {}) or {}).get("next_cursor")
        if not data.get("results"):
            break
        time.sleep(sleep_between)

    return results, first_page_meta_count, debug_url

# ============================================================
# 3. COLLECT FUNCTION – returns results_by_year
# ============================================================

def collect(phrases, years, email=None, api_key=None, sleep_between=0.1, work_type=None, 
            fetch_fn=search_openalex_phrase_year):
    by_year = {y: {} for y in years}
    debug_info = {}
    
    for year in years:
        for phrase in phrases:
            rows, count, debug_url = fetch_fn(
                phrase, year, email=email, api_key=api_key, 
                sleep_between=sleep_between, work_type=work_type
            )
            
            debug_info[(phrase, year)] = {
                "count": count,
                "url": debug_url
            }
            
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
    return {year: list(rows.values()) for year, rows in by_year.items()}, debug_info

# ============================================================
# 4. EXPORT – now takes flat_rows (deduped globally)
# ============================================================

ILLEGAL_CHARACTERS_RE = re.compile(r'[\000-\010]|[\013-\014]|[\016-\037]')

def sanitize_for_excel(value):
    if isinstance(value, str):
        return ILLEGAL_CHARACTERS_RE.sub('', value)
    return value

def export_to_excel_bytes(flat_rows, phrases, years, out_bytes, email=None, exclude_indices=None):
    """
    flat_rows: list of (year, row) tuples – already deduplicated globally.
    exclude_indices: set of indices (from the preview) to skip.
    """
    if exclude_indices is None:
        exclude_indices = set()
    
    wb = Workbook()
    wb.remove(wb.active)

    header = ["Title", "Authors", "Journal", "Abstract", "DOI", "PDF URL", "Matched Phrase(s)"]
    bold = Font(bold=True)
    missing_pdfs_rows = []

    # Filter out excluded rows
    filtered_flat = [(year, row) for idx, (year, row) in enumerate(flat_rows) if idx not in exclude_indices]

    # Group by year for sheets
    filtered_by_year = {}
    for year, row in filtered_flat:
        if year not in filtered_by_year:
            filtered_by_year[year] = []
        filtered_by_year[year].append(row)

    for year in years:
        ws = wb.create_sheet(title=str(year))
        for col_idx, h in enumerate(header, start=1):
            ws.cell(row=1, column=col_idx, value=h).font = bold

        rows = sorted(filtered_by_year.get(year, []), key=lambda r: r["title"].lower())
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
    settings_ws.cell(row=4, column=1, value="OpenAlex email (polite pool)")
    settings_ws.cell(row=4, column=2, value=email or "Not provided")
    settings_ws.cell(row=5, column=1, value="API Key used")
    settings_ws.cell(row=5, column=2, value="Yes (server-side)" if st.secrets.get("OPENALEX_API_KEY") else "No")
    settings_ws.cell(row=6, column=1, value="Total unique records (exported)")
    settings_ws.cell(row=6, column=2, value=len(filtered_flat))
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
# 5. STREAMLIT USER INTERFACE
# ============================================================

st.set_page_config(page_title="Literature Search", layout="wide")
st.title("📚 Literature Search")

# ---- SECURELY LOAD THE API KEY ----
api_key = st.secrets.get("OPENALEX_API_KEY")

if not api_key:
    st.error("🚨 **API Key Missing!** Please add OPENALEX_API_KEY to your secrets.")
    st.stop()

# ---- Session state init ----
if "email" not in st.session_state:
    st.session_state.email = ""
if "search_history" not in st.session_state:
    st.session_state.search_history = []

# ---- DISCLAIMER ----
st.info(
    "ℹ️ **Session‑only storage**: Your email and search history are stored **only in your current browser session**. "
    "If you close this tab or refresh the page, they will be cleared. "
    "No data is stored on any server or shared with anyone."
)

# ---- OpenAlex description (minimised by default) ----
with st.expander("ℹ️ About this search tool", expanded=False):
    st.markdown("""
    This tool searches **OpenAlex** – a free, open index of the world's research ecosystem.
    
    **What does OpenAlex cover?**
    - Over **320 million scholarly works**: journal articles, conference papers, books, book chapters, datasets, dissertations, preprints, and more
    - **Extra coverage** of humanities, non‑English languages, and the Global South
    - Data from **Crossref, PubMed, arXiv, HAL, DOAJ, ORCID, institutional repositories**, and many other sources
    - **60 million open access PDFs** parsed directly
    
    **Why OpenAlex?**
    - It is **free and open** – no paywalls, no API keys required (though providing your email gives you faster "polite pool" access)
    - It is **more comprehensive** than Scopus or Web of Science, with over 464 million works indexed
    - It includes **datasets, software, and other research objects** beyond just traditional publications
    
    **Search tips:**
    - Use **uppercase** `AND`, `OR`, `NOT` for boolean logic
    - Use **double quotes** for exact phrase matches (e.g., `"climate change"`)
    - **Parentheses** group terms (e.g., `(neural OR deep)`)
    - Non‑ASCII characters (e.g., 守破離) are fully supported
    - You can filter results by document type using the dropdown below
    
    ⚠️ **Cache & Rate limits:** 
    - The app caches results to speed up repeated searches. 
    - If you are getting **0 results unexpectedly**, check the **"Force Refresh"** box below and search again.
    """)

with st.form("search_form"):
    col1, col2 = st.columns(2)
    with col1:
        phrases_input = st.text_area(
            "🔍 Search Phrases",
            value=(
                'climate AND ("man-made change" OR change)\n'
                '"epistemic cognition" OR "personal epistemology" OR "epistemological beliefs"\n'
                '"climate change" NOT denial\n'
            ),
            help=(
                "Each line is a separate query. Use uppercase AND, OR, NOT. "
                "Put double quotes around exact phrases (e.g., \"climate change\")."
            )
        )
        start_year = st.number_input("Start Year", min_value=1900, max_value=2030, value=2020, step=1)
        end_year = st.number_input("End Year", min_value=1900, max_value=2030, value=2026, step=1)
        
        work_type = st.selectbox(
            "📚 Work Type (optional)",
            options=["All types", "article", "book", "book-chapter", "dataset", 
                     "dissertation", "preprint", "conference-paper", "conference-abstract",
                     "book-review", "report", "editorial", "letter", "erratum"],
            index=0,
            help="Filter results by document type. 'All types' searches everything OpenAlex has."
        )
        
        force_refresh = st.checkbox(
            "🔄 Force Refresh (Ignore Cache)", 
            value=False,
            help="Check this if you're getting 0 results unexpectedly or suspect cached data."
        )
        
    with col2:
        email = st.text_input(
            "📧 Your Email (for OpenAlex polite pool)",
            value=st.session_state.email,
            help="OPTIONAL but recommended: Use your real email address for better performance."
        )
        st.caption("Adding your email gives you access to the 'polite pool' for better performance.")
        st.success("🔒 **API Key: Configured securely** (server-side)")

    submitted = st.form_submit_button("🚀 Run Search", use_container_width=True)

# ---- THE CACHED FUNCTION ----
@st.cache_data(show_spinner=False)
def run_collection(phrases_tuple, years_tuple, email, api_key, work_type, refresh_seed):
    return collect(list(phrases_tuple), list(years_tuple), email=email, api_key=api_key, work_type=work_type)

if submitted:
    st.session_state.email = email
    
    phrases = [p.strip() for p in phrases_input.splitlines() if p.strip()]
    years = list(range(int(start_year), int(end_year) + 1))

    if not phrases:
        st.error("Please enter at least one search phrase.")
        st.stop()

    with st.status("⏳ Searching OpenAlex...", expanded=True) as status:
        refresh_seed = random.randint(0, 999999) if force_refresh else 0
        
        results_by_year, debug_info = run_collection(
            tuple(phrases), 
            tuple(years), 
            email, 
            api_key, 
            work_type,
            refresh_seed
        )
        
        # ---- FLATTEN AND DEDUPLICATE GLOBALLY ----
        seen_keys = set()
        flat_rows = []
        for year, rows in results_by_year.items():
            for r in rows:
                key = r.get("id") or r.get("doi_url") or f"{r['title'].strip().lower()}|{r['year']}"
                if key not in seen_keys:
                    seen_keys.add(key)
                    flat_rows.append((year, r))
        
        total = len(flat_rows)
        status.update(label=f"✅ Done! Found {total} unique records (after deduplication).", state="complete")

    if force_refresh:
        st.success("🔄 Cache bypassed! Results are freshly fetched from OpenAlex.")

    # ---- DEBUG: If total is 0, show the exact URL ----
    if total == 0:
        st.warning("⚠️ **0 results found.** This might be due to a rate limit, a malformed query, or genuinely 0 works.")
        st.warning("Here is the exact URL the app called for the first search (copy and paste it into your browser to test):")
        
        first_key = list(debug_info.keys())[0]
        first_debug = debug_info[first_key]
        st.code(first_debug["url"], language="text")
        st.caption(f"OpenAlex reported count: **{first_debug['count']}** for '{first_key[0]}' in {first_key[1]}.")
        st.caption("If this URL shows results in your browser, the app is having a parsing issue. If it shows 0 or an error, the problem is with your API key/rate limits.")

    # ---- Save to search history ----
    search_record = {
        "phrases": ", ".join(phrases),
        "years": f"{start_year}-{end_year}",
        "timestamp": time.strftime("%Y-%m-%d %H:%M"),
        "total_results": total,
        "work_type": work_type
    }
    st.session_state.search_history.append(search_record)

    # ---- Build preview from flat_rows ----
    if total > 0:
        all_rows = []
        for year, r in flat_rows:
            all_rows.append({
                "Year": year,
                "Title": r["title"],
                "Authors": r["authors"],
                "Journal": r["journal"],
                "Has PDF": "✅" if r["pdf_url"] else "❌",
                "Phrases": "; ".join(r.get("matched_phrases", []))
            })
        df = pd.DataFrame(all_rows)

        st.subheader("📄 Preview of Results")
        st.caption("Check the box next to any row you want to **exclude** from the final download.")

        # ---- Add Exclude column ----
        df_with_checkboxes = df.copy()
        df_with_checkboxes.insert(0, "Exclude", False)

        # ---- Manage session state for the editor ----
        # If the number of rows changed (new search), reset the editor state
        if "df_editor" not in st.session_state or len(st.session_state.df_editor) != len(df_with_checkboxes):
            st.session_state.df_editor = df_with_checkboxes

        # ---- Bulk action buttons ----
        colA, colB, colC = st.columns([1, 1, 3])
        if colA.button("✖️ Exclude All"):
            st.session_state.df_editor["Exclude"] = True
        if colB.button("✅ Include All"):
            st.session_state.df_editor["Exclude"] = False

        # ---- Data editor ----
        edited_df = st.data_editor(
            st.session_state.df_editor,
            use_container_width=True,
            height=400,
            column_config={
                "Exclude": st.column_config.CheckboxColumn("Exclude", help="Check to exclude this row from export"),
                "Title": st.column_config.TextColumn("Title", width="large"),
                "Authors": st.column_config.TextColumn("Authors", width="medium"),
                "Journal": st.column_config.TextColumn("Journal", width="medium"),
                "Year": st.column_config.NumberColumn("Year", width="small"),
                "Has PDF": st.column_config.TextColumn("PDF", width="small"),
                "Phrases": st.column_config.TextColumn("Matched Phrases", width="medium"),
            },
            hide_index=True,
        )

        # ---- Update session state with manual edits ----
        st.session_state.df_editor = edited_df

        # ---- Calculate excluded indices ----
        excluded_indices = set()
        for idx, row in edited_df.iterrows():
            if row["Exclude"]:
                excluded_indices.add(idx)

        missing_count = sum(1 for r in all_rows if r["Has PDF"] == "❌")
        st.caption(f"📌 {missing_count} records are missing a PDF link (they'll appear in the 'Missing PDFs' sheet).")
        if excluded_indices:
            st.info(f"🚫 {len(excluded_indices)} row(s) marked for exclusion from the download.")

        # ---- Export button ----
        if st.button("⬇️ Download Excel File", use_container_width=True):
            with st.spinner("Generating Excel..."):
                output = BytesIO()
                export_to_excel_bytes(
                    flat_rows,        # pass the deduped flat list
                    phrases, 
                    years, 
                    output, 
                    email=email,
                    exclude_indices=excluded_indices
                )
                output.seek(0)
                st.download_button(
                    label="📥 Click here to save the file",
                    data=output,
                    file_name=f"Literature_Search_{start_year}-{end_year}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    else:
        st.info("No results to preview.")

# ---- Display search history ----
if st.session_state.search_history:
    with st.expander("📜 Search History (this session only)"):
        for i, record in enumerate(reversed(st.session_state.search_history)):
            st.write(f"**{i+1}.** {record['timestamp']} – **{record['phrases']}** ({record['years']}) → {record['total_results']} results | Type: {record['work_type']}")
        st.caption("This history is stored only in your browser session and will be cleared when you close the tab.")
