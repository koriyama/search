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
from datetime import datetime

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
# 2. SEARCH FUNCTION – handles multiple work types
# ============================================================

def search_openalex_phrase_year(phrase, year, email=None, api_key=None, per_page=200, 
                                session=None, sleep_between=0.1, work_types=None,
                                just_count=False):
    session = session or requests.Session()
    cursor = "*"
    count = 0

    filter_parts = [
        f'title_and_abstract.search:"{phrase}"',
        f'publication_year:{year}'
    ]
    if work_types and "All types" not in work_types:
        type_filter = "|".join(work_types)
        filter_parts.append(f'type:{type_filter}')
    
    filter_string = ",".join(filter_parts)
    
    params = {
        "filter": filter_string,
        "per-page": 1 if just_count else per_page
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
        except requests.exceptions.HTTPError as e:
            st.error(f"❌ OpenAlex API error for count of '{phrase}' (year {year}):")
            st.error(f"Status code: {resp.status_code}")
            st.error(f"Response body (first 500 chars):\n{resp.text[:500]}")
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
# 3. COLLECT FUNCTION
# ============================================================

def collect(phrases, years, email=None, api_key=None, sleep_between=0.1, work_types=None, 
            fetch_fn=search_openalex_phrase_year):
    by_year = {y: {} for y in years}
    debug_info = {}
    
    for year in years:
        for phrase in phrases:
            rows, count, debug_url = fetch_fn(
                phrase, year, email=email, api_key=api_key, 
                sleep_between=sleep_between, work_types=work_types,
                just_count=False
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
# 4. EXPORT
# ============================================================

ILLEGAL_CHARACTERS_RE = re.compile(r'[\000-\010]|[\013-\014]|[\016-\037]')

def sanitize_for_excel(value):
    if isinstance(value, str):
        return ILLEGAL_CHARACTERS_RE.sub('', value)
    return value

def export_to_excel_bytes(flat_rows, phrases, years, out_bytes, email=None, exclude_indices=None):
    if exclude_indices is None:
        exclude_indices = set()
    
    wb = Workbook()
    wb.remove(wb.active)

    header = ["Title", "Authors", "Journal", "Abstract", "DOI", "PDF URL", "Matched Phrase(s)"]
    bold = Font(bold=True)
    missing_pdfs_rows = []

    filtered_flat = [(year, row) for idx, (year, row) in enumerate(flat_rows) if idx not in exclude_indices]

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
# 5. PRE‑FLIGHT COUNT
# ============================================================

def get_total_count(phrases, years, email, api_key, work_types, fetch_fn):
    total = 0
    for year in years:
        for phrase in phrases:
            count, _ = fetch_fn(phrase, year, email=email, api_key=api_key,
                                work_types=work_types, just_count=True)
            total += count
    return total

# ============================================================
# 6. MERGE HELPER
# ============================================================

def merge_flat_rows(existing, new_rows):
    """Merge new rows into existing, deduplicating by id/doi/title+year."""
    seen = set()
    for year, row in existing:
        key = row.get("id") or row.get("doi_url") or f"{row['title'].strip().lower()}|{row['year']}"
        seen.add(key)
    merged = existing.copy()
    for year, row in new_rows:
        key = row.get("id") or row.get("doi_url") or f"{row['title'].strip().lower()}|{row['year']}"
        if key not in seen:
            seen.add(key)
            merged.append((year, row))
    return merged

# ============================================================
# 7. STREAMLIT UI
# ============================================================

st.set_page_config(page_title="Literature Search", layout="wide")

# ---- Load API key ----
api_key = st.secrets.get("OPENALEX_API_KEY")

if not api_key:
    st.error("🚨 **API Key Missing!** Please add OPENALEX_API_KEY to your secrets.")
    st.stop()

# ---- Minimal CSS: floating card for the main content only ----
st.markdown("""
<style>
    .main-card {
        background: white;
        border-radius: 12px;
        padding: 1.5rem 2rem 2rem 2rem;
        box-shadow: 0 4px 20px rgba(0,0,0,0.08);
        border: 1px solid #e6e9ef;
        margin-top: 1rem;
        margin-bottom: 2rem;
    }
    .stButton button {
        width: 100%;
    }
</style>
""", unsafe_allow_html=True)

# ---- Title and blurb (outside card) ----
st.title("📚 Literature Search")
st.markdown("Search for scholarly works using **OpenAlex**. Enter your search phrases and years to find relevant publications – abstracts, authors, journals and PDF links are all included.")

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
    
    🛡️ **Pre‑flight check:** The app first counts how many works match your search. If the count exceeds 2000, it warns you to narrow your search to avoid excessive API calls. You can still force a full fetch if needed.
    
    ℹ️ **Session‑only storage**: Your email and search history are stored **only in your current browser session**. If you close this tab or refresh the page, they will be cleared. No data is stored on any server or shared with anyone.
    
    🔄 **Search Sessions**: Each new search is saved as a separate session. Newest sessions appear at the top of the list. You can view each one individually, merge selected sessions, or delete unwanted ones.
    """)

# ---- Wrap the interactive area in a card container ----
with st.container():
    st.markdown('<div class="main-card">', unsafe_allow_html=True)

    # ---- Initialise session state ----
    if "email" not in st.session_state:
        st.session_state.email = ""
    if "search_history" not in st.session_state:
        st.session_state.search_history = []
    if "preflight_count" not in st.session_state:
        st.session_state.preflight_count = 0
    if "force_fetch" not in st.session_state:
        st.session_state.force_fetch = False
    if "do_full_fetch" not in st.session_state:
        st.session_state.do_full_fetch = False

    # ---- Search sessions (store newest first) ----
    if "search_sessions" not in st.session_state:
        st.session_state.search_sessions = []  # list of dicts, newest first
    if "current_session_index" not in st.session_state:
        st.session_state.current_session_index = None

    # ---- THE MAIN FORM ----
    with st.form("search_form"):
        col1, col2 = st.columns([1, 1])
        with col1:
            phrases_input = st.text_area(
                "🔍 Search Phrases",
                value='"epistemic cognition" AND EFL',
                help="Each line is a separate query. Use uppercase AND, OR, NOT. Put double quotes around exact phrases."
            )
            start_year = st.number_input("Start Year", min_value=1900, max_value=2030, value=2020, step=1)
            end_year = st.number_input("End Year", min_value=1900, max_value=2030, value=2026, step=1)
            force_refresh = st.checkbox("🔄 Force Refresh (Ignore Cache)", value=False)
            
        with col2:
            email = st.text_input(
                "📧 Recommended: Your Email (for OpenAlex polite pool)",
                value=st.session_state.email,
                help="Use your real email address (e.g., name@university.edu). Without it, you'll have only 10 requests per day."
            )
            st.caption("**Use of a real email gives you 10x more daily searches.**")
            
            work_type_options = ["All types", "article", "book", "book-chapter", "dataset", 
                                 "dissertation", "preprint", "conference-paper", "conference-abstract",
                                 "book-review", "report", "editorial", "letter", "erratum"]
            work_types = st.multiselect(
                "📚 Work Types (optional – select multiple)",
                options=work_type_options,
                default=["All types"],
                help="Filter results by one or more document types. 'All types' means no filter."
            )
            if "All types" in work_types:
                work_types = ["All types"]

        submitted = st.form_submit_button("🚀 Run Search", use_container_width=True)

    # ---- HANDLE FORM SUBMISSION ----
    if submitted:
        if not email or email.strip() == "":
            st.error("❌ **Please enter your email address.**")
            st.error("Without an email, OpenAlex limits you to only ~10 requests per day. Enter your email and try again.")
            st.stop()

        phrases = [p.strip() for p in phrases_input.splitlines() if p.strip()]
        if not phrases:
            st.error("Please enter at least one search phrase.")
            st.stop()

        years = list(range(int(start_year), int(end_year) + 1))

        st.session_state.email = email
        st.session_state.phrases = phrases
        st.session_state.years = years
        st.session_state.work_types = work_types
        st.session_state.force_refresh = force_refresh

        # Pre‑flight count
        with st.status("🔎 Checking search scope...", expanded=True) as status:
            try:
                total_count = get_total_count(
                    phrases,
                    years,
                    email,
                    api_key,
                    work_types,
                    search_openalex_phrase_year
                )
                st.session_state.preflight_count = total_count
                status.update(label=f"🔎 Found approximately {total_count} matching works across all phrases and years.", state="running")
            except Exception as e:
                st.error(f"Pre‑flight count failed: {e}")
                st.stop()

        THRESHOLD = 2000
        if total_count > THRESHOLD:
            st.warning(f"⚠️ **Search too broad!** This search would return approximately **{total_count}** works, which is above the safety threshold of **{THRESHOLD}**. This may consume a large number of API requests and take a long time.")
            st.warning("Please narrow your search by:")
            st.warning("- Reducing the year range")
            st.warning("- Using more specific phrases (e.g., `\"deep learning\"` instead of `learning`)")
            st.warning("- Combining terms with `AND` (e.g., `\"climate change\" AND adaptation`)")
            st.warning("- Limiting the number of phrases per line")

            force_check = st.checkbox("⚠️ **Force full fetch anyway** (I understand the risks)")
            if force_check:
                if st.button("📥 Fetch all records", use_container_width=True):
                    st.session_state.do_full_fetch = True
                    st.rerun()
            else:
                st.stop()
        else:
            st.session_state.do_full_fetch = True
            st.rerun()

    # ---- PERFORM FULL FETCH AND CREATE A NEW SESSION (insert at beginning) ----
    if st.session_state.do_full_fetch:
        phrases = st.session_state.phrases
        years = st.session_state.years
        work_types = st.session_state.work_types
        email = st.session_state.email
        force_refresh = st.session_state.force_refresh

        @st.cache_data(show_spinner=False)
        def run_collection_cached(phrases_tuple, years_tuple, email, api_key, work_types_tuple, refresh_seed):
            return collect(list(phrases_tuple), list(years_tuple), email=email, api_key=api_key, work_types=list(work_types_tuple))

        with st.status("⏳ Fetching full records...", expanded=True) as status:
            refresh_seed = random.randint(0, 999999) if force_refresh else 0
            results_by_year, debug_info = run_collection_cached(
                tuple(phrases),
                tuple(years),
                email,
                api_key,
                tuple(work_types),
                refresh_seed
            )
            seen_keys = set()
            flat_rows = []
            for year, rows in results_by_year.items():
                for r in rows:
                    key = r.get("id") or r.get("doi_url") or f"{r['title'].strip().lower()}|{r['year']}"
                    if key not in seen_keys:
                        seen_keys.add(key)
                        flat_rows.append((year, r))
            total = len(flat_rows)
            status.update(label=f"✅ Done! Fetched {total} records.", state="complete")

        # Create new session and insert at index 0 (newest first)
        new_session = {
            "id": datetime.now().strftime("%Y%m%d_%H%M%S"),
            "phrases": phrases,
            "years": years,
            "work_types": work_types,
            "flat_rows": flat_rows,
            "total": total,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "excluded_indices": set()
        }
        st.session_state.search_sessions.insert(0, new_session)  # prepend
        st.session_state.current_session_index = 0  # show the newest
        st.session_state.do_full_fetch = False
        st.rerun()

    # ---- DISPLAY SESSIONS AND RESULTS ----
    if st.session_state.search_sessions:
        # ---- Session management expander ----
        with st.expander("📋 Search Sessions", expanded=True):
            st.write("Select which session to view, or merge multiple sessions. Newest sessions are at the top.")

            # Create a dataframe for the list (already in newest-first order)
            session_data = []
            for idx, sess in enumerate(st.session_state.search_sessions):
                session_data.append({
                    "Select": False,
                    "ID": sess["id"][:8],
                    "Phrases": "; ".join(sess["phrases"]),
                    "Years": f"{sess['years'][0]}-{sess['years'][-1]}",
                    "Records": sess["total"],
                    "Timestamp": sess["timestamp"]
                })
            df_sessions = pd.DataFrame(session_data)

            edited_sessions = st.data_editor(
                df_sessions,
                column_config={
                    "Select": st.column_config.CheckboxColumn("Select for merge", default=False),
                    "ID": st.column_config.TextColumn("ID", width="small"),
                    "Phrases": st.column_config.TextColumn("Phrases", width="large"),
                    "Years": st.column_config.TextColumn("Years", width="small"),
                    "Records": st.column_config.NumberColumn("Records", width="small"),
                    "Timestamp": st.column_config.TextColumn("Timestamp", width="medium"),
                },
                hide_index=True,
                use_container_width=True,
                height=200,
                disabled=["ID", "Phrases", "Years", "Records", "Timestamp"],
                key="session_editor"
            )

            col_merge, col_delete, col_clear = st.columns([1, 1, 1])
            with col_merge:
                if st.button("🔄 Merge Selected", use_container_width=True):
                    selected_indices = [idx for idx, row in edited_sessions.iterrows() if row["Select"]]
                    if len(selected_indices) < 2:
                        st.warning("Please select at least two sessions to merge.")
                    else:
                        merged_rows = []
                        for idx in selected_indices:
                            merged_rows.extend(st.session_state.search_sessions[idx]["flat_rows"])
                        merged_rows = merge_flat_rows([], merged_rows)
                        combined_phrases = []
                        for idx in selected_indices:
                            combined_phrases.extend(st.session_state.search_sessions[idx]["phrases"])
                        combined_phrases = list(dict.fromkeys(combined_phrases))
                        new_session = {
                            "id": datetime.now().strftime("%Y%m%d_%H%M%S"),
                            "phrases": combined_phrases,
                            "years": st.session_state.search_sessions[selected_indices[0]]["years"],
                            "work_types": st.session_state.search_sessions[selected_indices[0]]["work_types"],
                            "flat_rows": merged_rows,
                            "total": len(merged_rows),
                            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M"),
                            "excluded_indices": set()
                        }
                        st.session_state.search_sessions.insert(0, new_session)
                        st.session_state.current_session_index = 0
                        st.rerun()

            with col_delete:
                if st.button("🗑️ Delete Selected", use_container_width=True):
                    selected_indices = [idx for idx, row in edited_sessions.iterrows() if row["Select"]]
                    if not selected_indices:
                        st.warning("Select at least one session to delete.")
                    else:
                        for idx in sorted(selected_indices, reverse=True):
                            del st.session_state.search_sessions[idx]
                        if st.session_state.current_session_index in selected_indices or st.session_state.current_session_index is None:
                            st.session_state.current_session_index = 0 if st.session_state.search_sessions else None
                        st.rerun()

            with col_clear:
                if st.button("🧹 Clear All", use_container_width=True):
                    st.session_state.search_sessions = []
                    st.session_state.current_session_index = None
                    st.rerun()

            # Session selector dropdown (index 0 is always the newest)
            session_options = [f"{sess['id'][:8]} – {sess['phrases'][0][:40]}… ({sess['total']} records)" for sess in st.session_state.search_sessions]
            if session_options:
                current_idx = st.session_state.current_session_index if st.session_state.current_session_index is not None else 0
                selected_idx = st.selectbox(
                    "View session:",
                    options=range(len(session_options)),
                    index=min(current_idx, len(session_options)-1),
                    format_func=lambda i: session_options[i],
                    key="session_selector"
                )
                if selected_idx != st.session_state.current_session_index:
                    st.session_state.current_session_index = selected_idx
                    st.rerun()

        # ---- DISPLAY THE CURRENT SESSION ----
        if st.session_state.current_session_index is not None and st.session_state.current_session_index < len(st.session_state.search_sessions):
            session = st.session_state.search_sessions[st.session_state.current_session_index]
            flat_rows = session["flat_rows"]
            total = session["total"]
            phrases = session["phrases"]
            years = session["years"]
            email = st.session_state.email
            excluded_indices = session.get("excluded_indices", set())

            if total > 0:
                all_rows = []
                for year, r in flat_rows:
                    all_rows.append({
                        "Year": year,
                        "Title": r["title"],
                        "Authors": r["authors"],
                        "Journal": r["journal"],
                        "Abstract": r["abstract"],
                        "Has PDF": "✅" if r["pdf_url"] else "❌",
                        "Phrases": "; ".join(r.get("matched_phrases", []))
                    })
                df = pd.DataFrame(all_rows)

                st.subheader(f"📄 Preview of Results ({total} hits)")
                st.caption("Check the box next to any row you want to **exclude** from the final download.")

                df_with_checkboxes = df.copy()
                df_with_checkboxes.insert(0, "Exclude", False)
                for idx in excluded_indices:
                    if idx < len(df_with_checkboxes):
                        df_with_checkboxes.at[idx, "Exclude"] = True

                editor_key = f"data_editor_{session['id']}"

                colA, colB = st.columns([1, 1])
                if colA.button("✖️ Exclude All", use_container_width=True, key=f"exclude_all_{session['id']}"):
                    df_with_checkboxes["Exclude"] = True
                    st.session_state[f"df_editor_{session['id']}"] = df_with_checkboxes
                    st.rerun()
                if colB.button("✅ Include All", use_container_width=True, key=f"include_all_{session['id']}"):
                    df_with_checkboxes["Exclude"] = False
                    st.session_state[f"df_editor_{session['id']}"] = df_with_checkboxes
                    st.rerun()

                if f"df_editor_{session['id']}" not in st.session_state:
                    st.session_state[f"df_editor_{session['id']}"] = df_with_checkboxes

                edited_df = st.data_editor(
                    st.session_state[f"df_editor_{session['id']}"],
                    use_container_width=True,
                    height=400,
                    column_config={
                        "Exclude": st.column_config.CheckboxColumn(
                            "Exclude", 
                            width=80,
                            help="Check to exclude this row from export"
                        ),
                        "Year": st.column_config.NumberColumn("Year", width="small"),
                        "Title": st.column_config.TextColumn("Title", width="large"),
                        "Authors": st.column_config.TextColumn("Authors", width="medium"),
                        "Journal": st.column_config.TextColumn("Journal", width="medium"),
                        "Abstract": st.column_config.TextColumn(
                            "Abstract (double click to expand)", 
                            width="large",
                            disabled=False
                        ),
                        "Has PDF": st.column_config.TextColumn("PDF", width="small"),
                        "Phrases": st.column_config.TextColumn("Matched Phrases", width="medium"),
                    },
                    hide_index=True,
                    key=editor_key
                )

                excluded_indices = set()
                for idx, row in edited_df.iterrows():
                    if row["Exclude"]:
                        excluded_indices.add(idx)
                session["excluded_indices"] = excluded_indices
                st.session_state[f"df_editor_{session['id']}"] = edited_df

                missing_count = sum(1 for r in all_rows if r["Has PDF"] == "❌")
                st.caption(f"📌 {missing_count} records are missing a PDF link (they'll appear in the 'Missing PDFs' sheet).")
                if excluded_indices:
                    st.info(f"🚫 {len(excluded_indices)} row(s) marked for exclusion from the download.")

                # ---- Download section ----
                st.markdown("---")
                col_name, col_btn = st.columns([2, 1])
                with col_name:
                    default_filename = f"{datetime.now().strftime('%Y-%m-%d')}_{phrases[0][:40].replace(' ', '_')}.xlsx"
                    filename = st.text_input(
                        "📁 Filename for download",
                        value=default_filename,
                        help="Customise the file name before downloading.",
                        key=f"filename_{session['id']}"
                    )
                with col_btn:
                    st.write("")
                    st.write("")
                    if st.button("⬇️ Download Excel File", use_container_width=True, key=f"download_{session['id']}"):
                        with st.spinner("Generating Excel..."):
                            output = BytesIO()
                            export_to_excel_bytes(
                                flat_rows,
                                phrases,
                                years,
                                output,
                                email=email,
                                exclude_indices=excluded_indices
                            )
                            output.seek(0)
                            st.download_button(
                                label="📥 Click to save",
                                data=output,
                                file_name=filename if filename.strip() else default_filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True,
                                key=f"download_btn_{session['id']}"
                            )
            else:
                st.info("This session has no results.")
        else:
            st.info("No session selected. Run a search to begin.")
    else:
        # ---- FILL THE EMPTY CARD WITH A HELPFUL MESSAGE ----
        st.info("No searches yet. Fill in the form above and click 'Run Search'.")

    # Close the card container
    st.markdown('</div>', unsafe_allow_html=True)
