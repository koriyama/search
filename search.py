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
# 2. SEARCH FUNCTION
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
# 4. EXPORT – now exports only included rows from merged set
# ============================================================

def export_to_excel_bytes(flat_rows, phrases, years, out_bytes, email=None):
    # flat_rows here is the merged set (list of (year, row))
    wb = Workbook()
    wb.remove(wb.active)

    header = ["Title", "Authors", "Journal", "Abstract", "DOI", "PDF URL", "Matched Phrase(s)"]
    bold = Font(bold=True)
    missing_pdfs_rows = []

    # group by year (preserve order)
    by_year = {}
    for year, row in flat_rows:
        if year not in by_year:
            by_year[year] = []
        by_year[year].append(row)

    for year in years:
        ws = wb.create_sheet(title=str(year))
        for col_idx, h in enumerate(header, start=1):
            ws.cell(row=1, column=col_idx, value=h).font = bold

        rows = sorted(by_year.get(year, []), key=lambda r: r["title"].lower())
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

    # Settings sheet (optional, but kept)
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
    settings_ws.cell(row=6, column=1, value="Total included records")
    settings_ws.cell(row=6, column=2, value=len(flat_rows))
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

# sanitize helper
ILLEGAL_CHARACTERS_RE = re.compile(r'[\000-\010]|[\013-\014]|[\016-\037]')

def sanitize_for_excel(value):
    if isinstance(value, str):
        return ILLEGAL_CHARACTERS_RE.sub('', value)
    return value

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
# 6. STREAMLIT UI
# ============================================================

st.set_page_config(page_title="Literature Search", layout="wide")

api_key = st.secrets.get("OPENALEX_API_KEY")
if not api_key:
    st.error("🚨 **API Key Missing!** Please add OPENALEX_API_KEY to your secrets.")
    st.stop()

# ---- Title and info ----
st.title("📚 Literature Search")
st.markdown("Search for scholarly works using **OpenAlex**. Enter your search phrases and years to find relevant publications – abstracts, authors, journals and PDF links are all included.")

with st.expander("ℹ️ About this search tool", expanded=False):
    st.markdown("""
    This tool searches **OpenAlex** – a free, open index of the world's research ecosystem.
    
    - **Sessions**: each search is saved separately.  
    - **Include rows**: for each session, tick the rows you want to keep.  
    - **Merged set**: the union of all included rows across all sessions is used for download.  
    - Newest sessions appear at the top.
    """)

# ---- Session state ----
if "email" not in st.session_state:
    st.session_state.email = ""
if "search_sessions" not in st.session_state:
    st.session_state.search_sessions = []  # each: {id, phrases, years, work_types, flat_rows, total, timestamp, included_indices: set()}
if "current_session_index" not in st.session_state:
    st.session_state.current_session_index = None
if "do_full_fetch" not in st.session_state:
    st.session_state.do_full_fetch = False
if "force_refresh" not in st.session_state:
    st.session_state.force_refresh = False

# ---- FORM ----
with st.form("search_form"):
    col1, col2 = st.columns([1, 1])
    with col1:
        phrases_input = st.text_area(
            "🔍 Search Phrases",
            value='"epistemic cognition" AND EFL',
            help="Each line is a separate query."
        )
        start_year = st.number_input("Start Year", min_value=1900, max_value=2030, value=2020, step=1)
        end_year = st.number_input("End Year", min_value=1900, max_value=2030, value=2026, step=1)
        force_refresh = st.checkbox("🔄 Force Refresh (Ignore Cache)", value=False)
    with col2:
        email = st.text_input(
            "📧 Recommended: Your Email (for OpenAlex polite pool)",
            value=st.session_state.email,
            help="Use your real email address for higher rate limits."
        )
        st.caption("**Use of a real email gives you 10x more daily searches.**")
        work_type_options = ["All types", "article", "book", "book-chapter", "dataset", 
                             "dissertation", "preprint", "conference-paper", "conference-abstract",
                             "book-review", "report", "editorial", "letter", "erratum"]
        work_types = st.multiselect(
            "📚 Work Types (optional – select multiple)",
            options=work_type_options,
            default=["All types"],
            help="Filter by document type. 'All types' means no filter."
        )
        if "All types" in work_types:
            work_types = ["All types"]
    submitted = st.form_submit_button("🚀 Run Search", use_container_width=True)

# ---- Handle submission ----
if submitted:
    if not email or email.strip() == "":
        st.error("❌ **Please enter your email address.**")
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

    # Pre-flight count
    with st.status("🔎 Checking search scope...", expanded=True) as status:
        try:
            total_count = get_total_count(
                phrases, years, email, api_key, work_types, search_openalex_phrase_year
            )
            st.session_state.preflight_count = total_count
            status.update(label=f"🔎 Found approximately {total_count} matching works.", state="running")
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

# ---- Fetch and create new session ----
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
        results_by_year, _ = run_collection_cached(
            tuple(phrases), tuple(years), email, api_key, tuple(work_types), refresh_seed
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
        status.update(label=f"✅ Fetched {total} records.", state="complete")

    new_session = {
        "id": datetime.now().strftime("%Y%m%d_%H%M%S"),
        "phrases": phrases,
        "years": years,
        "work_types": work_types,
        "flat_rows": flat_rows,
        "total": total,
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "included_indices": set()  # empty by default
    }
    st.session_state.search_sessions.insert(0, new_session)
    st.session_state.current_session_index = 0
    st.session_state.do_full_fetch = False
    st.rerun()

# ---- Display current session preview (ABOVE session management) ----
if st.session_state.search_sessions:
    current_idx = st.session_state.current_session_index
    if current_idx is None or current_idx >= len(st.session_state.search_sessions):
        current_idx = 0
        st.session_state.current_session_index = 0
    session = st.session_state.search_sessions[current_idx]
    flat_rows = session["flat_rows"]
    total = session["total"]
    phrases = session["phrases"]
    years = session["years"]
    included_indices = session.get("included_indices", set())

    # ---- Preview of current session ----
    if total > 0:
        all_rows = []
        for i, (year, r) in enumerate(flat_rows):
            all_rows.append({
                "Include": i in included_indices,
                "Year": year,
                "Title": r["title"],
                "Authors": r["authors"],
                "Journal": r["journal"],
                "Abstract": r["abstract"],
                "Has PDF": "✅" if r["pdf_url"] else "❌",
                "Phrases": "; ".join(r.get("matched_phrases", []))
            })
        df = pd.DataFrame(all_rows)

        st.subheader(f"📄 Current Session – {phrases[0]} ({total} hits)")
        st.caption("Tick the **Include** box for rows you want to keep. Included rows from all sessions will be merged for download.")

        # Data editor with "Include" checkbox
        df_with_include = df.copy()
        editor_key = f"include_editor_{session['id']}"

        # Store the edited df in session state
        if f"include_df_{session['id']}" not in st.session_state:
            st.session_state[f"include_df_{session['id']}"] = df_with_include

        edited_df = st.data_editor(
            st.session_state[f"include_df_{session['id']}"],
            use_container_width=True,
            height=400,
            column_config={
                "Include": st.column_config.CheckboxColumn(
                    "Include", 
                    width=80,
                    help="Check to include this row in the merged set"
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

        # Update included_indices from the edited_df
        new_included = set()
        for idx, row in edited_df.iterrows():
            if row["Include"]:
                new_included.add(idx)
        session["included_indices"] = new_included
        st.session_state[f"include_df_{session['id']}"] = edited_df

        # Show count of included rows
        st.caption(f"✅ {len(new_included)} row(s) included from this session.")

        # ---- Quick actions for this session ----
        colA, colB = st.columns([1, 1])
        if colA.button("✅ Include All", use_container_width=True, key=f"include_all_{session['id']}"):
            for idx in range(len(edited_df)):
                edited_df.at[idx, "Include"] = True
            session["included_indices"] = set(range(len(edited_df)))
            st.session_state[f"include_df_{session['id']}"] = edited_df
            st.rerun()
        if colB.button("❌ Include None", use_container_width=True, key=f"include_none_{session['id']}"):
            for idx in range(len(edited_df)):
                edited_df.at[idx, "Include"] = False
            session["included_indices"] = set()
            st.session_state[f"include_df_{session['id']}"] = edited_df
            st.rerun()

    else:
        st.info("This session has no results.")

    # ---- Session management (below the preview) ----
    with st.expander("📋 Manage Sessions", expanded=False):
        st.write("All your search sessions. Newest first.")

        session_data = []
        for idx, sess in enumerate(st.session_state.search_sessions):
            included_count = len(sess.get("included_indices", set()))
            session_data.append({
                "Select": False,
                "ID": sess["id"][:8],
                "Phrases": "; ".join(sess["phrases"])[:60],
                "Years": f"{sess['years'][0]}-{sess['years'][-1]}",
                "Total": sess["total"],
                "Included": included_count,
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
                "Total": st.column_config.NumberColumn("Total", width="small"),
                "Included": st.column_config.NumberColumn("Included", width="small"),
                "Timestamp": st.column_config.TextColumn("Timestamp", width="medium"),
            },
            hide_index=True,
            use_container_width=True,
            height=200,
            disabled=["ID", "Phrases", "Years", "Total", "Included", "Timestamp"],
            key="session_manager"
        )

        col_merge, col_delete, col_clear = st.columns([1, 1, 1])
        with col_merge:
            if st.button("🔄 Merge Selected (keep included rows)", use_container_width=True):
                selected_indices = [idx for idx, row in edited_sessions.iterrows() if row["Select"]]
                if len(selected_indices) < 2:
                    st.warning("Select at least two sessions to merge.")
                else:
                    # Collect all included rows from selected sessions
                    merged_rows = []
                    for idx in selected_indices:
                        sess = st.session_state.search_sessions[idx]
                        included = sess.get("included_indices", set())
                        for row_idx in included:
                            if row_idx < len(sess["flat_rows"]):
                                merged_rows.append(sess["flat_rows"][row_idx])
                    # Deduplicate (by id/doi/title+year)
                    seen = set()
                    unique_merged = []
                    for year, r in merged_rows:
                        key = r.get("id") or r.get("doi_url") or f"{r['title'].strip().lower()}|{r['year']}"
                        if key not in seen:
                            seen.add(key)
                            unique_merged.append((year, r))
                    # Create new session
                    combined_phrases = []
                    for idx in selected_indices:
                        combined_phrases.extend(st.session_state.search_sessions[idx]["phrases"])
                    combined_phrases = list(dict.fromkeys(combined_phrases))
                    new_session = {
                        "id": datetime.now().strftime("%Y%m%d_%H%M%S"),
                        "phrases": combined_phrases,
                        "years": st.session_state.search_sessions[selected_indices[0]]["years"],
                        "work_types": st.session_state.search_sessions[selected_indices[0]]["work_types"],
                        "flat_rows": unique_merged,
                        "total": len(unique_merged),
                        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M"),
                        "included_indices": set(range(len(unique_merged)))  # all included by default
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

        # Switch to another session
        session_options = [f"{sess['id'][:8]} – {sess['phrases'][0][:40]}… (included: {len(sess.get('included_indices', set()))}/{sess['total']})" for sess in st.session_state.search_sessions]
        if session_options:
            current_idx = st.session_state.current_session_index if st.session_state.current_session_index is not None else 0
            selected_idx = st.selectbox(
                "Switch to session:",
                options=range(len(session_options)),
                index=min(current_idx, len(session_options)-1),
                format_func=lambda i: session_options[i],
                key="session_switch"
            )
            if selected_idx != st.session_state.current_session_index:
                st.session_state.current_session_index = selected_idx
                st.rerun()

    # ---- Show merged set and download ----
    # Compute merged rows from all sessions' included indices
    all_included = []
    for sess in st.session_state.search_sessions:
        included = sess.get("included_indices", set())
        for idx in included:
            if idx < len(sess["flat_rows"]):
                all_included.append(sess["flat_rows"][idx])
    # Deduplicate globally
    seen = set()
    merged_rows = []
    for year, r in all_included:
        key = r.get("id") or r.get("doi_url") or f"{r['title'].strip().lower()}|{r['year']}"
        if key not in seen:
            seen.add(key)
            merged_rows.append((year, r))

    st.markdown("---")
    st.subheader(f"📦 Merged Set – {len(merged_rows)} unique included records across all sessions")

    if len(merged_rows) > 0:
        st.caption("This is the union of all rows you marked 'Include' in any session. Download this set below.")
        # Show a quick preview of merged rows (optional)
        with st.expander("Preview merged records", expanded=False):
            merged_df = pd.DataFrame([{"Year": y, "Title": r["title"]} for y, r in merged_rows])
            st.dataframe(merged_df, use_container_width=True, height=200)

        # Download button for merged set
        col_name, col_btn = st.columns([2, 1])
        with col_name:
            default_filename = f"{datetime.now().strftime('%Y-%m-%d')}_merged_{len(merged_rows)}_records.xlsx"
            filename = st.text_input(
                "📁 Filename for download",
                value=default_filename,
                help="Customise the file name before downloading.",
                key="merged_filename"
            )
        with col_btn:
            st.write("")
            st.write("")
            if st.button("⬇️ Download Merged Excel", use_container_width=True, key="download_merged"):
                with st.spinner("Generating Excel..."):
                    output = BytesIO()
                    # Pass the merged_rows and a combined phrases list
                    all_phrases = []
                    for sess in st.session_state.search_sessions:
                        all_phrases.extend(sess["phrases"])
                    all_phrases = list(dict.fromkeys(all_phrases))
                    # Use years from first session (or merge years)
                    all_years = []
                    for sess in st.session_state.search_sessions:
                        all_years.extend(sess["years"])
                    all_years = sorted(set(all_years))
                    export_to_excel_bytes(
                        merged_rows,
                        all_phrases,
                        all_years,
                        output,
                        email=st.session_state.email
                    )
                    output.seek(0)
                    st.download_button(
                        label="📥 Click to save",
                        data=output,
                        file_name=filename if filename.strip() else default_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="merged_download_btn"
                    )
    else:
        st.info("No rows included yet. Tick the 'Include' boxes in the session preview to add records to the merged set.")

else:
    st.info("No searches yet. Fill in the form above and click 'Run Search'.")
