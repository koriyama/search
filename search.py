# ... (the rest of the code is unchanged until the download section) ...

# ---- Download button for merged set ----
col_name, col_btn = st.columns([2, 1])
with col_name:
    # Build a summary of search terms from the first session's phrases
    if st.session_state.search_sessions:
        first_phrases = st.session_state.search_sessions[0]["phrases"]
        summary = "_".join(first_phrases)[:50].replace(" ", "_").replace('"', '')  # clean
        default_filename = f"{datetime.now().strftime('%Y-%m-%d')}_{summary}.xlsx"
    else:
        default_filename = f"{datetime.now().strftime('%Y-%m-%d')}_search_results.xlsx"
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
            # ... (rest of download logic) ...
