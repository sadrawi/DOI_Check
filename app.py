"""
Streamlit front end for the Crossref DOI checker.

Run locally:   streamlit run app.py
Deploy:        push to GitHub, then share.streamlit.io -> New app
"""

import streamlit as st
import pandas as pd

import doi_core as core


st.set_page_config(page_title="Reference DOI checker", page_icon="check", layout="wide")


# ----------------------------------------------------------------------
# Cached lookups - Streamlit reruns the whole script on every interaction,
# so without this each button click would hit Crossref again.
# ----------------------------------------------------------------------

@st.cache_data(show_spinner=False, ttl=3600)
def cached_check(reference):
    return core.check_reference(reference)


def check_with_cache(references, progress_bar, status_line):
    rows = []
    for i, entry in enumerate(references):
        status_line.write("Checking " + str(i + 1) + " of " + str(len(references)) + ": " + entry[0][:80] + "...")
        row = dict(cached_check(entry[0]))
        row["n"] = i + 1
        row["paragraphs"] = entry[1]
        rows.append(row)
        progress_bar.progress((i + 1) / len(references))
    status_line.empty()
    return pd.DataFrame(rows)


# ----------------------------------------------------------------------
# Layout
# ----------------------------------------------------------------------

st.title("Reference DOI checker")
st.caption("Supported by Claude Opus 5.0 (Anthropic PBC, San Francisco, California, U.S.")


tab_docx, tab_paste = st.tabs(["Word document", "Paste a list"])

with tab_docx:
    uploaded = st.file_uploader("Manuscript (.docx)", type="docx")
    st.caption("The reference section is found by heading: "
               + ", ".join(core.REF_HEADINGS[:4]) + ", ...")
    run_docx = st.button("Check references", type="primary", disabled=uploaded is None)

with tab_paste:
    pasted = st.text_area("One reference per line", height=200)
    run_paste = st.button("Check list", type="primary", disabled=pasted.strip() == "")


def render(results, docx_bytes, filename):
    counts = {}
    for status in results["status"]:
        key = status.split(" - ")[0]
        counts[key] = counts.get(key, 0) + 1

    columns = st.columns(4)
    for i, key in enumerate(["OK", "WRONG PAPER", "DOI DOES NOT EXIST", "NO DOI"]):
        columns[i].metric(key.title(), counts.get(key, 0))

    problems = results[results["status"] != "OK"]
    if len(problems) == 0:
        st.success("Every DOI resolves to the paper cited.")
    else:
        st.warning(str(len(problems)) + " entries need attention.")

    table = results[["n", "status", "given_doi"]]
    st.dataframe(table, use_container_width=True, hide_index=True)

    with st.expander("Full detail"):
        for _, row in results.iterrows():
            # st.markdown("**[" + str(row["n"]) + "]**")
            st.markdown("**[" + str(row["n"]) + "] " + row["reference"][:300] + "   [" + row["status"] + "]**")
            # if row["resolves_to"] != "":
            #     st.caption("Given DOI points to: " + row["resolves_to"])
            if row["correct_doi"] != "":
                st.caption("https://doi.org/" + row["correct_doi"])
            st.markdown(" ")

    # left, right = st.columns(2)
    # left.download_button(
    #     "Download results CSV",
    #     results.drop(columns=["paragraphs"]).to_csv(index=False).encode("utf-8"),
    #     file_name="doi_check_results.csv",
    #     mime="text/csv",
    # )
    # if docx_bytes is not None:
    #     right.download_button(
    #         "Download marked document",
    #         docx_bytes,
    #         file_name=filename,
    #         mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    #     )


if run_docx and uploaded is not None:
    try:
        document, references = core.read_references_from_docx(uploaded)
    except ValueError as error:
        st.error(str(error))
        st.stop()

    st.info("Found " + str(len(references)) + " references.")
    bar = st.progress(0.0)
    line = st.empty()
    results = check_with_cache(references, bar, line)
    bar.empty()

    docx_bytes = core.mark_docx(document, results)
    stem = uploaded.name.rsplit(".", 1)[0]
    render(results, docx_bytes, stem + "_checked.docx")

if run_paste and pasted.strip() != "":
    references = core.load_references_from_text(pasted)
    st.info("Checking " + str(len(references)) + " references.")
    bar = st.progress(0.0)
    line = st.empty()
    results = check_with_cache(references, bar, line)
    bar.empty()
    render(results, None, "")
