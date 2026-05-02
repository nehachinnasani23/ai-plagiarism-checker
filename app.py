import re
import csv
from difflib import SequenceMatcher
from io import BytesIO, StringIO

import streamlit as st
from ddgs import DDGS
from docx import Document
from pypdf import PdfReader


MIN_LINE_LENGTH = 35
SIMILARITY_THRESHOLD = 0.72


def clean_line(line):
    return re.sub(r"\s+", " ", line).strip()


def split_lines(text):
    lines = []

    for line_number, raw_line in enumerate(text.splitlines(), start=1):
        line = clean_line(raw_line)
        if len(line) >= MIN_LINE_LENGTH:
            lines.append((line_number, line))

    return lines


def read_txt(file):
    return file.read().decode("utf-8", errors="ignore")


def read_pdf(file):
    reader = PdfReader(BytesIO(file.read()))
    pages = []

    for page in reader.pages:
        pages.append(page.extract_text() or "")

    return "\n".join(pages)


def read_docx(file):
    document = Document(BytesIO(file.read()))
    return "\n".join(paragraph.text for paragraph in document.paragraphs)


def read_uploaded_file(file):
    name = file.name.lower()

    if name.endswith(".txt"):
        return read_txt(file)
    if name.endswith(".pdf"):
        return read_pdf(file)
    if name.endswith(".docx"):
        return read_docx(file)

    raise ValueError("Please upload a .txt, .pdf, or .docx file.")


def similarity(a, b):
    return SequenceMatcher(None, a.lower(), b.lower()).ratio()


def search_line(line, max_results):
    query = f'"{line}"'

    with DDGS() as ddgs:
        results = list(ddgs.text(query, max_results=max_results))

    matches = []
    for result in results:
        title = result.get("title", "")
        href = result.get("href", "")
        body = result.get("body", "")
        score = max(similarity(line, title), similarity(line, body))

        if line.lower() in body.lower() or score >= SIMILARITY_THRESHOLD:
            matches.append(
                {
                    "title": title or "Untitled source",
                    "url": href,
                    "snippet": body,
                    "score": round(score, 2),
                }
            )

    return matches


def check_plagiarism(text, max_results):
    checked_lines = split_lines(text)
    report = []

    progress = st.progress(0)
    status = st.empty()

    for index, (line_number, line) in enumerate(checked_lines, start=1):
        status.write(f"Checking line {line_number}...")

        try:
            matches = search_line(line, max_results=max_results)
        except Exception as error:
            matches = [
                {
                    "title": "Search failed",
                    "url": "",
                    "snippet": str(error),
                    "score": 0,
                }
            ]

        report.append(
            {
                "line_number": line_number,
                "line": line,
                "matches": matches,
            }
        )

        progress.progress(index / len(checked_lines))

    status.empty()
    progress.empty()
    return report


def report_score(report):
    if not report:
        return 0

    matched_lines = len([item for item in report if item["matches"]])
    return round((matched_lines / len(report)) * 100)


def build_csv_report(report):
    output = StringIO()
    writer = csv.writer(output)
    writer.writerow(["line_number", "status", "line", "source_title", "source_url", "similarity_score", "snippet"])

    for item in report:
        if not item["matches"]:
            writer.writerow([item["line_number"], "No match found", item["line"], "", "", "", ""])
            continue

        for match in item["matches"]:
            writer.writerow(
                [
                    item["line_number"],
                    "Possible plagiarism",
                    item["line"],
                    match["title"],
                    match["url"],
                    match["score"],
                    match["snippet"],
                ]
            )

    return output.getvalue()


def render_report(report):
    plagiarized = [item for item in report if item["matches"]]
    score = report_score(report)

    st.subheader("Summary")

    col1, col2, col3 = st.columns(3)
    col1.metric("Lines checked", len(report))
    col2.metric("Lines with possible matches", len(plagiarized))
    col3.metric("Possible plagiarism score", f"{score}%")

    st.download_button(
        "Download CSV report",
        data=build_csv_report(report),
        file_name="plagiarism_report.csv",
        mime="text/csv",
    )

    st.subheader("Line-by-line results")

    if not report:
        st.info("No long enough lines found to check.")
        return

    for item in report:
        found = bool(item["matches"])
        label = "Possible plagiarism" if found else "No match found"

        with st.expander(f"Line {item['line_number']}: {label}", expanded=found):
            st.write(item["line"])

            if not found:
                st.success("No likely source found for this line.")
                continue

            for match in item["matches"]:
                st.markdown(f"**{match['title']}**")
                if match["url"]:
                    st.markdown(f"[Open source]({match['url']})")
                st.caption(f"Similarity score: {match['score']}")
                st.write(match["snippet"])
                st.divider()


def main():
    st.set_page_config(page_title="AI Plagiarism Checker", page_icon="Search", layout="wide")
    st.title("AI Plagiarism Checker")

    st.write(
        "Upload a file or paste text to scan each line for possible copied content and source links."
    )

    uploaded_file = st.file_uploader("Upload a file", type=["txt", "pdf", "docx"])
    pasted_text = st.text_area("Or paste text here", height=220)

    max_results = st.slider("Search results per line", min_value=1, max_value=5, value=3)

    if st.button("Check plagiarism", type="primary"):
        text = pasted_text.strip()

        if uploaded_file is not None:
            try:
                text = read_uploaded_file(uploaded_file)
            except Exception as error:
                st.error(error)
                return

        if not text.strip():
            st.warning("Please upload a file or paste some text first.")
            return

        report = check_plagiarism(text, max_results=max_results)
        render_report(report)


if __name__ == "__main__":
    main()
