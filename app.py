import re
import csv
from html import escape
from difflib import SequenceMatcher
from io import BytesIO, StringIO

import streamlit as st
from ddgs import DDGS
from docx import Document
from pypdf import PdfReader


MIN_LINE_LENGTH = 35
SIMILARITY_THRESHOLD = 0.72
HIGHLIGHT_WORD_LENGTH = 5


def apply_styles():
    st.markdown(
        """
        <style>
            .block-container {
                padding-top: 2rem;
                max-width: 1180px;
            }

            .hero {
                border-bottom: 1px solid #e5e7eb;
                margin-bottom: 1.25rem;
                padding-bottom: 1rem;
            }

            .hero h1 {
                margin-bottom: 0.25rem;
            }

            .help-text {
                color: #4b5563;
                font-size: 1rem;
                margin: 0;
            }

            .result-card {
                border: 1px solid #e5e7eb;
                border-radius: 8px;
                margin: 0.7rem 0;
                padding: 1rem;
                background: #ffffff;
            }

            .plag-card {
                border-color: #fecaca;
                background: #fff7f7;
            }

            .clean-card {
                border-color: #bbf7d0;
                background: #f7fff9;
            }

            .line-text {
                border-radius: 6px;
                font-size: 0.97rem;
                line-height: 1.6;
                margin: 0.65rem 0;
                padding: 0.75rem;
            }

            .copied-line {
                background: #fee2e2;
                border-left: 5px solid #dc2626;
                color: #7f1d1d;
                font-weight: 600;
            }

            .clean-line {
                background: #dcfce7;
                border-left: 5px solid #16a34a;
                color: #14532d;
            }

            .source-card {
                background: #ffffff;
                border: 1px solid #e5e7eb;
                border-radius: 8px;
                margin-top: 0.75rem;
                padding: 0.85rem;
            }

            .source-title {
                color: #111827;
                font-weight: 700;
                margin-bottom: 0.25rem;
            }

            .source-url {
                color: #2563eb;
                font-size: 0.9rem;
                overflow-wrap: anywhere;
            }

            .source-snippet {
                color: #374151;
                font-size: 0.95rem;
                line-height: 1.55;
                margin-top: 0.55rem;
            }

            mark {
                background: #fecaca;
                border-radius: 4px;
                color: #7f1d1d;
                font-weight: 700;
                padding: 0 0.15rem;
            }

            .badge {
                border-radius: 999px;
                display: inline-block;
                font-size: 0.78rem;
                font-weight: 700;
                letter-spacing: 0.01em;
                padding: 0.25rem 0.6rem;
            }

            .badge-danger {
                background: #dc2626;
                color: #ffffff;
            }

            .badge-clean {
                background: #16a34a;
                color: #ffffff;
            }

            .score-pill {
                background: #f3f4f6;
                border-radius: 999px;
                color: #374151;
                display: inline-block;
                font-size: 0.82rem;
                font-weight: 600;
                margin-top: 0.35rem;
                padding: 0.25rem 0.55rem;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )


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


def important_words(text):
    words = re.findall(r"[A-Za-z0-9']+", text.lower())
    return {
        word
        for word in words
        if len(word) >= HIGHLIGHT_WORD_LENGTH
    }


def highlight_source_snippet(line, snippet):
    words = important_words(line)

    if not snippet:
        return "<em>No preview text was returned for this source.</em>"

    def replace_word(match):
        word = match.group(0)
        if word.lower() in words:
            return f"<mark>{escape(word)}</mark>"
        return escape(word)

    parts = re.split(r"([A-Za-z0-9']+)", snippet)
    return "".join(replace_word(re.match(r"[A-Za-z0-9']+", part)) if re.match(r"[A-Za-z0-9']+", part) else escape(part) for part in parts)


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

    st.subheader("Review Summary")

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

    if score >= 50:
        st.error("High number of possible matches found. Review the red lines and source links carefully.")
    elif score > 0:
        st.warning("Some possible matches were found. Check the highlighted lines before submitting your work.")
    else:
        st.success("No possible source matches were found for the checked lines.")

    st.subheader("Line-by-line Review")

    if not report:
        st.info("No long enough lines found to check.")
        return

    for item in report:
        found = bool(item["matches"])
        label = "Possible plagiarism found" if found else "No match found"
        card_class = "plag-card" if found else "clean-card"
        line_class = "copied-line" if found else "clean-line"
        badge_class = "badge-danger" if found else "badge-clean"

        with st.expander(f"Line {item['line_number']}: {label}", expanded=found):
            st.markdown(
                f"""
                <div class="result-card {card_class}">
                    <span class="badge {badge_class}">{escape(label)}</span>
                    <div class="line-text {line_class}">
                        Line {item['line_number']}: {escape(item['line'])}
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            if not found:
                st.success("This line did not return a likely web source match.")
                continue

            for match in item["matches"]:
                source_link = (
                    f'<a class="source-url" href="{escape(match["url"])}" target="_blank">Open source</a>'
                    if match["url"]
                    else '<span class="source-url">No source URL available</span>'
                )
                highlighted_snippet = highlight_source_snippet(item["line"], match["snippet"])

                st.markdown(
                    f"""
                    <div class="source-card">
                        <div class="source-title">{escape(match['title'])}</div>
                        {source_link}
                        <div class="score-pill">Similarity score: {match['score']}</div>
                        <div class="source-snippet">{highlighted_snippet}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )


def main():
    st.set_page_config(page_title="AI Plagiarism Checker", page_icon="Search", layout="wide")
    apply_styles()

    st.markdown(
        """
        <div class="hero">
            <h1>AI Plagiarism Checker</h1>
            <p class="help-text">
                Upload a file or paste text. The app highlights suspicious lines in red,
                shows source links, and marks matching words inside the source preview.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    input_tab, guidance_tab = st.tabs(["Check content", "How to read results"])

    with input_tab:
        left, right = st.columns([1, 1])

        with left:
            uploaded_file = st.file_uploader("Upload a file", type=["txt", "pdf", "docx"])

        with right:
            max_results = st.slider("Search results per line", min_value=1, max_value=5, value=3)

        pasted_text = st.text_area("Or paste text here", height=220)

        if st.button("Check plagiarism", type="primary", use_container_width=True):
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

    with guidance_tab:
        st.markdown(
            """
            **Red lines** mean the checker found a possible source match.

            **Green lines** mean no likely web source was found for that line.

            **Highlighted source words** show which words from your line also appeared
            in the source preview.

            **Similarity score** is a rough match score from `0` to `1`. A higher score
            means the source preview is closer to the checked line.
            """
        )


if __name__ == "__main__":
    main()
