import re
import csv
import os
from html import escape
from difflib import SequenceMatcher
from io import BytesIO, StringIO

import streamlit as st
from ddgs import DDGS
from docx import Document
from pypdf import PdfReader

try:
    from openai import OpenAI
except ImportError:
    OpenAI = None


MIN_LINE_LENGTH = 35
SIMILARITY_THRESHOLD = 0.72
HIGHLIGHT_WORD_LENGTH = 5
MIN_MATCHED_WORDS = 3
REWRITE_MODEL = "gpt-4.1-mini"


THEME = {
    "page": "#f5f7fb",
    "panel": "rgba(255, 255, 255, 0.92)",
    "hero_a": "rgba(37, 99, 235, 0.16)",
    "hero_b": "rgba(244, 114, 182, 0.14)",
    "accent": "#2563eb",
    "accent_dark": "#1d4ed8",
    "danger": "#e11d48",
    "clean": "#059669",
}


def apply_styles():
    css = """
        <style>
            @keyframes fadeUp {
                from {
                    opacity: 0;
                    transform: translateY(14px);
                }
                to {
                    opacity: 1;
                    transform: translateY(0);
                }
            }

            @keyframes glowPulse {
                0% {
                    box-shadow: 0 0 0 rgba(20, 184, 166, 0);
                }
                50% {
                    box-shadow: 0 0 28px rgba(20, 184, 166, 0.28);
                }
                100% {
                    box-shadow: 0 0 0 rgba(20, 184, 166, 0);
                }
            }

            :root {
                color-scheme: light;
            }

            html, body, [data-testid="stAppViewContainer"] {
                background:
                    radial-gradient(circle at 20% 10%, __HERO_A__, transparent 28rem),
                    radial-gradient(circle at 80% 0%, __HERO_B__, transparent 24rem),
                    __PAGE__;
                color: #172033;
            }

            [data-testid="stHeader"] {
                background: rgba(245, 247, 251, 0.72);
                backdrop-filter: blur(16px);
            }

            .block-container {
                padding-top: 2.4rem;
                max-width: 1220px;
            }

            .hero {
                animation: fadeUp 0.5s ease-out;
                background:
                    linear-gradient(135deg, rgba(255, 255, 255, 0.96), rgba(240, 246, 255, 0.92)),
                    linear-gradient(135deg, rgba(37, 99, 235, 0.12), rgba(244, 114, 182, 0.1));
                border: 1px solid rgba(37, 99, 235, 0.14);
                border-radius: 8px;
                box-shadow: 0 24px 70px rgba(37, 99, 235, 0.12);
                margin-bottom: 1.4rem;
                overflow: hidden;
                padding: 1.8rem;
                position: relative;
            }

            .hero:before {
                background: linear-gradient(90deg, __ACCENT__, #f472b6, #14b8a6);
                content: "";
                height: 4px;
                left: 0;
                position: absolute;
                right: 0;
                top: 0;
            }

            .hero h1 {
                color: #111827;
                font-size: 2.45rem;
                letter-spacing: 0;
                margin: 0 0 0.45rem;
            }

            .hero-kicker {
                color: __ACCENT__;
                font-size: 0.78rem;
                font-weight: 800;
                letter-spacing: 0.12em;
                margin-bottom: 0.35rem;
                text-transform: uppercase;
            }

            .help-text {
                color: #445069;
                font-size: 1rem;
                margin: 0;
                max-width: 760px;
            }

            .feature-strip {
                display: flex;
                flex-wrap: wrap;
                gap: 0.55rem;
                margin-top: 1.15rem;
            }

            .feature-chip {
                background: rgba(255, 255, 255, 0.8);
                border: 1px solid rgba(37, 99, 235, 0.15);
                border-radius: 999px;
                color: #1f2a44;
                font-size: 0.86rem;
                font-weight: 700;
                padding: 0.45rem 0.7rem;
            }

            [data-testid="stTabs"] button {
                color: #475569;
                font-weight: 700;
            }

            [data-testid="stTabs"] button[aria-selected="true"] {
                color: __ACCENT__;
            }

            [data-testid="stFileUploader"] {
                background: rgba(255, 255, 255, 0.82);
                border: 1px dashed __ACCENT__;
                border-radius: 8px;
                padding: 0.75rem;
            }

            [data-testid="stTextArea"] textarea {
                background: rgba(255, 255, 255, 0.92);
                border: 1px solid rgba(37, 99, 235, 0.18);
                border-radius: 8px;
                color: #111827;
            }

            .stButton button, .stDownloadButton button {
                animation: glowPulse 3s ease-in-out infinite;
                background: linear-gradient(135deg, __ACCENT__, __ACCENT_DARK__) !important;
                border: 0 !important;
                border-radius: 8px !important;
                color: #ffffff !important;
                font-weight: 800 !important;
                min-height: 3rem;
                transition: transform 0.18s ease, filter 0.18s ease;
            }

            .stButton button:hover, .stDownloadButton button:hover {
                filter: brightness(1.08);
                transform: translateY(-1px);
            }

            [data-testid="stMetric"] {
                animation: fadeUp 0.42s ease-out;
                background: rgba(255, 255, 255, 0.92);
                border: 1px solid rgba(37, 99, 235, 0.12);
                border-radius: 8px;
                padding: 1rem;
            }

            [data-testid="stMetricLabel"] {
                color: #64748b;
            }

            [data-testid="stMetricValue"] {
                color: #111827;
            }

            .result-card {
                animation: fadeUp 0.36s ease-out;
                border: 1px solid rgba(148, 163, 184, 0.24);
                border-radius: 8px;
                margin: 0.7rem 0;
                padding: 1rem;
                background: __PANEL__;
                transition: border-color 0.18s ease, transform 0.18s ease;
            }

            .result-card:hover, .source-card:hover {
                transform: translateY(-1px);
            }

            .plag-card {
                border-color: __DANGER__;
                background: linear-gradient(135deg, rgba(255, 241, 242, 0.98), rgba(255, 255, 255, 0.94));
            }

            .clean-card {
                border-color: __CLEAN__;
                background: linear-gradient(135deg, rgba(236, 253, 245, 0.98), rgba(255, 255, 255, 0.94));
            }

            .line-text {
                border-radius: 6px;
                font-size: 0.97rem;
                line-height: 1.6;
                margin: 0.65rem 0;
                padding: 0.75rem;
            }

            .copied-line {
                background: #fff1f2;
                border-left: 5px solid __DANGER__;
                color: #881337;
                font-weight: 600;
            }

            .clean-line {
                background: #ecfdf5;
                border-left: 5px solid __CLEAN__;
                color: #064e3b;
            }

            .source-card {
                animation: fadeUp 0.34s ease-out;
                background: rgba(255, 255, 255, 0.88);
                border: 1px solid rgba(148, 163, 184, 0.22);
                border-radius: 8px;
                margin-top: 0.75rem;
                padding: 0.85rem;
                transition: border-color 0.18s ease, transform 0.18s ease;
            }

            .source-title {
                color: #111827;
                font-weight: 700;
                margin-bottom: 0.25rem;
            }

            .source-url {
                color: #1d4ed8;
                font-size: 0.9rem;
                overflow-wrap: anywhere;
            }

            .source-snippet {
                color: #334155;
                font-size: 0.95rem;
                line-height: 1.55;
                margin-top: 0.55rem;
            }

            mark {
                background: __DANGER__;
                border-radius: 4px;
                color: #ffffff;
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
                background: __DANGER__;
                color: #ffffff;
            }

            .badge-clean {
                background: __CLEAN__;
                color: #ffffff;
            }

            .score-pill {
                background: #eef2ff;
                border-radius: 999px;
                color: #1e3a8a;
                display: inline-block;
                font-size: 0.82rem;
                font-weight: 600;
                margin-top: 0.35rem;
                padding: 0.25rem 0.55rem;
            }

            .evidence-row {
                display: flex;
                flex-wrap: wrap;
                gap: 0.4rem;
                margin-top: 0.65rem;
            }

            .word-chip {
                background: #fff1f2;
                border: 1px solid __DANGER__;
                border-radius: 999px;
                color: #9f1239;
                display: inline-block;
                font-size: 0.78rem;
                font-weight: 700;
                padding: 0.22rem 0.5rem;
            }

            .source-location {
                color: #64748b;
                font-size: 0.86rem;
                font-weight: 700;
                margin-top: 0.5rem;
            }

            .search-note {
                background: rgba(234, 179, 8, 0.12);
                border: 1px solid rgba(234, 179, 8, 0.3);
                border-radius: 8px;
                color: #92400e;
                margin-top: 0.7rem;
                padding: 0.75rem;
            }

            .control-help, .source-explanation {
                animation: fadeUp 0.36s ease-out;
                background: rgba(255, 255, 255, 0.88);
                border: 1px solid rgba(37, 99, 235, 0.14);
                border-radius: 8px;
                color: #475569;
                font-size: 0.92rem;
                line-height: 1.55;
                margin-top: 0.65rem;
                padding: 0.85rem;
            }

            .source-explanation {
                border-color: rgba(251, 113, 133, 0.34);
            }

            .visible-url {
                color: #1d4ed8;
                font-size: 0.82rem;
                margin-top: 0.2rem;
                overflow-wrap: anywhere;
            }

            .rewrite-note {
                background: #fff7ed;
                border: 1px solid #fdba74;
                border-radius: 8px;
                color: #9a3412;
                margin: 0.8rem 0;
                padding: 0.85rem;
            }

            .guide-panel {
                animation: fadeUp 0.42s ease-out;
                background: rgba(255, 255, 255, 0.9);
                border: 1px solid rgba(37, 99, 235, 0.14);
                border-radius: 8px;
                color: #475569;
                line-height: 1.65;
                padding: 1.2rem;
            }

            section[data-testid="stExpander"] {
                background: rgba(255, 255, 255, 0.76);
                border: 1px solid rgba(148, 163, 184, 0.2);
                border-radius: 8px;
            }
        </style>
    """
    css = (
        css.replace("__PAGE__", THEME["page"])
        .replace("__PANEL__", THEME["panel"])
        .replace("__HERO_A__", THEME["hero_a"])
        .replace("__HERO_B__", THEME["hero_b"])
        .replace("__ACCENT_DARK__", THEME["accent_dark"])
        .replace("__ACCENT__", THEME["accent"])
        .replace("__DANGER__", THEME["danger"])
        .replace("__CLEAN__", THEME["clean"])
    )
    st.markdown(css, unsafe_allow_html=True)


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


def matched_words(line, source_text):
    line_words = important_words(line)
    source_words = important_words(source_text)
    return sorted(line_words.intersection(source_words))


def highlight_text_with_words(text, words):
    words = set(words)

    if not text:
        return "<em>No preview text was returned for this source.</em>"

    def replace_word(match):
        word = match.group(0)
        if word.lower() in words:
            return f"<mark>{escape(word)}</mark>"
        return escape(word)

    parts = re.split(r"([A-Za-z0-9']+)", text)
    return "".join(replace_word(re.match(r"[A-Za-z0-9']+", part)) if re.match(r"[A-Za-z0-9']+", part) else escape(part) for part in parts)


def explain_match(match):
    if match["exact_phrase_found"]:
        return (
            "This line is flagged because the same phrase appears in the source preview. "
            "That is strong evidence of copied or uncited text."
        )

    locations = []
    if match["words_in_title"]:
        locations.append("source title")
    if match["words_in_snippet"]:
        locations.append("source preview")

    location_text = " and ".join(locations) if locations else "the source result"
    return (
        f"This line is flagged because {len(match['matched_words'])} important words from your sentence "
        f"also appear in the {location_text}, and the similarity score is {match['score']}."
    )


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
        words_in_title = matched_words(line, title)
        words_in_snippet = matched_words(line, body)
        evidence_words = sorted(set(words_in_title + words_in_snippet))
        exact_phrase_found = line.lower() in body.lower()

        if exact_phrase_found or (score >= SIMILARITY_THRESHOLD and len(evidence_words) >= MIN_MATCHED_WORDS):
            matches.append(
                {
                    "title": title or "Untitled source",
                    "url": href,
                    "snippet": body,
                    "score": round(score, 2),
                    "matched_words": evidence_words,
                    "words_in_title": words_in_title,
                    "words_in_snippet": words_in_snippet,
                    "exact_phrase_found": exact_phrase_found,
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
            search_error = ""
        except Exception as error:
            matches = []
            search_error = str(error)

        report.append(
            {
                "line_number": line_number,
                "line": line,
                "matches": matches,
                "search_error": search_error,
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
    writer.writerow(
        [
            "line_number",
            "status",
            "line",
            "matched_words",
            "found_in_title",
            "found_in_snippet",
            "source_title",
            "source_url",
            "similarity_score",
            "snippet",
            "search_note",
        ]
    )

    for item in report:
        if not item["matches"]:
            writer.writerow([item["line_number"], "No match found", item["line"], "", "", "", "", "", "", "", item.get("search_error", "")])
            continue

        for match in item["matches"]:
            writer.writerow(
                [
                    item["line_number"],
                    "Possible plagiarism",
                    item["line"],
                    ", ".join(match["matched_words"]),
                    ", ".join(match["words_in_title"]),
                    ", ".join(match["words_in_snippet"]),
                    match["title"],
                    match["url"],
                    match["score"],
                    match["snippet"],
                    item.get("search_error", ""),
                ]
            )

    return output.getvalue()


def get_openai_api_key():
    try:
        secret_key = st.secrets.get("OPENAI_API_KEY", "")
    except Exception:
        secret_key = ""

    return secret_key or os.environ.get("OPENAI_API_KEY", "")


def rewrite_line_with_ai(line, matches):
    api_key = get_openai_api_key()
    if not api_key or OpenAI is None:
        return ""

    source_context = "\n".join(
        f"- {match['title']}: {match['url']}\n  Snippet: {match['snippet']}"
        for match in matches[:2]
    )
    prompt = f"""
Rewrite the student's sentence in original wording while preserving the meaning.
Do not copy the source wording. Keep it clear, academic, and concise.
Do not add quotation marks. Return only the rewritten sentence.

Student sentence:
{line}

Possible source evidence:
{source_context}
"""

    try:
        client = OpenAI(api_key=api_key)
        response = client.responses.create(
            model=REWRITE_MODEL,
            input=prompt,
        )
        return response.output_text.strip()
    except Exception:
        return ""


def build_corrected_docx(original_text, report):
    doc = Document()
    doc.add_heading("Corrected Plagiarism Review Document", level=1)

    matched_items = [item for item in report if item["matches"]]
    api_enabled = bool(get_openai_api_key() and OpenAI is not None)

    doc.add_paragraph(f"Lines checked: {len(report)}")
    doc.add_paragraph(f"Lines with possible plagiarism: {len(matched_items)}")
    doc.add_paragraph(f"Possible plagiarism score: {report_score(report)}%")

    if api_enabled:
        doc.add_paragraph(
            "Flagged lines were rewritten with AI assistance where possible. Review the changes before final submission."
        )
    else:
        doc.add_paragraph(
            "AI rewriting is not enabled because OPENAI_API_KEY is not configured. Flagged lines are marked for rewriting and include source links."
        )

    doc.add_heading("Corrected / Reviewed Text", level=2)

    report_by_line = {item["line_number"]: item for item in report}

    for line_number, raw_line in enumerate(original_text.splitlines(), start=1):
        item = report_by_line.get(line_number)
        if not raw_line.strip():
            doc.add_paragraph("")
            continue

        if item and item["matches"]:
            rewritten = rewrite_line_with_ai(item["line"], item["matches"])
            paragraph = doc.add_paragraph()
            paragraph.add_run("Revised line: ").bold = True
            paragraph.add_run(rewritten or item["line"])

            if not rewritten:
                note = doc.add_paragraph()
                note.add_run("Action needed: ").bold = True
                note.add_run("Rewrite this line in your own words or add a citation before final use.")

            source_note = doc.add_paragraph()
            source_note.add_run("Possible source(s): ").bold = True
            source_note.add_run(", ".join(match["url"] for match in item["matches"][:3] if match["url"]))
        else:
            doc.add_paragraph(raw_line)

    doc.add_page_break()
    doc.add_heading("Plagiarism Review Notes", level=2)

    if not matched_items:
        doc.add_paragraph("No possible source matches were found for the checked lines.")
    else:
        for item in matched_items:
            doc.add_heading(f"Line {item['line_number']}", level=3)
            doc.add_paragraph(item["line"])
            for match in item["matches"][:3]:
                doc.add_paragraph(f"Source: {match['title']}")
                doc.add_paragraph(f"URL: {match['url']}")
                doc.add_paragraph(f"Matched words: {', '.join(match['matched_words'])}")
                doc.add_paragraph(f"Similarity score: {match['score']}")

    output = BytesIO()
    doc.save(output)
    return output.getvalue()


def render_report(report, original_text):
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

    api_enabled = bool(get_openai_api_key() and OpenAI is not None)
    if api_enabled:
        st.markdown(
            """
            <div class="rewrite-note">
                AI rewrite export is enabled. The corrected DOCX will rewrite flagged lines and include source notes.
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            """
            <div class="rewrite-note">
                AI rewrite export is not enabled on this deployment. The DOCX will still mark flagged lines,
                include source links, and show what needs rewriting. Add <strong>OPENAI_API_KEY</strong> in
                Streamlit secrets to enable automatic rewriting.
            </div>
            """,
            unsafe_allow_html=True,
        )

    if st.button("Prepare corrected DOCX", use_container_width=True):
        with st.spinner("Building corrected document..."):
            st.session_state["corrected_docx"] = build_corrected_docx(original_text, report)

    if "corrected_docx" in st.session_state:
        st.download_button(
            "Download corrected DOCX",
            data=st.session_state["corrected_docx"],
            file_name="corrected_plagiarism_review.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
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
            words_for_line = sorted({word for match in item["matches"] for word in match.get("matched_words", [])})
            line_html = highlight_text_with_words(item["line"], words_for_line) if found else escape(item["line"])

            st.markdown(
                f"""
                <div class="result-card {card_class}">
                    <span class="badge {badge_class}">{escape(label)}</span>
                    <div class="line-text {line_class}">
                        Line {item['line_number']}: {line_html}
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            if not found:
                st.success("This line did not return a likely web source match.")
                if item.get("search_error"):
                    st.markdown(
                        f"""
                        <div class="search-note">
                            Search note: {escape(item['search_error'])}
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                continue

            for match in item["matches"]:
                visible_url = escape(match["url"]) if match["url"] else "No source URL available"
                source_link = (
                    f'<a class="source-url" href="{visible_url}" target="_blank">Open source link</a>'
                    if match["url"]
                    else '<span class="source-url">No source URL available</span>'
                )
                highlighted_title = highlight_text_with_words(match["title"], match["words_in_title"])
                highlighted_snippet = highlight_text_with_words(match["snippet"], match["words_in_snippet"])
                word_chips = "".join(f'<span class="word-chip">{escape(word)}</span>' for word in match["matched_words"])
                exact_label = "Exact phrase found in snippet" if match["exact_phrase_found"] else "Similar words found in source"
                explanation = explain_match(match)

                st.markdown(
                    f"""
                    <div class="source-card">
                        <div class="source-title">{highlighted_title}</div>
                        {source_link}
                        <div class="visible-url">{visible_url}</div>
                        <div class="score-pill">Similarity score: {match['score']}</div>
                        <div class="score-pill">{escape(exact_label)}</div>
                        <div class="source-explanation">
                            <strong>Why this may be plagiarism:</strong><br>
                            {escape(explanation)}
                        </div>
                        <div class="source-location">
                            Matched words from your sentence:
                        </div>
                        <div class="evidence-row">{word_chips}</div>
                        <div class="source-location">
                            Found in source preview:
                        </div>
                        <div class="source-snippet">{highlighted_snippet}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )


def main():
    st.set_page_config(page_title="Plagiarism Source Finder", page_icon="Search", layout="wide")
    apply_styles()

    st.markdown(
        """
        <div class="hero">
            <div class="hero-kicker">Source Evidence Review</div>
            <h1>Plagiarism Source Finder</h1>
            <p class="help-text">
                Upload a file or paste text to find possible copied lines, see the matching source links,
                and understand exactly which words triggered each result.
            </p>
            <div class="feature-strip">
                <span class="feature-chip">Checks every line</span>
                <span class="feature-chip">Shows source URLs</span>
                <span class="feature-chip">Highlights matched words</span>
                <span class="feature-chip">Explains each flag</span>
                <span class="feature-chip">Exports corrected DOCX</span>
                <span class="feature-chip">Exports CSV report</span>
            </div>
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
            st.markdown(
                """
                <div class="control-help">
                    <strong>What does this mean?</strong><br>
                    For every line in your text, the app checks this many web search results.
                    <br><br>
                    <strong>1</strong> = faster but fewer sources checked<br>
                    <strong>3</strong> = best default for normal checking<br>
                    <strong>5</strong> = slower but checks more possible sources
                </div>
                """,
                unsafe_allow_html=True,
            )

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
            render_report(report, text)

    with guidance_tab:
        st.markdown(
            """
            <div class="guide-panel">
                <strong>What is plagiarism?</strong>
                Plagiarism means using someone else's words, ideas, or structure without giving proper credit.
                In this app, a red result means the line may need a citation, quotation marks, rewriting, or review.
                <br><br>
                <strong>Red lines</strong> mean the checker found a possible source match.
                <br><br>
                <strong>Green lines</strong> mean no likely web source was found for that line.
                <br><br>
                <strong>Open source link</strong> takes you to the page where matching text or matching words were found.
                The full URL is also shown below the link.
                <br><br>
                <strong>Highlighted source words</strong> show which words from your line also appeared
                in the source title or preview.
                <br><br>
                <strong>Why this may be plagiarism</strong> explains the evidence for each result, including exact phrase
                matches, matched words, and similarity score.
                <br><br>
                <strong>Search notes</strong> are not counted as plagiarism. They only mean the search provider
                returned no result or had a temporary lookup issue.
                <br><br>
                <strong>Similarity score</strong> is a rough match score from <code>0</code> to <code>1</code>.
                A higher score means the source preview is closer to the checked line.
                <br><br>
                <strong>Corrected DOCX</strong> creates a new Word document. If an OpenAI API key is configured
                in Streamlit secrets, flagged lines are rewritten. If not, the document marks each flagged line
                and includes source links so you can rewrite it manually.
            </div>
            """,
            unsafe_allow_html=True,
        )


if __name__ == "__main__":
    main()
