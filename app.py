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
MIN_MATCHED_WORDS = 3


THEMES = {
    "Neon Evidence": {
        "page": "#070b12",
        "panel": "rgba(15, 23, 42, 0.86)",
        "hero_a": "rgba(20, 184, 166, 0.18)",
        "hero_b": "rgba(220, 38, 38, 0.14)",
        "accent": "#14b8a6",
        "accent_dark": "#0f766e",
        "danger": "#fb7185",
        "clean": "#2dd4bf",
    },
    "Midnight Violet": {
        "page": "#090816",
        "panel": "rgba(24, 18, 43, 0.88)",
        "hero_a": "rgba(168, 85, 247, 0.18)",
        "hero_b": "rgba(14, 165, 233, 0.14)",
        "accent": "#a78bfa",
        "accent_dark": "#7c3aed",
        "danger": "#f43f5e",
        "clean": "#38bdf8",
    },
    "Emerald Focus": {
        "page": "#06110d",
        "panel": "rgba(6, 30, 24, 0.88)",
        "hero_a": "rgba(16, 185, 129, 0.2)",
        "hero_b": "rgba(245, 158, 11, 0.13)",
        "accent": "#10b981",
        "accent_dark": "#047857",
        "danger": "#f97316",
        "clean": "#34d399",
    },
}


def apply_styles(theme):
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
                color-scheme: dark;
            }

            html, body, [data-testid="stAppViewContainer"] {
                background:
                    radial-gradient(circle at 20% 10%, __HERO_A__, transparent 28rem),
                    radial-gradient(circle at 80% 0%, __HERO_B__, transparent 24rem),
                    __PAGE__;
                color: #e5e7eb;
            }

            [data-testid="stHeader"] {
                background: rgba(7, 11, 18, 0.72);
                backdrop-filter: blur(16px);
            }

            .block-container {
                padding-top: 2.4rem;
                max-width: 1220px;
            }

            .hero {
                animation: fadeUp 0.5s ease-out;
                background:
                    linear-gradient(135deg, rgba(15, 23, 42, 0.96), rgba(17, 24, 39, 0.88)),
                    linear-gradient(135deg, rgba(20, 184, 166, 0.2), rgba(248, 113, 113, 0.14));
                border: 1px solid rgba(148, 163, 184, 0.22);
                border-radius: 8px;
                box-shadow: 0 28px 80px rgba(0, 0, 0, 0.36);
                margin-bottom: 1.4rem;
                overflow: hidden;
                padding: 1.8rem;
                position: relative;
            }

            .hero:before {
                background: linear-gradient(90deg, __ACCENT__, #f97316, __DANGER__);
                content: "";
                height: 4px;
                left: 0;
                position: absolute;
                right: 0;
                top: 0;
            }

            .hero h1 {
                color: #f8fafc;
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
                color: #cbd5e1;
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
                background: rgba(15, 23, 42, 0.82);
                border: 1px solid rgba(148, 163, 184, 0.25);
                border-radius: 999px;
                color: #e2e8f0;
                font-size: 0.86rem;
                font-weight: 700;
                padding: 0.45rem 0.7rem;
            }

            [data-testid="stTabs"] button {
                color: #cbd5e1;
                font-weight: 700;
            }

            [data-testid="stTabs"] button[aria-selected="true"] {
                color: __ACCENT__;
            }

            [data-testid="stFileUploader"] {
                background: rgba(15, 23, 42, 0.72);
                border: 1px dashed __ACCENT__;
                border-radius: 8px;
                padding: 0.75rem;
            }

            [data-testid="stTextArea"] textarea {
                background: rgba(15, 23, 42, 0.86);
                border: 1px solid rgba(148, 163, 184, 0.26);
                border-radius: 8px;
                color: #f8fafc;
            }

            .stButton button, .stDownloadButton button {
                animation: glowPulse 3s ease-in-out infinite;
                background: linear-gradient(135deg, __ACCENT__, __ACCENT_DARK__) !important;
                border: 0 !important;
                border-radius: 8px !important;
                color: #f8fafc !important;
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
                background: rgba(15, 23, 42, 0.82);
                border: 1px solid rgba(148, 163, 184, 0.2);
                border-radius: 8px;
                padding: 1rem;
            }

            [data-testid="stMetricLabel"] {
                color: #94a3b8;
            }

            [data-testid="stMetricValue"] {
                color: #f8fafc;
            }

            .result-card {
                animation: fadeUp 0.36s ease-out;
                border: 1px solid rgba(148, 163, 184, 0.22);
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
                background: linear-gradient(135deg, rgba(127, 29, 29, 0.42), rgba(15, 23, 42, 0.9));
            }

            .clean-card {
                border-color: __CLEAN__;
                background: linear-gradient(135deg, rgba(20, 83, 45, 0.36), rgba(15, 23, 42, 0.9));
            }

            .line-text {
                border-radius: 6px;
                font-size: 0.97rem;
                line-height: 1.6;
                margin: 0.65rem 0;
                padding: 0.75rem;
            }

            .copied-line {
                background: rgba(127, 29, 29, 0.78);
                border-left: 5px solid __DANGER__;
                color: #fee2e2;
                font-weight: 600;
            }

            .clean-line {
                background: rgba(20, 83, 45, 0.76);
                border-left: 5px solid __CLEAN__;
                color: #dcfce7;
            }

            .source-card {
                animation: fadeUp 0.34s ease-out;
                background: rgba(2, 6, 23, 0.6);
                border: 1px solid rgba(148, 163, 184, 0.18);
                border-radius: 8px;
                margin-top: 0.75rem;
                padding: 0.85rem;
                transition: border-color 0.18s ease, transform 0.18s ease;
            }

            .source-title {
                color: #f8fafc;
                font-weight: 700;
                margin-bottom: 0.25rem;
            }

            .source-url {
                color: #67e8f9;
                font-size: 0.9rem;
                overflow-wrap: anywhere;
            }

            .source-snippet {
                color: #cbd5e1;
                font-size: 0.95rem;
                line-height: 1.55;
                margin-top: 0.55rem;
            }

            mark {
                background: __DANGER__;
                border-radius: 4px;
                color: #450a0a;
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
                color: #450a0a;
            }

            .badge-clean {
                background: __CLEAN__;
                color: #042f2e;
            }

            .score-pill {
                background: rgba(148, 163, 184, 0.16);
                border-radius: 999px;
                color: #e2e8f0;
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
                background: rgba(251, 113, 133, 0.18);
                border: 1px solid __DANGER__;
                border-radius: 999px;
                color: #fecdd3;
                display: inline-block;
                font-size: 0.78rem;
                font-weight: 700;
                padding: 0.22rem 0.5rem;
            }

            .source-location {
                color: #94a3b8;
                font-size: 0.86rem;
                font-weight: 700;
                margin-top: 0.5rem;
            }

            .search-note {
                background: rgba(234, 179, 8, 0.12);
                border: 1px solid rgba(234, 179, 8, 0.3);
                border-radius: 8px;
                color: #fde68a;
                margin-top: 0.7rem;
                padding: 0.75rem;
            }

            .control-help, .source-explanation {
                animation: fadeUp 0.36s ease-out;
                background: rgba(15, 23, 42, 0.74);
                border: 1px solid rgba(148, 163, 184, 0.2);
                border-radius: 8px;
                color: #cbd5e1;
                font-size: 0.92rem;
                line-height: 1.55;
                margin-top: 0.65rem;
                padding: 0.85rem;
            }

            .source-explanation {
                border-color: rgba(251, 113, 133, 0.34);
            }

            .visible-url {
                color: #93c5fd;
                font-size: 0.82rem;
                margin-top: 0.2rem;
                overflow-wrap: anywhere;
            }

            .guide-panel {
                animation: fadeUp 0.42s ease-out;
                background: rgba(15, 23, 42, 0.78);
                border: 1px solid rgba(148, 163, 184, 0.22);
                border-radius: 8px;
                color: #cbd5e1;
                line-height: 1.65;
                padding: 1.2rem;
            }

            section[data-testid="stExpander"] {
                background: rgba(15, 23, 42, 0.5);
                border: 1px solid rgba(148, 163, 184, 0.16);
                border-radius: 8px;
            }
        </style>
    """
    css = (
        css.replace("__PAGE__", theme["page"])
        .replace("__PANEL__", theme["panel"])
        .replace("__HERO_A__", theme["hero_a"])
        .replace("__HERO_B__", theme["hero_b"])
        .replace("__ACCENT_DARK__", theme["accent_dark"])
        .replace("__ACCENT__", theme["accent"])
        .replace("__DANGER__", theme["danger"])
        .replace("__CLEAN__", theme["clean"])
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
    st.set_page_config(page_title="AI Plagiarism Checker", page_icon="Search", layout="wide")
    theme_name = st.sidebar.selectbox("Theme", list(THEMES.keys()))
    apply_styles(THEMES[theme_name])

    st.sidebar.markdown(
        """
        <div class="control-help">
            Choose a visual theme for screenshots, demos, or LinkedIn project presentation.
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="hero">
            <div class="hero-kicker">Document Integrity Scanner</div>
            <h1>AI Plagiarism Checker</h1>
            <p class="help-text">
                A dark-mode review dashboard for checking pasted text or uploaded files.
                Suspicious lines glow red, clean lines stay green, and matching source words are highlighted for fast review.
            </p>
            <div class="feature-strip">
                <span class="feature-chip">Line-by-line scan</span>
                <span class="feature-chip">Source links</span>
                <span class="feature-chip">Red match highlights</span>
                <span class="feature-chip">Match explanations</span>
                <span class="feature-chip">Theme picker</span>
                <span class="feature-chip">CSV export</span>
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
            render_report(report)

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
            </div>
            """,
            unsafe_allow_html=True,
        )


if __name__ == "__main__":
    main()
