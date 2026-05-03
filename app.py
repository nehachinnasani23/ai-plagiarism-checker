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


SCAN_MODES = {
    "Fast precheck": {"max_results": 1, "min_length": 55, "max_lines": 35},
    "Publication review": {"max_results": 3, "min_length": 40, "max_lines": 90},
    "Deep editorial scan": {"max_results": 5, "min_length": 35, "max_lines": 160},
}


THEME = {
    "page": "#080b13",
    "panel": "rgba(17, 24, 39, 0.92)",
    "hero_a": "rgba(196, 164, 94, 0.18)",
    "hero_b": "rgba(217, 79, 69, 0.16)",
    "accent": "#d6b76a",
    "accent_dark": "#8f6f2a",
    "danger": "#ff6b5f",
    "clean": "#2dd4bf",
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
                color-scheme: dark;
            }

            html, body, [data-testid="stAppViewContainer"] {
                background:
                    radial-gradient(circle at 20% 10%, __HERO_A__, transparent 28rem),
                    radial-gradient(circle at 80% 0%, __HERO_B__, transparent 24rem),
                    linear-gradient(180deg, rgba(16, 24, 39, 0.7), transparent 18rem),
                    __PAGE__;
                color: #f5efe5;
            }

            [data-testid="stHeader"] {
                background: rgba(8, 11, 19, 0.78);
                backdrop-filter: blur(16px);
            }

            .block-container {
                padding-top: 2.4rem;
                max-width: 1220px;
            }

            .hero {
                animation: fadeUp 0.5s ease-out;
                background:
                    linear-gradient(135deg, rgba(15, 23, 42, 0.98), rgba(31, 41, 55, 0.9)),
                    linear-gradient(135deg, rgba(214, 183, 106, 0.13), rgba(255, 107, 95, 0.1));
                border: 1px solid rgba(214, 183, 106, 0.2);
                border-radius: 8px;
                box-shadow: 0 28px 90px rgba(0, 0, 0, 0.42);
                margin-bottom: 1.4rem;
                overflow: hidden;
                padding: 1.8rem;
                position: relative;
            }

            .hero:before {
                background: linear-gradient(90deg, #f5efe5, __ACCENT__, __DANGER__, #2dd4bf);
                content: "";
                height: 4px;
                left: 0;
                position: absolute;
                right: 0;
                top: 0;
            }

            .hero h1 {
                color: #fff7ed;
                font-size: 2.65rem;
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
                background: rgba(15, 23, 42, 0.72);
                border: 1px solid rgba(214, 183, 106, 0.25);
                border-radius: 999px;
                color: #f8fafc;
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
                background: rgba(15, 23, 42, 0.78);
                border: 1px dashed __ACCENT__;
                border-radius: 8px;
                padding: 0.75rem;
            }

            [data-testid="stTextArea"] textarea {
                background: rgba(15, 23, 42, 0.84);
                border: 1px solid rgba(214, 183, 106, 0.2);
                border-radius: 8px;
                color: #f8fafc;
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
                background: rgba(17, 24, 39, 0.92);
                border: 1px solid rgba(214, 183, 106, 0.16);
                border-radius: 8px;
                padding: 1rem;
            }

            [data-testid="stMetricLabel"] {
                color: #94a3b8;
            }

            [data-testid="stMetricValue"] {
                color: #fff7ed;
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
                background: linear-gradient(135deg, rgba(91, 28, 28, 0.62), rgba(17, 24, 39, 0.94));
            }

            .clean-card {
                border-color: __CLEAN__;
                background: linear-gradient(135deg, rgba(13, 76, 67, 0.52), rgba(17, 24, 39, 0.94));
            }

            .line-text {
                border-radius: 6px;
                font-size: 0.97rem;
                line-height: 1.6;
                margin: 0.65rem 0;
                padding: 0.75rem;
            }

            .copied-line {
                background: rgba(91, 28, 28, 0.78);
                border-left: 5px solid __DANGER__;
                color: #ffe4e0;
                font-weight: 600;
            }

            .clean-line {
                background: rgba(13, 76, 67, 0.7);
                border-left: 5px solid __CLEAN__;
                color: #ccfbf1;
            }

            .source-card {
                animation: fadeUp 0.34s ease-out;
                background: rgba(15, 23, 42, 0.76);
                border: 1px solid rgba(148, 163, 184, 0.18);
                border-radius: 8px;
                margin-top: 0.75rem;
                padding: 0.85rem;
                transition: border-color 0.18s ease, transform 0.18s ease;
            }

            .source-title {
                color: #fff7ed;
                font-weight: 700;
                margin-bottom: 0.25rem;
            }

            .source-url {
                color: #93c5fd;
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
                background: rgba(214, 183, 106, 0.16);
                border-radius: 999px;
                color: #fde68a;
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
                background: rgba(255, 107, 95, 0.14);
                border: 1px solid __DANGER__;
                border-radius: 999px;
                color: #fecaca;
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
                background: rgba(15, 23, 42, 0.78);
                border: 1px solid rgba(214, 183, 106, 0.18);
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
                background: rgba(15, 23, 42, 0.82);
                border: 1px solid rgba(214, 183, 106, 0.18);
                border-radius: 8px;
                color: #cbd5e1;
                line-height: 1.65;
                padding: 1.2rem;
            }

            section[data-testid="stExpander"] {
                background: rgba(15, 23, 42, 0.62);
                border: 1px solid rgba(148, 163, 184, 0.16);
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


def is_heading_or_structure(line):
    if re.match(r"^(chapter|section|figure|table|references|bibliography)\b", line, re.I):
        return True
    if re.match(r"^\d+(\.\d+)*\s+[A-Z]", line):
        return True
    if line.startswith("|") or line.count("|") >= 2:
        return True
    if line.startswith("![") or "attachment:" in line:
        return True
    return False


def split_lines(text, min_length=MIN_LINE_LENGTH, skip_structure=True, max_lines=None):
    lines = []
    in_references = False

    for line_number, raw_line in enumerate(text.splitlines(), start=1):
        line = clean_line(raw_line)
        if re.match(r"^(references|bibliography|works cited)\b", line, re.I):
            in_references = True
        if in_references:
            continue
        if skip_structure and is_heading_or_structure(line):
            continue
        if len(line) >= min_length:
            lines.append((line_number, line))
        if max_lines and len(lines) >= max_lines:
            break

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


def classify_match(match):
    if match["exact_phrase_found"] or match["score"] >= 0.9:
        return "Rewrite required"
    if match["score"] >= 0.78 or len(match["matched_words"]) >= 8:
        return "Citation needed"
    return "Editor review"


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
            matches[-1]["review_label"] = classify_match(matches[-1])

    return matches


def check_plagiarism(text, max_results, min_length=MIN_LINE_LENGTH, max_lines=None):
    checked_lines = split_lines(text, min_length=min_length, skip_structure=True, max_lines=max_lines)
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


def publication_readiness(report):
    similarity = report_score(report)
    rewrite_required = sum(
        1
        for item in report
        for match in item["matches"]
        if match.get("review_label") == "Rewrite required"
    )
    citation_needed = sum(
        1
        for item in report
        for match in item["matches"]
        if match.get("review_label") == "Citation needed"
    )
    score = max(0, 100 - similarity - (rewrite_required * 8) - (citation_needed * 3))
    if score >= 85:
        label = "Ready after normal editorial review"
    elif score >= 65:
        label = "Needs citation and wording review"
    else:
        label = "Not publication-ready yet"
    return score, label


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
            "editorial_label",
            "snippet",
            "search_note",
        ]
    )

    for item in report:
        if not item["matches"]:
            writer.writerow([item["line_number"], "No match found", item["line"], "", "", "", "", "", "", "", "", item.get("search_error", "")])
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
                    match.get("review_label", "Editor review"),
                    match["snippet"],
                    item.get("search_error", ""),
                ]
            )

    return output.getvalue()


def render_report(report, original_text):
    plagiarized = [item for item in report if item["matches"]]
    score = report_score(report)
    readiness_score, readiness_label = publication_readiness(report)

    st.subheader("Publication Similarity Summary")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Lines checked", len(report))
    col2.metric("Lines with source similarity", len(plagiarized))
    col3.metric("Similarity score", f"{score}%")
    col4.metric("Publication readiness", f"{readiness_score}%")

    st.download_button(
        "Download CSV report",
        data=build_csv_report(report),
        file_name="plagiarism_report.csv",
        mime="text/csv",
    )

    if score >= 50:
        st.error(f"{readiness_label}. Review the red lines, source links, and citation labels.")
    elif score > 0:
        st.warning(f"{readiness_label}. Some source similarity was found; review before submission.")
    else:
        st.success(f"{readiness_label}. No source similarity matches were found for the checked lines.")

    st.subheader("Editor Line-by-line Review")

    if not report:
        st.info("No long enough lines found to check.")
        return

    for item in report:
        found = bool(item["matches"])
        labels = sorted({match.get("review_label", "Editor review") for match in item["matches"]})
        label = ", ".join(labels) if found else "No similarity match"
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
                review_label = match.get("review_label", "Editor review")

                st.markdown(
                    f"""
                    <div class="source-card">
                        <div class="source-title">{highlighted_title}</div>
                        {source_link}
                        <div class="visible-url">{visible_url}</div>
                        <div class="score-pill">Similarity score: {match['score']}</div>
                        <div class="score-pill">Editorial label: {escape(review_label)}</div>
                        <div class="score-pill">{escape(exact_label)}</div>
                        <div class="source-explanation">
                            <strong>Why this needs publication review:</strong><br>
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
    st.set_page_config(page_title="Manuscript Originality Studio", page_icon="Search", layout="wide")
    apply_styles()

    st.markdown(
        """
        <div class="hero">
            <div class="hero-kicker">Publication Review Desk</div>
            <h1>Manuscript Originality Studio</h1>
            <p class="help-text">
                A polished pre-publication workspace for checking book chapters, tracing source overlap,
                and reviewing citation risks before submission.
            </p>
            <div class="feature-strip">
                <span class="feature-chip">Editorial similarity review</span>
                <span class="feature-chip">Source evidence links</span>
                <span class="feature-chip">Reference-aware scanning</span>
                <span class="feature-chip">Citation risk labels</span>
                <span class="feature-chip">CSV evidence report</span>
                <span class="feature-chip">Publication readiness score</span>
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
            scan_mode = st.selectbox("Publication scan mode", list(SCAN_MODES.keys()), index=1)
            mode_settings = SCAN_MODES[scan_mode]
            max_results = mode_settings["max_results"]
            st.markdown(
                f"""
                <div class="control-help">
                    <strong>{escape(scan_mode)}</strong><br>
                    Checks up to <strong>{max_results}</strong> source results per line and reviews up to
                    <strong>{mode_settings["max_lines"]}</strong> manuscript lines.
                    <br><br>
                    Publisher workflows use similarity reports as editorial evidence, not automatic proof of plagiarism.
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

            report = check_plagiarism(
                text,
                max_results=max_results,
                min_length=mode_settings["min_length"],
                max_lines=mode_settings["max_lines"],
            )
            render_report(report, text)

    with guidance_tab:
        st.markdown(
            """
            <div class="guide-panel">
                <strong>How publishers use similarity checks</strong><br>
                Publishers usually screen manuscripts for text similarity, then editors decide whether the overlap is
                plagiarism, missing citation, text recycling, or harmless common wording.
                <br><br>
                <strong>Rewrite required</strong> means the line appears very close to a source and should be rewritten,
                quoted, or cited before submission.
                <br><br>
                <strong>Citation needed</strong> means the idea or wording may be source-dependent and should be checked
                against the source link.
                <br><br>
                <strong>Editor review</strong> means some similarity was found, but it may be common terminology or
                acceptable overlap.
                <br><br>
                <strong>Skipped content</strong> includes headings, tables, image artifacts, and references, because
                these often inflate similarity scores in book chapters.
                <br><br>
                <strong>Publication readiness</strong> is a practical risk score for revision planning. It is not a
                publisher certificate or an iThenticate replacement.
                <br><br>
            </div>
            """,
            unsafe_allow_html=True,
        )


if __name__ == "__main__":
    main()
