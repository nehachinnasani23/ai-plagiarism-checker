# Plagiarism Source Finder

A Python and Streamlit pre-publication similarity review app for book chapters and manuscripts. The app scans uploaded text, finds possible source overlap, labels editorial risk, and exports a corrected DOCX review file.

## Live Demo

[Open the deployed app](https://ai-plagiarism-checker-rvpxuf4wj8bmy43xqngxuz.streamlit.app/)

## Project Goal

The goal of this project is to help authors, students, editors, and reviewers identify text that may need rewriting, quotation, or citation before publication. It follows the same general idea used by publication workflows: generate a similarity report first, then let a human editor decide whether the overlap is plagiarism, missing citation, text recycling, or harmless common wording.

## Features

- Paste text directly into the app
- Upload `.txt`, `.pdf`, or `.docx` files
- Animated result cards
- Polished light research-dashboard theme
- Publication scan modes: Fast precheck, Publication review, and Deep editorial scan
- Skip headings, tables, figures, image artifacts, and references
- Show possible source links
- Show full source URLs
- Highlight suspicious lines in red
- Highlight matching words inside source snippets
- Show which words matched between your sentence and the source
- Label matches as Rewrite required, Citation needed, or Editor review
- Keep search failures separate from similarity matches
- Display similarity scores
- Calculate a similarity percentage and publication readiness score
- Download a CSV report
- Download a corrected DOCX with flagged lines rewritten when `OPENAI_API_KEY` is configured
- Run locally in the browser

## Tech Stack

- Python
- Streamlit
- OpenAI Python SDK
- DDGS search
- pypdf
- python-docx
- difflib similarity matching

## Install

```zsh
cd ~/ai-plagiarism-checker
python3.13 -m pip install --user --break-system-packages -r requirements.txt
```

## Run

```zsh
cd ~/ai-plagiarism-checker
python3.13 -m streamlit run app.py
```

Then open:

```text
http://localhost:8501
```

## How It Works

1. The user uploads a file or pastes text.
2. The app extracts text from the input.
3. Headings, tables, image artifacts, and reference sections are skipped.
4. Meaningful manuscript lines are searched as exact phrases.
5. Search snippets are compared with the original line.
6. Matches are labeled as Rewrite required, Citation needed, or Editor review.
7. Results show source links, full URLs, matched words, highlighted source words, scores, and an explanation of why the line was flagged.
8. The user can download the results as a CSV report or create a corrected DOCX.

## Publication Workflow Basis

Publication workflows typically use similarity screening as editorial evidence rather than an automatic plagiarism verdict.

- Crossref Similarity Check describes similarity screening as feedback on manuscript similarity to published and web content.
- Crossref notes that iThenticate checks similarity, not plagiarism by itself.
- Springer Nature says copied sentences without proper citation can be considered plagiarism.
- Elsevier describes plagiarism as copying or paraphrasing substantial parts of another work without attribution.
- COPE guidance treats suspected plagiarism and text recycling as editorial review matters.

## Corrected DOCX Export

The app can generate a new Word document after scanning.

- If `OPENAI_API_KEY` is configured in Streamlit secrets or your local environment, flagged lines are rewritten in original wording.
- If no API key is configured, the DOCX still marks the flagged lines and includes source links so they can be rewritten manually.
- The original uploaded file is not changed.

## Limitations

This project is a portfolio-friendly pre-publication review tool, not an iThenticate, Turnitin, or publisher certification replacement. It works best for exact or near-exact copied text from searchable web pages. Final publication decisions still require human editorial review.

## Future Improvements

- Add AI-based paraphrase detection
- Add account login and saved reports
- Generate downloadable PDF reports
- Highlight copied text inside uploaded documents
- Add citation suggestions
- Improve hosted performance for large files
