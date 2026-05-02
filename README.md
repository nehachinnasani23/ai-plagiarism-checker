# Plagiarism Source Finder

A Python and Streamlit web app that checks pasted text or uploaded files for possible plagiarism. The app scans long lines, searches the web for matching content, and explains which words matched with source links and similarity scores.

## Live Demo

[Open the deployed app](https://ai-plagiarism-checker-rvpxuf4wj8bmy43xqngxuz.streamlit.app/)

## Project Goal

The goal of this project is to help students, writers, and reviewers quickly identify lines that may need citations or rewriting. It is designed as a practical career portfolio project showing Python, file processing, web search integration, and interactive app development.

## Features

- Paste text directly into the app
- Upload `.txt`, `.pdf`, or `.docx` files
- Animated result cards
- Polished light research-dashboard theme
- Check content line by line
- Show possible source links
- Show full source URLs
- Highlight suspicious lines in red
- Highlight matching words inside source snippets
- Show which words matched between your sentence and the source
- Explain why each line may be plagiarism
- Explain the search results control inside the app
- Keep search failures separate from plagiarism matches
- Display similarity scores
- Calculate a possible plagiarism percentage
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
3. Text is split into individual lines.
4. Short lines are skipped to reduce weak matches.
5. Each long line is searched as an exact phrase.
6. Search snippets are compared with the original line.
7. Possible matches are shown with red line highlights, source links, full URLs, matched words, highlighted source words, scores, and an explanation of why the line was flagged.
8. The user can download the results as a CSV report or create a corrected DOCX.

## Corrected DOCX Export

The app can generate a new Word document after scanning.

- If `OPENAI_API_KEY` is configured in Streamlit secrets or your local environment, flagged lines are rewritten in original wording.
- If no API key is configured, the DOCX still marks the flagged lines and includes source links so they can be rewritten manually.
- The original uploaded file is not changed.

## Limitations

This starter version works best for exact or near-exact copied text from searchable web pages. It is not a replacement for academic plagiarism systems such as Turnitin. Future improvements could include semantic similarity, citation detection, AI paraphrase detection, and PDF report generation.

## Future Improvements

- Add AI-based paraphrase detection
- Add account login and saved reports
- Generate downloadable PDF reports
- Highlight copied text inside uploaded documents
- Add citation suggestions
- Improve hosted performance for large files
