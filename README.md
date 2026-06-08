```markdown
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
Languages: [简体中文](README.zh-CN.md)

# ⚡ LocalKB | Local Knowledge Base
An offline document knowledge base written in pure Python, all data stored locally without cloud upload, compatible with Windows / macOS / Linux.

## Introduction
- Multi-format document parsing & FTS5 full-text local search
- Two import modes with different backup policies
- Dual view: full-file preview or split-item pagination view
- Export search result as Markdown / TXT / Word

### Supported Formats
`PDF / DOCX / PPTX / TXT / MD / HTML / HTM`
⚠️ Unsupported for now: .doc(Office 97-2003), .ppt(Office 97-2003), encrypted/scanned PDFs and PDFs with custom embedded fonts

## Backup Policy (Important)
1. **Upload via webpage**: Original files auto-saved into `upload_files` folder
2. **Import from local folder**: Only extract text to database, NO original file copy & backup.

## Install Dependencies
```bash
pip install streamlit PyPDF2 python-docx python-pptx chardet olefile
```

## Run Project

```bash
streamlit run knowledge_base.py
```

Default URL: `http://127.0.0.1:8501`

## Usage Guide

### 1. Sidebar - Import Documents

- Upload files: auto backup source files to `upload_files`

- Import folder: parse content only without backup original files

- Delete: select filename to remove all related records from database

### 2. Two View Modes

- File Mode: page list, click to view full document content

- Item Mode: split document into paragraphs and browse by page

### 3. Fixed Top Search Bar

Search box at top of page, full-text search and multi-format export available.

## Project Structure

```Plain Text
LocalKB/
├─ knowledge_base.py    # Main source code
├─ LICENSE                      # MIT License
├─ README.zh-CN.md    # Chinese Readme
├─ README.md               # English Readme
├─ knowledge.db            # Auto-generated database
└─ upload_files                # Backup folder for uploaded files only
```

## License
Released under MIT License, free to use, modify and commercial use.
Copyright (c) 2026 CCking2022

## ⚠️ Limitations
The program runs stably with all exceptions caught to avoid crashes. The following formats are not yet fully supported:
1. PDF: Encrypted PDFs, image-only scanned PDFs and PDFs with special embedded fonts may produce garbled text or fail extraction; standard text-based PDFs parse correctly.
2. Office: Legacy binary formats .doc / .ppt (Office 97~2003) are unsupported; modern .docx / .pptx files work perfectly.

## Future Plan
- Improve PDF parsing and fix font garbling issues;
- Integrate antiword + catdoc to add support for legacy .doc and .ppt files;
- Add Tesseract-OCR integration to handle scanned PDF documents.
- Add Excel & EPUB parsing support
- Local LLM semantic search
- Document tag management
- Package to standalone EXE

⭐ Star this repo if helpful!
