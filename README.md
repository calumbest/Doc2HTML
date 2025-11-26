# Doc2HTML

A Google Apps Script add-on that converts Google Docs to HTML format.

## What it does

This tool automates the conversion of Google Docs into clean HTML code, handling:
- Rich text formatting (bold, italic, underline, links)
- Headers (H2, H3, H4, H5)
- Bulleted and numbered lists (including nested lists)
- Tables
- Footnotes
- Special characters and HTML entities

The add-on provides a step-by-step interface within Google Docs, guiding users through the conversion process with buttons for each formatting operation.

## Project history

Most of the code in this repository was written with GPT-4 in late 2023 as part of my first coding project. As of this commit, the resulting app is used for virtually all research page publication at GiveWell.

It's possible that the app could be improved substantially, as I haven't returned to this project in over a year.

## Files

- `Code.js` - Main application code
- `appsscript.json` - Apps Script manifest
- `.clasp.json` - Configuration for Google's clasp CLI tool
- Image files - Logo and screenshot assets for the add-on
