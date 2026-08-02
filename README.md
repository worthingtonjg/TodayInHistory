# Today In History Email => Document

This repository turns a Google Takeout Gmail export for the `School` label into:

1. grouped JSON files for each supported newsletter type
2. formatted DOCX files for review or reuse

It is intentionally narrow in scope. The parser is built for two specific newsletter formats:

- `Back Then History`
- `The Retrospect`

Messages from other senders or unrelated mail in the same export are ignored.

The workflow is split into two local steps:

- `import_mbox.py` reads the Gmail `.mbox` export and writes JSON for the supported newsletters
- `export_docs.py` reads those JSON files and generates DOCX files

## What You Need

- A Gmail account with the `School` label export you want to process
- Google Takeout access
- Python installed on Windows

You do not need to install Python packages manually for the current workflow. `export_docs.py` is set up to run with the bundled Codex Python runtime if your local Python is missing the required DOCX/image dependencies.

## Folder Layout

The repo expects this structure:

- `input/` for the Gmail export
- `output/` for generated JSON and DOCX files

Typical paths:

- `input/School.mbox`
- `output/Back Then History.json`
- `output/The Retrospect.json`
- `output/Back Then History/<doc>.docx`
- `output/The Retrospect/<doc>.docx`

## Step 1: Export Gmail With Google Takeout

Use Google Takeout to export the messages from Gmail.

1. Open Google Takeout.
2. Select Gmail.
3. Filter or limit the export to the `School` label.
4. Choose the export format that includes an `.mbox` file.
5. Download the archive when Google finishes building it.
6. Extract the archive.
7. Copy the `School.mbox` file into this repo's `input/` folder.

If the Takeout archive contains other mail besides the supported newsletters, that is fine. The import step will skip anything it does not recognize.

## Step 2: Convert the MBOX To JSON

Run:

```powershell
python import_mbox.py
```

What this does:

- reads `input/School.mbox`
- parses each message
- groups supported emails into:
  - `Back Then History`
  - `The Retrospect`
- writes one JSON file per supported type into `output/`

Unsupported email types are skipped and not written into the JSON output. This is not a general-purpose Gmail parser.

Helpful flags:

- `--mbox` to point at a different `.mbox`
- `--output-dir` to write the JSON somewhere else
- `--limit` to print only the first few parsed messages for inspection

Example:

```powershell
python import_mbox.py --mbox input\School.mbox --output-dir output
```

## Step 3: Generate DOCX Files

Run:

```powershell
python export_docs.py
```

What this does:

- reads `output/Back Then History.json`
- reads `output/The Retrospect.json`
- generates DOCX files under `output/`
- creates subfolders by message type automatically

Helpful flags:

- `--input-dir` to read JSON from a different folder
- `--backthen-limit` to generate only the first N `Back Then History` documents
- `--retrospect-limit` to generate only the first N `The Retrospect` documents
- `--skip-existing` to resume a partial run without overwriting DOCX files that are already present

Example:

```powershell
python export_docs.py --input-dir output --backthen-limit 20 --retrospect-limit 20
```

To resume a partially completed export without overwriting files that already exist:

```powershell
python export_docs.py --skip-existing
```

## End-to-End Command Sequence

For a fresh export, the usual sequence is:

```powershell
python import_mbox.py
python export_docs.py
```

After that, inspect the DOCX files in `output\Back Then History\` and `output\The Retrospect\`.

## Repeating The Process Later

To export more mail in the future:

1. Use Google Takeout again.
2. Replace `input/School.mbox` with the new `.mbox` file.
3. Rerun `python import_mbox.py`.
4. Rerun `python export_docs.py`.

The scripts overwrite files with the same names, but they do not clear the output folder first.

## Notes

- `import_mbox.py` no longer writes `.eml` files.
- The export stage reads the grouped JSON, not the `.mbox` directly.
- The repo is designed so you can rerun the same pipeline against new exports of these same newsletter formats without rewriting the scripts.
