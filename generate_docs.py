"""
generate_docs.py
Production-ready CSV → HTML-template → PDF generator using xhtml2pdf.

Requirements:
  - Python 3.9+
  - xhtml2pdf (only external dependency, pure pip install)

Usage:
  python generate_docs.py --csv CSV/input.csv --template HTML/template.html --out output

  If --csv and/or --template are omitted, you can choose interactively from:
    - CSV/  (input data files)
    - HTML/ (HTML templates)
"""

from __future__ import annotations

import argparse
import csv
import os
import platform
import re
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List, Optional


PLACEHOLDER_RE = re.compile(r"\{\{\s*([a-zA-Z0-9_]+)\s*\}\}")


DEFAULT_CSS = r"""
/* ---- Fonts (Cyrillic-friendly if installed) ---- */
@font-face {
  font-family: "DocFont";
  src: local("Roboto"), local("Roboto Regular"), local("Liberation Sans"), local("Arial");
  font-weight: normal;
  font-style: normal;
}
@font-face {
  font-family: "DocFont";
  src: local("Roboto Bold"), local("Roboto-Bold"), local("Liberation Sans Bold"), local("Arial Bold");
  font-weight: bold;
  font-style: normal;
}

/* ---- Page / print settings ---- */
@page {
  size: A4;
  margin: 14mm 12mm 16mm 12mm;
}

html, body {
  font-family: "DocFont", "Roboto", "Liberation Sans", Arial, sans-serif;
  font-size: 12pt;
  line-height: 1.35;
  color: #111;
}

/* Long strings (emails, ids, urls, etc.) */
* {
  overflow-wrap: anywhere;
  word-break: break-word;
}

h1, h2, h3 {
  margin: 0 0 10px 0;
}

.meta {
  color: #555;
  font-size: 10pt;
  margin-bottom: 10px;
}

/* ---- Table styling ---- */
table {
  width: 100%;
  border-collapse: collapse;
  margin: 10px 0;
}
th, td {
  border: 1px solid #222;
  padding: 8px 10px;
  vertical-align: top;
}
th {
  text-align: center;
  background: #f3f3f3;
}
td {
  text-align: center;
}

/* Utility */
.left { text-align: left; }
.right { text-align: right; }
.center { text-align: center; }
"""


@dataclass(frozen=True)
class RenderResult:
    html: str
    missing_placeholders: List[str]


def _list_files(dir_path: Path, allowed_suffixes: List[str]) -> List[Path]:
    if not dir_path.exists() or not dir_path.is_dir():
        return []
    files = [p for p in dir_path.iterdir() if p.is_file() and p.suffix.lower() in allowed_suffixes]
    files.sort(key=lambda p: p.name.lower())
    return files


def _prompt_choice(title: str, files: List[Path]) -> Path:
    """
    Prompt user to choose a file by number from a list.
    """
    if not files:
        raise FileNotFoundError(f"No files available for selection: {title}")

    print(f"\n{title}")
    for i, p in enumerate(files, start=1):
        print(f"  {i}) {p.name}")

    while True:
        choice = input("Select number: ").strip()
        if not choice:
            print("Please enter a number.")
            continue
        if not choice.isdigit():
            print("Invalid input. Enter a number like 1, 2, 3...")
            continue
        n = int(choice)
        if 1 <= n <= len(files):
            return files[n - 1]
        print(f"Out of range. Enter 1..{len(files)}.")


def _resolve_or_select_file(
    provided: Optional[str],
    project_root: Path,
    folder_name: str,
    allowed_suffixes: List[str],
    title: str,
) -> Path:
    """
    If provided path is given, resolve it. Otherwise, select from project_root/folder_name.
    """
    if provided:
        p = Path(provided).expanduser()
        # If it's a relative path, resolve relative to current working directory.
        return p.resolve()

    dir_path = (project_root / folder_name).resolve()
    files = _list_files(dir_path, allowed_suffixes)
    return _prompt_choice(f"{title} (from {dir_path})", files)


def load_csv(csv_path: Path) -> List[Dict[str, str]]:
    """
    Load a UTF-8 CSV into a list of dicts.
    - Uses utf-8-sig to gracefully handle BOM.
    - Empty cells become "".
    """
    if not csv_path.exists() or not csv_path.is_file():
        raise FileNotFoundError(f"CSV file not found: {csv_path}")

    rows: List[Dict[str, str]] = []
    with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        if reader.fieldnames is None:
            raise ValueError("CSV has no header row (column names).")

        for i, row in enumerate(reader, start=1):
            # Normalize None values from DictReader to empty strings
            cleaned = {k: (v if v is not None else "") for k, v in row.items()}
            rows.append(cleaned)

    if not rows:
        print("WARNING: CSV has headers but contains zero data rows.")

    return rows


def _inject_css_into_html(html: str, css: str) -> str:
    """
    Inject a <style> block into HTML.
    Prefer inserting into <head> if present; otherwise prepend to document.
    """
    style_block = f"<style>\n{css}\n</style>\n"

    # Insert before </head> if it exists (case-insensitive).
    m = re.search(r"</head\s*>", html, flags=re.IGNORECASE)
    if m:
        idx = m.start()
        return html[:idx] + style_block + html[idx:]

    # If there's <html> but no head, inject after <html...>
    m = re.search(r"<html[^>]*>", html, flags=re.IGNORECASE)
    if m:
        idx = m.end()
        return html[:idx] + "\n<head>\n" + style_block + "</head>\n" + html[idx:]

    # Fallback: just prepend
    return style_block + html


def render_template(template_text: str, row: Dict[str, str]) -> RenderResult:
    """
    Render an HTML template by replacing {{column_name}} placeholders with CSV row values.
    Missing placeholders are replaced with "" and reported.
    """
    missing: List[str] = []

    def repl(match: re.Match[str]) -> str:
        key = match.group(1)
        if key not in row:
            missing.append(key)
            return ""
        return str(row.get(key, ""))

    rendered = PLACEHOLDER_RE.sub(repl, template_text)

    # Add a tiny footer timestamp (non-intrusive) if the template wants it.
    # Users can include {{generated_at}} in templates without putting it in CSV.
    if "generated_at" in missing:
        missing.remove("generated_at")
    rendered = rendered.replace("{{generated_at}}", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    return RenderResult(html=rendered, missing_placeholders=sorted(set(missing)))


def html_to_pdf(html: str, out_pdf_path: Path, base_url: Optional[str] = None) -> None:
    """
    Convert HTML string into a PDF file via xhtml2pdf.

    Note: xhtml2pdf supports a useful subset of HTML/CSS. The included DEFAULT_CSS
    is kept within what xhtml2pdf usually handles (tables, basic fonts, margins).
    """
    try:
        from xhtml2pdf import pisa  # type: ignore
    except Exception as e:  # pragma: no cover
        raise RuntimeError(
            "xhtml2pdf is not installed or failed to import. Install with:\n"
            "  pip install xhtml2pdf\n"
        ) from e

    try:
        out_pdf_path.parent.mkdir(parents=True, exist_ok=True)
        with out_pdf_path.open("wb") as pdf_file:
            # pisa.CreatePDF returns an object with .err count; 0 means success.
            result = pisa.CreatePDF(html, dest=pdf_file, link_callback=None)
        if result.err:
            raise RuntimeError(f"xhtml2pdf reported {result.err} error(s) while creating PDF.")
    except Exception as e:
        raise RuntimeError(f"xhtml2pdf failed to generate PDF: {out_pdf_path}\nReason: {e}") from e


def open_file(path: Path) -> None:
    """
    Open a file in the default viewer.
    - Windows: os.startfile
    - macOS: open
    - Other OS: no-op (prints a message)
    """
    system = platform.system()
    try:
        if system == "Windows":
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif system == "Darwin":
            subprocess.run(["open", str(path)], check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        else:
            print(f"INFO: Auto-open is not implemented for OS '{system}'. File saved at: {path}")
    except Exception as e:
        print(f"WARNING: Failed to open file automatically: {path}\nReason: {e}")


def _sanitize_filename(name: str) -> str:
    """
    Make a safe filename for Windows/macOS.
    """
    name = name.strip()
    if not name:
        return "document"
    # Replace forbidden characters for Windows filenames: <>:"/\\|?*
    name = re.sub(r'[<>:"/\\\\|?*]+', "_", name)
    # Avoid trailing dots/spaces on Windows
    name = name.rstrip(". ").strip()
    return name or "document"


def _choose_output_filename(row: Dict[str, str], index: int) -> str:
    """
    Prefer a user-provided column if present (filename / id / name), otherwise use an index.
    """
    for key in ("filename", "file_name", "doc_name", "document_name", "id", "name"):
        val = (row.get(key) or "").strip()
        if val:
            return _sanitize_filename(val)
    return f"document_{index:04d}"


def main(argv: Optional[Iterable[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Generate PDFs from a CSV and an HTML template (WeasyPrint).")
    parser.add_argument("--csv", required=False, help="Path to input CSV (UTF-8). If omitted, choose from CSV/ folder.")
    parser.add_argument(
        "--template",
        required=False,
        help="Path to HTML template with {{column_name}} placeholders. If omitted, choose from HTML/ folder.",
    )
    parser.add_argument("--out", default="output", help="Output directory for generated PDFs (default: output).")
    args = parser.parse_args(list(argv) if argv is not None else None)

    project_root = Path(__file__).resolve().parent

    try:
        csv_path = _resolve_or_select_file(
            args.csv,
            project_root=project_root,
            folder_name="CSV",
            allowed_suffixes=[".csv"],
            title="Choose CSV file",
        )
        template_path = _resolve_or_select_file(
            args.template,
            project_root=project_root,
            folder_name="HTML",
            allowed_suffixes=[".html", ".htm"],
            title="Choose HTML template",
        )
    except FileNotFoundError as e:
        print(f"ERROR: {e}")
        print(f"TIP: Put your CSV files into: {project_root / 'CSV'}")
        print(f"TIP: Put your HTML templates into: {project_root / 'HTML'}")
        return 2

    out_dir = Path(args.out).expanduser().resolve()

    # Input validation
    if not csv_path.exists():
        print(f"ERROR: CSV file does not exist: {csv_path}")
        return 2
    if not template_path.exists():
        print(f"ERROR: Template file does not exist: {template_path}")
        return 2
    if not template_path.is_file():
        print(f"ERROR: Template path is not a file: {template_path}")
        return 2

    try:
        template_text = template_path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        print("ERROR: Template file is not UTF-8. Save it as UTF-8 and retry.")
        return 2
    except Exception as e:
        print(f"ERROR: Failed to read template: {template_path}\nReason: {e}")
        return 2

    try:
        rows = load_csv(csv_path)
    except Exception as e:
        print(f"ERROR: Failed to load CSV: {csv_path}\nReason: {e}")
        return 2

    # Ensure output directory exists
    try:
        out_dir.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        print(f"ERROR: Failed to create output directory: {out_dir}\nReason: {e}")
        return 2

    print(f"CSV:      {csv_path}")
    print(f"Template: {template_path}")
    print(f"Output:   {out_dir}")
    print(f"Rows:     {len(rows)}")

    # Inject our baseline CSS for print/table/wrapping/fonts
    template_with_css = _inject_css_into_html(template_text, DEFAULT_CSS)

    failures = 0
    generated_paths: List[Path] = []

    for idx, row in enumerate(rows, start=1):
        print(f"\n[{idx}/{len(rows)}] Rendering row...")
        render = render_template(template_with_css, row)
        if render.missing_placeholders:
            print("WARNING: Missing columns for placeholders: " + ", ".join(render.missing_placeholders))

        filename = _choose_output_filename(row, idx) + ".pdf"
        out_pdf = out_dir / filename

        print(f"[{idx}/{len(rows)}] Generating PDF: {out_pdf.name}")
        try:
            html_to_pdf(render.html, out_pdf, base_url=str(template_path.parent))
            generated_paths.append(out_pdf)
            print(f"[{idx}/{len(rows)}] OK")
        except Exception as e:
            failures += 1
            print(f"[{idx}/{len(rows)}] ERROR: {e}")
            continue

    print("\nDone.")
    print(f"Generated: {len(generated_paths)} PDFs")
    if failures:
        print(f"Failed:    {failures} rows (see errors above)")

    # Auto-open each PDF (as requested)
    for p in generated_paths:
        open_file(p)

    return 0 if failures == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())


# -----------------------------
# Example HTML template (UTF-8)
# -----------------------------
# Note: this same sample is also saved as a file in the project:
#   HTML/template_example.html
EXAMPLE_TEMPLATE_HTML = r"""<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8" />
  <title>Документ</title>
</head>
<body>
  <h1>Справка</h1>
  <div class="meta">Сгенерировано: {{generated_at}}</div>

  <p><b>ФИО:</b> {{full_name}}</p>
  <p><b>Должность:</b> {{position}}</p>
  <p><b>Отдел:</b> {{department}}</p>

  <h2>Детали</h2>
  <table>
    <thead>
      <tr>
        <th>Параметр</th>
        <th>Значение</th>
      </tr>
    </thead>
    <tbody>
      <tr>
        <td class="left">Email</td>
        <td class="left">{{email}}</td>
      </tr>
      <tr>
        <td class="left">Телефон</td>
        <td class="left">{{phone}}</td>
      </tr>
      <tr>
        <td class="left">Комментарий</td>
        <td class="left">{{comment}}</td>
      </tr>
    </tbody>
  </table>
</body>
</html>
"""


# -----------------------------
# Example CSV (UTF-8)
# -----------------------------
# Note: this same sample is also saved as a file in the project:
#   CSV/data_example.csv
EXAMPLE_CSV = """full_name,position,department,email,phone,comment,filename
Иванов Иван Иванович,Инженер,ИТ,ivanov@example.com,+7 999 123-45-67,"Очень длинная строка для проверки переноса: https://example.com/very/long/path?with=params&and=more",ivanov_doc
Петров Пётр Петрович,Менеджер,Продажи,petrov@example.com,+7 999 222-33-44,"Комментарий с кириллицей и знаками препинания.",petrov_doc
"""


# -----------------------------
# Installation notes
# -----------------------------
"""
Install xhtml2pdf (no external system libraries required):
  pip install xhtml2pdf

Or inside this project (using requirements.txt):
  pip install -r requirements.txt

The library is pure Python + ReportLab and works the same way on Windows and macOS
without installing GTK/Cairo/Pango.
"""
