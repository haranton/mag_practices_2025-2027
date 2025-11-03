#!/usr/bin/env bash
set -euo pipefail
file="${1:-}"
[ -z "$file" ] && { echo "Usage: $0 <filename>" >&2; exit 1; }

mime=$(file --mime-type -b "$file")
ext="${file##*.}"
ext=$(echo "$ext" | tr '[:upper:]' '[:lower:]')

convert_pdf() { pdftotext -layout -enc UTF-8 "$file" - 2>/dev/null || true; }
convert_doc() { antiword "$file" 2>/dev/null || catdoc "$file" 2>/dev/null || pandoc -f doc -t plain "$file" 2>/dev/null || true; }
convert_docx_odt() { pandoc -t plain "$file" 2>/dev/null || soffice --headless --convert-to txt:Text --outdir /tmp "$file" >/dev/null 2>&1 && cat "/tmp/${file%.*}.txt"; }
convert_xls() { xls2csv "$file" 2>/dev/null || ssconvert --export-type=Gnumeric_stf:stf_csv "$file" fd://1 2>/dev/null || true; }
convert_xlsx_ods() { ssconvert --export-type=Gnumeric_stf:stf_csv "$file" fd://1 2>/dev/null || soffice --headless --convert-to csv --outdir /tmp "$file" >/dev/null 2>&1 && cat "/tmp/${file%.*}.csv"; }
convert_ppt() { catppt "$file" 2>/dev/null || true; }
convert_pptx_odp() { pandoc -t plain "$file" 2>/dev/null || soffice --headless --convert-to txt:Text --outdir /tmp "$file" >/dev/null 2>&1 && cat "/tmp/${file%.*}.txt"; }
convert_text() { cat "$file"; }

case "$mime" in
    application/pdf) convert_pdf ;;
    application/msword) convert_doc ;;
    application/vnd.openxmlformats-officedocument.wordprocessingml.document|application/vnd.oasis.opendocument.text) convert_docx_odt ;;
    application/vnd.ms-excel) convert_xls ;;
    application/vnd.openxmlformats-officedocument.spreadsheetml.sheet|application/vnd.oasis.opendocument.spreadsheet) convert_xlsx_ods ;;
    application/vnd.ms-powerpoint) convert_ppt ;;
    application/vnd.openxmlformats-officedocument.presentationml.presentation|application/vnd.oasis.opendocument.presentation) convert_pptx_odp ;;
    text/*) convert_text ;;
    *)
        case "$ext" in
            pdf) convert_pdf ;;
            doc) convert_doc ;;
            docx|odt) convert_docx_odt ;;
            xls) convert_xls ;;
            xlsx|ods) convert_xlsx_ods ;;
            ppt) convert_ppt ;;
            pptx|odp) convert_pptx_odp ;;
            txt|csv|md|log) convert_text ;;
            *) echo "Unsupported file type: $mime ($ext)" >&2; exit 0 ;;
        esac
        ;;
esac
