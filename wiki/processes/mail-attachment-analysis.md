# Mail Attachment Analysis

## Summary

This process describes how the project analyzes email attachments after they have already been safely archived in `raw/mail/`.

The raw archive keeps original binaries untouched.
The derived layer produces structured attachment insights for operator review and future automation.

## Flow

1. Outlook mail is synced into `raw/mail/messages/...`.
2. Each message keeps `attachments.json` plus the original saved files.
3. The processing pipeline analyzes attachments into `derived/mail/<message-key>/attachment_analysis.json`.
4. The SQLite read model stores the message-level attachment summary and attachment insights.
5. `Needs Decision` shows attachment summaries directly in the operator queue.

## Current supported attachment analysis

- Excel
  - `.xlsx`, `.xlsm`, `.xltx`, `.xltm`
  - workbook sheet names and sheet count
  - `.csv` preview rows
- PDF
  - page count
  - text preview when optional PDF text extraction is available
- Image
  - dimensions for PNG, JPEG, GIF, and BMP
  - OCR status explicitly marked as not configured in the current MVP
- Other
  - metadata-only storage for unsupported formats such as `docx`

## Storage contract

Raw layer:

- `raw/mail/messages/<message-folder>/attachments/`
- `raw/mail/messages/<message-folder>/attachments.json`

Derived layer:

- `derived/mail/<message-key>/attachment_analysis.json`

Database/read-model layer:

- analyzed attachment count
- attachment summary
- attachment kinds
- attachment analysis path
- detailed attachment insight list

## Design rules

- never modify the original attachment binary in `raw/`
- keep derived analysis outside `raw/`
- degrade gracefully to `metadata_only` instead of failing the pipeline
- make attachment kinds visible in labels such as `attachment:excel`
- keep the operator queue informative even when OCR or deep parsing is not configured

## Current limitations

- OCR is not configured yet for images
- PDF text extraction depends on optional parser availability
- legacy Excel formats such as `.xls` and `.xlsb` currently stay metadata-first
- some filenames still arrive with encoding issues from the upstream mail source

## Next recommended improvements

1. add OCR for image attachments
2. improve PDF text extraction consistently across environments
3. extract richer workbook previews from Excel
4. promote attachment insights into separate analytics-ready entities when needed
