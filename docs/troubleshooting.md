# Troubleshooting

## Login fails with 401

- Check plugin sends form fields (`username`, `password`), not JSON.
- Validate API base URL points to MDR backend root (example: `http://127.0.0.1:8000`).

## Publish fails with `validation_error` for PDF hash

- Ensure item has either:
  - `PdfFilePath` to an existing file, or
  - precomputed `FileSha256`
- In command flow, verify export output directory is writable.

## Publish fails with `file_policy_rejected`

- Backend storage policy rejected mime/size.
- Confirm generated file extensions and upload content are acceptable.

## Publish has failed items after run

- Enable `retryFailedItems`.
- Retry path resubmits only items with `state=failed` from previous run.

## No log files generated

- Verify write access to `%LocalAppData%/MDR/RevitPlugin/logs`.
