# Test Plan

## EDMS Publish MVP

1. Login path
   - verify `/api/v1/auth/login` uses form payload (`username`, `password`)
   - reject empty credentials
2. Publish payload mode
   - with local PDF path -> `multipart/form-data` with `items_json`, `files_manifest`, `files[]`
   - hash-only fallback -> JSON body with `file_sha256`
3. Retry flow
   - initial run with mixed `completed/failed`
   - `RetryFailedAsync` sends only previously failed items
4. Validation
   - missing `projectCode`, `sheetUniqueId`, `requestedRevision`, and PDF/hash should fail fast
   - non-existing `PdfFilePath` should fail fast

## Backend contract alignment

1. `POST /api/v1/bim/edms/publish-batch` multipart request with real PDF file should pass.
2. idempotency and revision conflict behavior should stay unchanged.
