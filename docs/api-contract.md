# API Contract Notes

## Authentication

- `POST /api/v1/auth/login`
- Content-Type: `application/x-www-form-urlencoded`
- Form fields: `username`, `password`
- Token source in response: `access_token`

## EDMS Publish Batch

- `POST /api/v1/bim/edms/publish-batch`
- Mode 1 (MVP default): `multipart/form-data`
  - Scalar fields: `run_client_id`, `project_code`, `revit_version`, `model_guid`, `model_title`, `plugin_version`
  - JSON string fields:
    - `items_json`: list of publish items
    - `files_manifest`: list of filename bindings
  - File fields:
    - repeated `files[]`
- Mode 2 fallback: `application/json` (hash-only publish without file upload)

### Publish item payload

`item_index`, `sheet_unique_id`, `sheet_number`, `sheet_name`, `doc_number`, `requested_revision`, `status_code`, `include_native`, `metadata`, `file_sha256`

### Files manifest payload

`item_index`, `sheet_unique_id`, `pdf_file_name`, `native_file_name`, `file_sha256`

## Retry Strategy

- Plugin retries transient network failures in HTTP client layer.
- Use-case retry is item-level:
  - take previous batch response
  - rebuild request only for `state=failed`
  - submit a new run with new `run_client_id`
