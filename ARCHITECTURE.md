# Architecture Overview

## Layer boundaries

- Addin: command handlers, user interaction
- Core: business rules and use-case orchestration
- Client: API communication
- RevitAdapter: Revit-side extraction/writing behavior
- Infra: platform utilities and local runtime concerns

## API domains

- `/api/v1/bim/edms/*`
- `/api/v1/bim/schedules/*`
- `/api/v1/bim/site-logs/revit/*`

## Data quality rules

- Publish uses idempotency hash and partial-continue execution.
- Schedule ingestion uses staging then approve/reject.
- Writeback uses verified site-log rows only.
