# MDR Revit Plugin

This folder contains the Revit plugin workspace for the MDR BIM integration.

## Scope

- Revit to MDR EDMS publish (sheets and files)
- Revit to MDR schedule ingest (MTO and Equipment)
- MDR to Revit site-log writeback (verified data only)

## Solution projects

- `src/Mdr.Revit.Addin`: command and UI entry points
- `src/Mdr.Revit.Core`: domain models and use-cases
- `src/Mdr.Revit.Client`: HTTP clients for MDR APIs
- `src/Mdr.Revit.RevitAdapter`: Revit extraction/write adapters
- `src/Mdr.Revit.Infra`: config, logging, utility helpers
- `tests/*`: unit test projects

## Current runtime status

- Addin app orchestrates commands through `src/Mdr.Revit.Addin/App.cs`.
- `Login`, `Publish Selected Sheets`, `Push Schedules`, `Sync Contractor Reports`, `Google Sheets Sync`, `Smart Numbering`, `Check Updates`, and `Settings` command handlers are wired.
- EDMS publish command executes: login -> extract -> export -> publish -> retry failed items.

## Step-by-step execution baseline

1. Build and config baseline
2. Core models/contracts
3. EDMS publish flow (login + multipart publish + failed-item retry)
4. Schedule ingest flow
5. Site-log writeback flow
6. Packaging and rollout
