# TadOcRevitBridge

`TadOcRevitBridge` is the Revit-side execution layer used by the Windows bridge (`TAD-Bridge-OC`) and the MCP/backend orchestrator (`TAD-BIM-Host-Runner`).

## Project Variants

This repository contains two project variants that produce the same `TadOcRevitBridge.dll` add-in, each targeting a different Revit generation:

| Variant | Folder | Target Framework | Revit versions |
|---|---|---|---|
| Multi-version (default) | `TadOcRevitBridge/` | `net8.0-windows` (2025/2026) or `net48` (2024) via `RevitVersion` property | 2024, 2025, 2026 |
| Revit 2024 standalone | `TadOcRevitBridge2024/` | `net48` (.NET Framework 4.8) hardcoded | 2024 only |

**Revit 2025 and 2026** — use `TadOcRevitBridge/TadOcRevitBridge.csproj` with `RevitVersion=2025` or `RevitVersion=2026`. Targets `net8.0-windows`.

**Revit 2024** — you have two options:
- Use `TadOcRevitBridge/TadOcRevitBridge.csproj` with `-p:RevitVersion=2024` (targets `net48`, requires .NET Framework 4.8 developer pack).
- Use `TadOcRevitBridge2024/TadOcRevitBridge2024.csproj` — a standalone project hardcoded for `net48` and the Revit 2024 API DLLs at `C:\Program Files\Autodesk\Revit 2024\`. This is the simpler option when you only need to build for Revit 2024 and want to open a single self-contained project in Visual Studio.

Both variants share identical C# source files and produce a `TadOcRevitBridge.dll` with the same tool contracts and JSON protocol.

The add-in stays intentionally narrow:

- Revit API execution happens here.
- Job orchestration stays outside the add-in.
- Requests are delivered through the existing file queue pattern.

## Current Architecture

- Entry point: `Autodesk.Revit.UI.IExternalApplication` implemented by `TadOcRevitBridge.App`
- Execution trigger: `UIControlledApplication.Idling`
- IPC pattern:
  - inbox: `D:\TAD\revit-bridge\inbox`
  - outbox: `D:\TAD\revit-bridge\outbox`
  - archive: `D:\TAD\revit-bridge\archive`
  - health file: `D:\TAD\revit-bridge\outbox\revit-addin-alive.json`
- Dispatch model: one JSON request file is processed per idle tick, then archived
- Result model: JSON-friendly response envelope written to `<jobId>.result.json`

There is no `ExternalEvent` pattern in the current add-in. The live execution model remains the original queued file/polling design, extended into a tool dispatcher instead of replacing it.

## Supported Actions

Implemented:

- `revit_create_wall`
- `revit_session_status`
- `revit_open_cloud_model`
- `revit_list_3d_views`
- `revit_export_nwc`

### Request Envelope

All requests use the same envelope shape:

```json
{
  "jobId": "job-123",
  "tool": "revit_session_status",
  "createdAt": "2026-04-07T10:00:00Z",
  "status": "queued",
  "payload": {}
}
```

### Response Envelope

All responses include:

```json
{
  "ok": true,
  "source": "revit-addin",
  "status": "completed",
  "jobId": "job-123",
  "tool": "revit_session_status",
  "revitVersion": "2025",
  "timeUtc": "2026-04-07T10:00:01.0000000Z"
}
```

Failures add:

```json
{
  "error": {
    "code": "no_active_document",
    "message": "Tool 'revit_export_nwc' requires an active Revit document."
  }
}
```

## Tool Contracts

### `revit_create_wall`

Preserved from the original add-in behavior.

Request payload:

```json
{
  "start": { "x": 0.0, "y": 0.0, "z": 0.0 },
  "end": { "x": 3.0, "y": 0.0, "z": 0.0 },
  "height": 3.0,
  "level": "Level 1",
  "wallType": "Basic Wall"
}
```

### `revit_session_status`

Reports the live Revit session state visible to the add-in.

Example response:

```json
{
  "ok": true,
  "tool": "revit_session_status",
  "revitRunning": true,
  "revitVersion": "2025",
  "activeDocument": {
    "isOpen": true,
    "isActive": true,
    "title": "ModelName",
    "isCloudModel": true,
    "projectGuid": "11111111-1111-1111-1111-111111111111",
    "modelGuid": "22222222-2222-2222-2222-222222222222",
    "region": "US"
  }
}
```

If no active document exists, `activeDocument.isOpen` is `false` and cloud metadata fields are `null`.

### `revit_open_cloud_model`

Opens a Revit cloud model using `ModelPathUtils.ConvertCloudGUIDsToCloudPath(...)` plus `Application.OpenDocumentFile(...)`.

Request payload:

```json
{
  "region": "US",
  "projectGuid": "11111111-1111-1111-1111-111111111111",
  "modelGuid": "22222222-2222-2222-2222-222222222222",
  "openInUi": false,
  "audit": false,
  "worksets": {
    "mode": "default"
  },
  "cloudOpenConflictPolicy": "use_default"
}
```

Supported workset modes:

- `default`
- `open_all`
- `close_all`
- `open_last_viewed`

Supported cloud conflict policies:

- `use_default`
- `discard_local_changes_and_open_latest_version`
- `keep_local_changes`
- `detach_from_central`
- `cancel`

Structured failure codes include:

- `unauthenticated_user`
- `unauthorized_access`
- `invalid_identifiers`
- `wrong_region`
- `communication_failure`
- `unsupported_context`

Important limitation:

- `openInUi=false` is fully implemented and opens the document in the background.
- `openInUi=true` is not implemented in this queue path yet, because the add-in currently executes from `Idling` and this repo does not yet have a safe UI-activation path wired for that request.

### `revit_list_3d_views`

Lists `View3D` elements from the active document.

Request payload:

```json
{
  "onlyExportable": true,
  "excludeTemplates": true
}
```

Example response:

```json
{
  "ok": true,
  "tool": "revit_list_3d_views",
  "documentTitle": "ModelName",
  "navisworksExporterAvailable": true,
  "views": [
    {
      "id": "12345",
      "name": "{3D}",
      "isTemplate": false,
      "canExport": true,
      "isPerspective": false
    }
  ]
}
```

### `revit_export_nwc`

Exports Navisworks-compatible `.nwc` output from the active document.

Request payload:

```json
{
  "viewNames": ["{3D}"],
  "outputPath": "C:\\Exports\\MyModel.nwc",
  "exportScope": "selected_views"
}
```

Supported export scopes:

- `selected_views`
- `model`

Notes:

- Revit exposes a single `ViewId` on `NavisworksExportOptions` when exporting a view-scoped NWC.
- Because of that, multiple requested view names are exported as multiple `.nwc` files derived from the requested base output path.
- The response returns the concrete output path for each exported view.

Structured failure codes include:

- `no_active_document`
- `requested_view_not_found`
- `navisworks_exporter_unavailable`
- `invalid_output_path`
- `export_failed`

## Version Strategy

The add-in keeps one logical tool contract across Revit 2024, 2025, and 2026.

Build strategy:

- Revit 2024 builds against `net48` (.NET Framework 4.8)
- Revit 2025 builds against `net8.0-windows`
- Revit 2026 builds against `net8.0-windows`
- Command names and JSON contracts do not change by Revit version

Multi-version project (`TadOcRevitBridge/`):

- project file: `TadOcRevitBridge/TadOcRevitBridge.csproj`
- required build property: `RevitVersion=2024|2025|2026`
- optional override: `RevitInstallDir=C:\Program Files\Autodesk\Revit 2025`

Revit 2024 standalone project (`TadOcRevitBridge2024/`):

- project file: `TadOcRevitBridge2024/TadOcRevitBridge2024.csproj`
- no extra build properties required
- Revit API DLLs expected at `C:\Program Files\Autodesk\Revit 2024\` (override with `RevitInstallDir`)

## Bridge Expectations

The bridge should:

- write one request JSON file per job into the inbox
- wait for `<jobId>.result.json` in the outbox
- relay the JSON response as-is to the MCP/backend
- treat the add-in as the source of truth for Revit-specific execution outcomes

The bridge should not reinterpret successful/failed Revit operations beyond transport concerns.

## Prerequisites

- Revit must already be running with this add-in loaded
- For cloud-model open:
  - the Revit user must be signed in to Autodesk
  - the signed-in user must have access to the requested Autodesk Docs/ACC model
- For view listing and NWC export:
  - an active document is required
- For NWC export:
  - a compatible Navisworks exporter must be available in the current Revit session

## Build Instructions

Run these on Windows from the repo root.

Build prerequisites:

- Revit installed for the target year, or `RevitInstallDir` pointed at the matching API DLLs
- .NET Framework 4.8 developer pack for Revit 2024 builds
- .NET 8 SDK for Revit 2025 and 2026 builds
- Visual Studio 2022 17.8 or later if you are building the .NET 8 variants inside Visual Studio

### Revit 2024 (multi-version project)

```powershell
dotnet build .\TadOcRevitBridge\TadOcRevitBridge.csproj -c Release -p:RevitVersion=2024
```

Output: `TadOcRevitBridge\bin\Release\Revit2024\`

### Revit 2024 (standalone project)

```powershell
dotnet build .\TadOcRevitBridge2024\TadOcRevitBridge2024.csproj -c Release
```

Output: `TadOcRevitBridge2024\bin\Release\`

To override the Revit install path:

```powershell
dotnet build .\TadOcRevitBridge2024\TadOcRevitBridge2024.csproj -c Release -p:RevitInstallDir="D:\Autodesk\Revit 2024"
```

### Revit 2025

```powershell
dotnet build .\TadOcRevitBridge\TadOcRevitBridge.csproj -c Release -p:RevitVersion=2025
```

Output: `TadOcRevitBridge\bin\Release\Revit2025\`

### Revit 2026

```powershell
dotnet build .\TadOcRevitBridge\TadOcRevitBridge.csproj -c Release -p:RevitVersion=2026
```

Output: `TadOcRevitBridge\bin\Release\Revit2026\`

### Build all targeted versions (multi-version project)

```powershell
dotnet msbuild .\TadOcRevitBridge\TadOcRevitBridge.csproj /t:BuildAllRevitVersions /p:Configuration=Release
```

Expected output folders:

- `TadOcRevitBridge\bin\Release\Revit2024\`
- `TadOcRevitBridge\bin\Release\Revit2025\`
- `TadOcRevitBridge\bin\Release\Revit2026\`

## Deploy Instructions

1. Build the version that matches the installed Revit major version.
2. Copy `TadOcRevitBridge.dll` and any copied dependency DLLs from the matching output folder to your deployment location.
3. Copy the `.addin` manifest to the matching Revit addins directory:
   - For Revit 2024 (standalone project): use `TadOcRevitBridge2024\TadOcRevitBridge2024.addin` → `%AppData%\Autodesk\Revit\Addins\2024\`
   - For Revit 2024/2025/2026 (multi-version project): create or update the manifest manually under `%AppData%\Autodesk\Revit\Addins\<year>\`
4. Point the manifest `Assembly` path at the built `TadOcRevitBridge.dll`.
5. Restart Revit and confirm `revit-addin-alive.json` appears in the outbox.

## Known Limitations

- `revit_open_cloud_model` currently implements background open only (`openInUi=false`).
- `revit_list_3d_views` and `revit_export_nwc` intentionally operate on the active document only.
- This repo cannot validate Revit API behavior without a Windows machine that has the matching Revit version and exporter installed.
