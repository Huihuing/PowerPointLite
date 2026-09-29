# Agent Instructions

`AGENTS.md` is the shared source of truth for Claude Code, OpenAI Codex, ChatGPT coding sessions, and other coding agents working on this repository.

Do not reconstruct the project from chat history when Git and repository files already contain the needed state.

## Project Goal

PowerPointLite currently has a stable Windows PPTX/PPTM Viewer on `main` and an experimental multi-format document/editor foundation on `feature/office-foundation`.

Long-term target:

```text
PPTX / PPTM  read / create / edit / save / present / print
DOCX / DOCM  read / create / edit / save / print
XLSX / XLSM  read / create / edit / save / print
HWPX         read / create / edit / save
HWP 5.x      read first; writer only after feasibility/safety review
ODT / ODS / ODP read / create / edit / save
PDF          export
```

The project must remain an independent implementation. Do not copy proprietary application code, DLLs, templates, logos, icons, or UI artwork.

## Source of Truth

Use sources in this order:

1. current code
2. current branch and `git status`
3. recent `git log`
4. this file
5. format/architecture/legal docs under `docs/`
6. Issues/PRs only when relevant
7. old chat/session context only when repository state is insufficient

Do not treat a previous Claude/ChatGPT/Codex session as the authoritative state.

When moving between home and academy PCs, recover with:

```text
AGENTS.md
→ CLAUDE.md when using Claude Code
→ git status
→ current branch
→ recent git log
→ files relevant to the current task
```

Do not rescan the whole repository when these are sufficient.

## Architecture

Core internal models:

```text
PresentationDocument
TextDocument
SpreadsheetDocument
```

Package layers:

```text
OOXML/OPC: src/OpcPackage.cs, src/OpcPreservation.cs
ODF:       src/OdfPackage.cs
```

Format implementations are separated from UI where possible. UI should edit internal models rather than directly mutating ZIP/XML package contents.

Important architectural rule:

```text
Reader → internal model → Renderer/Editor
Editor → internal model → Writer
```

Never narrow an arbitrary external document into a smaller model and silently overwrite the original. Unsupported/unknown content must be preserved or editing must be denied.

## Important Files

Viewer:

```text
src/MainForm.Part01.cs ... MainForm.Part05.cs
src/InternalPptxRenderer.Part01.cs ... Part08.cs
src/PresenterView.cs
src/Printing.cs
```

Workspace / UI:

```text
src/UiTheme.cs
src/OfficeWorkspaceDialog.cs
src/MainFormWorkspaceToolbar.cs
src/EditorIntegration.cs
```

PPTX editor/writer:

```text
src/PresentationModel.cs
src/PresentationEditSession.cs
src/PptxEditableReader.cs
src/PptxWriter.cs
src/PptxCompatibilityAnalyzer.cs
src/AdvancedPresentationEditorForm.cs
src/AdvancedEditorThumbnailExtension.cs
src/AdvancedEditorTableExtension.cs
```

Text documents:

```text
src/TextDocumentModel.cs
src/TextDocumentEditSession.cs
src/UnifiedTextDocumentEditorForm.cs
src/DocxReader.cs
src/DocxWriter.cs
src/HwpxCore.cs
src/OdtCore.cs
```

Spreadsheets:

```text
src/SpreadsheetCore.cs
src/SpreadsheetEditorForm.cs
src/OdsCore.cs
src/OdsSpreadsheetEditorForm.cs
```

ODP:

```text
src/OdpCore.cs
src/OdpPresentationEditorForm.cs
```

HWP read-only:

```text
src/CompoundFileReader.cs
src/HwpReader.cs
src/HwpReadOnlyViewerForm.cs
```

PDF / conversion:

```text
src/PdfExport.cs
src/DocumentPdfExport.cs
src/DocumentConversion.cs
```

Fonts / licensing:

```text
src/FontLicensing.cs
src/FontLicenseService.cs
src/FontBundlePolicy.cs
src/FontLicenseInspectorForm.cs
src/FontLicenseDiagnostics.cs
```

## Build / Test / Run

Target environment:

```text
Windows 10/11
.NET Framework 4.x csc.exe
WinForms
No .NET 8 SDK requirement
No mandatory NuGet dependency
```

Build:

```bat
BUILD_EXE.cmd
```

Pre-merge automated gate:

```bat
RUN_PREMERGE_CHECKS.cmd
```

Format self-tests:

```bat
RUN_ALL_FORMAT_SELFTESTS.cmd
RUN_WRITER_SELFTEST.cmd
RUN_DOCX_SELFTEST.cmd
RUN_XLSX_SELFTEST.cmd
RUN_HWPX_SELFTEST.cmd
RUN_HWP_PARSER_SELFTEST.cmd
RUN_ODT_SELFTEST.cmd
RUN_ODS_SELFTEST.cmd
RUN_ODP_SELFTEST.cmd
RUN_PDF_SELFTEST.cmd
RUN_CONVERSION_SELFTEST.cmd
RUN_FONT_LICENSE_SELFTEST.cmd
```

Structural self-tests do not prove real Microsoft Office / LibreOffice / Hancom interoperability.

If the current environment cannot run Windows `csc.exe`, do not claim `BUILD SUCCESS`.

Known compatibility hazards:

```text
GUI Timer → use System.Windows.Forms.Timer explicitly
Graphics.DrawImage → use .NET Framework-compatible overloads; Rectangle.Round when needed
Avoid modern C#/.NET-only syntax or APIs that Framework csc.exe cannot compile
```

## Development Rules

Priority order:

```text
user document safety
→ copyright/trademark/font-license safety
→ existing PPTX Viewer regression safety
→ Windows build compatibility
→ new features
```

Git rules:

- work in meaningful units
- do not create a commit for every tiny edit
- commit message format: `한글 - English`
- do not merge experimental office/editor work to `main` before Windows build and regression checks
- never commit API keys, tokens, passwords, or private credentials
- Git history is the long-term work log; do not maintain duplicate chronological logs in Markdown

Viewer behavior that must not regress:

- Fit uses the actual `SplitContainer.Panel2` viewport
- slide-area wheel = previous/next
- Ctrl+wheel = zoom
- thumbnail wheel = normal list scroll
- F5 / Shift+F5 slideshow
- Presenter View / Notes / Print
- startup crash logging

## Copyright / Trademark / Font Rules

Read when relevant:

```text
docs/LEGAL_ASSET_POLICY.md
THIRD_PARTY_NOTICES.md
licenses/README.md
```

Do not bundle or copy:

- Microsoft Office executables/DLLs
- Hancom executables/DLLs
- Microsoft/Hancom official logos or icons
- copied Office/Hancom UI artwork
- commercial templates/clipart without explicit rights
- third-party font files without an explicit redistribution/app-embedding license
- arbitrary third-party documents as repository fixtures

Font policy:

- system-installed font families may be used for rendering/editing
- do not copy system font binaries into the app
- store font family names in documents by default
- embedding is deny-by-default
- `OS/2.fsType` is advisory metadata, not the final license grant
- actual license text takes priority
- application bundling cannot be authorized by fsType alone
- current raster PDF export does not embed source TTF/OTF binaries

## UI Rules

Use the project-owned dark/flat visual language from `src/UiTheme.cs`.

- keep UI clean and lightweight
- do not clone Microsoft Office/Hancom Ribbon appearance
- do not use proprietary product icons/logos
- common UI concepts such as toolbar, sidebar, grid, canvas, property panel are fine
- prefer consistent spacing, alignment, typography, and hover/focus states
- retain keyboard accessibility and test common DPI levels when Windows validation is available

See `docs/UI_DESIGN_GUIDE.md` for durable design details.

## External Services

This repository does not require Render, Vercel, Aiven, Slack, Notion, or other external services for ordinary local development.

Default tools:

```text
local files
Git
terminal/compiler
```

Use an external service only when the current task actually depends on that service.

## MCP Usage Rules

Do not enumerate or probe every registered MCP just because it exists.

Before using MCP, check:

1. Is the information/action actually required for the current task?
2. Can local files, Git, or the terminal answer it?
3. Will the MCP result change the implementation decision?

If not, do not call it.

Examples:

```text
code edit/build → local files + Git + terminal
GitHub remote operation explicitly required → GitHub
production incident → only the relevant deployment/database MCP
Slack request → Slack only when the user actually asks about Slack
```

## Subagent Usage Rules

Use subagents only when isolated context provides real value, such as:

- broad multi-file investigation
- large log analysis
- specialized security review
- independent large research tracks

Do not use subagents for:

- one-file edits
- short searches
- routine Git operations
- small UI/CSS fixes
- tasks already well understood in the main context

Goal: keep the main context clean, not maximize agent count.

No project subagent is required by default. Add `.claude/agents/` entries only when a repeated workflow justifies them.

## Skill Usage Rules

Create skills only for genuinely repeated procedures.

Do not create skills just to wrap one-off commands or ordinary development steps.

For a user-triggered skill, consider:

```yaml
disable-model-invocation: true
```

so the model does not invoke it opportunistically.

No project skill is required by default. Add `.claude/skills/` only after a repeated workflow is clear.

## Documentation Rules

Root working documents:

```text
README.md   human-facing project overview
AGENTS.md   shared AI coding source of truth
CLAUDE.md   short Claude Code entry point
```

Do not create a persistent `PROGRESS.md`.

Keep actual project knowledge in `docs/`, for example:

```text
docs/ARCHITECTURE_ROADMAP.md
docs/EDITOR_FOUNDATION.md
docs/HWPX_FOUNDATION.md
docs/HWP_FOUNDATION.md
docs/ODF_FOUNDATION.md
docs/PDF_EXPORT_FOUNDATION.md
docs/LEGAL_ASSET_POLICY.md
docs/UI_DESIGN_GUIDE.md
```

These are durable architecture/format/legal documents, not session handoff logs.

Do not accumulate:

- daily status logs
- commit SHA diaries
- solved-error journals
- duplicated TODO files
- old AI handoff documents

Use Git history for historical work reconstruction.

## Current State

### Completed

- stable `main` PPTX/PPTM Viewer foundation
- multi-format workspace foundation on `feature/office-foundation`
- PPTX Writer + advanced experimental Editor
- DOCX Reader/Writer/Editor foundation
- XLSX Reader/Writer/Editor foundation
- HWPX Reader/Writer/Editor foundation
- HWP 5.x read-only parser foundation
- ODT/ODS/ODP Reader/Writer/Editor foundations
- raster PDF export foundation without font-binary embedding
- cross-format conversion foundation with safety guards
- OpenType/TrueType fsType reader and conservative font-license service
- font-license inspector and synthetic licensing-policy self-test
- project-owned dark/flat Workspace UI
- legal/asset/font policy documents

### Current

- make `feature/office-foundation` compile cleanly with Windows Framework `csc.exe`
- keep all structural self-tests consistent with current source
- reduce obvious .NET Framework compatibility hazards before Windows build
- strengthen deny-by-default preservation/safety paths
- keep UI consistent while expanding editors

### Next

1. run `BUILD_EXE.cmd` on Windows and fix the first compiler error until clean
2. run `RUN_PREMERGE_CHECKS.cmd`
3. run Viewer regression checks
4. validate PPTX/DOCX/XLSX against PowerPoint/Word/Excel or LibreOffice where available
5. validate HWPX against Hancom
6. validate ODT/ODS/ODP against LibreOffice
7. visually verify/print PDF output
8. expand unknown-part preservation before allowing arbitrary external-file editing
9. continue richer DOCX/XLSX/HWP support only after the build gate is healthy

### Blocked

- current non-Windows agent environment cannot prove Windows Framework `csc.exe` build success
- real Office/LibreOffice/Hancom interoperability requires those applications or a suitable Windows validation machine
- arbitrary external-document editing remains intentionally restricted until preservation coverage is broader

Do not mark these blockers as solved without actual verification.
