# POAP Builder — Local Windows Portal

Turn Microsoft Project MPP, Excel, CSV and TXT project plans into an editable executive Plan on a Page (POAP) PowerPoint.

## Run locally on Windows
1. Install **Python 3.11 or 3.12** and select **Add Python to PATH** during installation.
2. For native `.mpp` files, install a **64-bit Java 17+ JDK/runtime**.
3. Clone/download this repository.
4. Double-click `start.bat`.
5. The portal opens at `http://127.0.0.1:8000`.

On first run, `start.bat` creates a private `.venv` and installs the Python dependencies.

## Supported inputs
- Microsoft Project `.mpp` through MPXJ + Java
- Excel `.xlsx` / `.xlsm`
- `.csv`
- delimited or line-based `.txt`

## Current POAP output
- Project dates and task count
- Workstreams/phases
- Key milestones
- Rule-based RAG status
- Completion metric
- Dependencies count
- Risks/attention flags
- Executive summary
- Editable `.pptx` output

## Data handling
Uploads and generated files remain in the local `uploads/` and `outputs/` folders. No cloud service or AI service is required.

## Next roadmap
- In-browser POAP editor before export
- Template selector / corporate branding
- PDF export
- Configurable RAG rules
- Optional AI executive summarisation when an approved model is available