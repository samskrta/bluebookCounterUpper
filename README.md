## Blue Book Counter Upper

Counts unique Blue Book quotes per technician from the Excel export.
Also flags rows whose Labor was manually modified (detected via cell comments)
and summarizes whether changes trend up/down per technician.
Requires download of Blue Book report from https://account.mypartshelp.com/company

### Setup

```bash
pip3 install -r requirements.txt
```

### Run

```bash
python3 bluebook_count.py --file data/BlueBook_Report_10_01_2025.xlsx --sheet ALL
```

Output is CSV to stdout with a TOTAL row. Optionally write a counts CSV file:

```bash
python3 bluebook_count.py --file data/BlueBook_Report_10_01_2025.xlsx --csv out/bluebook_counts.csv
```

Additionally, write per-line modification flags and a per-tech summary:

```bash
python3 bluebook_count.py \
  --file data/BlueBook_Report_10_01_2025.xlsx \
  --mods-csv out/bluebook_mods.csv \
  --mods-summary-csv out/bluebook_mods_summary.csv
```

Notes:
- Labor changes are identified by the presence of a comment on the Labor cell.
- Direction (up/down/even) is inferred from amounts mentioned in the comment
  (e.g., "from $149 to $179"), or compared to the Labor cell value when only
  one amount is present. If not inferable, direction is marked "unknown".

### Distribute to non-technical users (macOS)

1) Build the app bundle:

```bash
chmod +x build_macos.sh
./build_macos.sh
```

2) Share the app found at `dist/BlueBook Counter Upper.app` (or the CLI at `dist/BlueBook Counter Upper`). On first run, Gatekeeper may warn you; right‑click → Open to allow.

How users run it:
- Double-click the app, choose the `.xlsx` file when prompted, and it will save a CSV next to the source file (suffix `_counts.csv`).
- Alternatively, run from Terminal for advanced options:

```bash
./dist/BlueBook\ Counter\ Upper --gui
```


### Distribute to non-technical users (Windows)

1) Build the exe:

```bat
build_windows.bat
```

2) Share `dist\\BlueBook Counter Upper.exe`.

How users run it:
- Double‑click the `.exe`, choose the `.xlsx` file when prompted, and it will save a CSV next to the source file (suffix `_counts.csv`) and auto-open it.
- Alternatively, from Command Prompt:

```bat
"dist\\BlueBook Counter Upper.exe" --gui
```

### Build Windows exe via GitHub Actions

Manual run:
- In GitHub, go to Actions → "Build Windows exe" → Run workflow.
- Download the artifact named `BlueBook-Counter-Upper-windows` which contains `BlueBook Counter Upper.exe`.

Release build:
- Tag a commit, e.g.:

```bash
git tag v1.0.0 && git push origin v1.0.0
```

- The workflow will attach the exe to the new GitHub Release.

### Download prebuilt apps (for non‑technical users)

- Go to the repo’s Releases page and download:
  - `BlueBook-Counter-Upper-macOS.zip` for macOS
  - `BlueBook Counter Upper.exe` for Windows
- On macOS, unzip and drag the app to Applications, right‑click → Open on first run.




