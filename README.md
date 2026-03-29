# Revit LOD3 Exporter

This Revit plugin exports building models to CityJSON LOD3 format and supports LEED-oriented retrofit analysis.

## Features
- Export Revit models to CityJSON LOD3 (`Export LOD3`)
- Import retrofit parameters from CSV and apply deterministic or AI-based upgrades (`Import CSV`)
- Generate retrofit reports and write results to Revit parameters
- Export window/door retrofit advice as TXT

## Getting Started

### Clone the Repository
Clone this repository from GitHub:

```
git clone https://github.com/yueyingwu965-prog/RevitLOD3Exporter.git
```

Open the solution in Visual Studio.

## Installation
1. **Build the project** in Visual Studio (targeting Revit API, .NET Framework).
2. **Copy the compiled DLL** to your Revit Addins folder (e.g., `%APPDATA%\Autodesk\REVIT\Addins\2023`).
3. **Add a .addin manifest** pointing to the DLL if needed.

## Usage
1. Open your Revit model.
2. On the `LOD Tools` ribbon tab, use:
   - **Export LOD3**: Export CityJSON LOD3 from a LOD2 CityJSON file and Revit model.
   - **Import CSV**: Import retrofit parameters from a CSV file and apply to the model. Choose between deterministic rules or AI-based analysis (requires API key for AI).
   - **Retrofit Apply**: Apply upgrades and export window/door retrofit TXT report.

## Requirements
- Autodesk Revit (tested on 2023+)
- .NET Framework (as required by your Revit version)
- [Optional] API key for AI features: set `DEEPSEEK_API_KEY`, `OPENAI_API_KEY`, or `GEMINI_API_KEY` as environment variables for AI-powered analysis.

## Notes
- Input/output file dialogs are shown for CityJSON and CSV files.
- Reports are written to Revit project parameters and optionally exported as TXT.
- For more details, see code comments in `ExportLOD3.cs`, `ImportRetrofitCSV.cs`, and `RetrofitCommand.cs`.
