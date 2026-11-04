ResultsAnalysisTerm02
![Platform](https://img.shields.io/badge/Platform-Excel%20(Microsoft%20365)%20%7C%20Windows-217346?style=flat-square&logo=microsoft-excel&logoColor=white)
![Language](https://img.shields.io/badge/Language-VBA-blue?style=flat-square)
![Domain](https://img.shields.io/badge/Domain-Educational%20Data%20Analytics-success?style=flat-square)
![Scope](https://img.shields.io/badge/Scope-Source%20modules%20only-lightgrey?style=flat-square)
ResultsAnalysisTerm02 is an Excel/VBA application for processing second-term (الفصل الثاني) results of secondary schools. It generates statistical analysis workbooks, compares terms, produces annual averages, and runs the preliminary stream-orientation workflow (التوجيه المسبق), starting from the official exports of the Algerian Ministry of National Education's Digitalization Platform (أرضية الرقمنة).
This repository contains the exported VBA modules (`.bas`) only. The macro-enabled application workbook, the source workbooks and all student data are not part of it (see Repository scope).
---
Table of contents
Features
How it works
Repository structure
Inputs and outputs
Getting started
Design notes
Requirements and compatibility
Validation status
Repository scope
Author
---
Features
1. Second-term analysis workbooks
`Build_All` reads the original workbooks from the `Source02` folder and generates new analysis workbooks. Each generated workbook contains:
Worksheet	Content
`Data`	The consolidated student records, stored as an Excel table.
`Analysis01` / `Analysis02`	Statistical analysis of the first and second terms.
`Compare`	First-term versus second-term comparison.
`Cart_Tec`	An individual student card, designed for display or projection.
The analysis sheets treat each class or pedagogical group (الأفواج التربوية) as a complete statistical population and report its mean, standard deviation and distribution indicators. Subjects are discovered from the headers of the `Data` table, so the engine does not depend on a fixed number of students or fixed subject-column positions.
2. Term-result processing
Split combined term results into separate first-term and second-term workbooks (`modSplitTermResults`).
Annual averages generated from the available term results (`modAverage02`).
Term comparison on the second-term source data (`modCompaire`).
Examination analysis (`modExamenAnalysis02`).
3. Preliminary orientation workflow
Coordinated by `Run_OrientationAll02`, in four stages:
Data transfer: eligible student data moves from `Source02` into `Term02Guidance.xlsx`.
Stream preferences: students' choices are read from `StreamPreferences02.xlsx` and linked to their records.
Orientation decisions: the configured rules are applied to the first- and second-term grades and the orientation subject groups (مجموعات التوجيه), then matched against the students' preferences (بطاقة الرغبات), producing the decision data and a general summary.
Results export: the final results are written to a separate workbook.
Student matching is validated before each stage proceeds: missing or ambiguous matches stop the process, whereas a valid match with empty preference fields is reported without being treated as an identity error. When run through the coordinator, the stages work in silent mode and end with one consolidated status message; run individually, each module reports on its own.
---
How it works
```
Source02\ (ministry exports) ──► Build_All ──► Data · Analysis01 · Analysis02 · Compare · Cart_Tec
                                                  (new workbooks, existing files never overwritten)

Source02\ ──► modTransfer02 ──► Term02Guidance.xlsx ◄── modStreamPreferences02 ◄── StreamPreferences02.xlsx
                                      │
                                modOrientation02 ──► modOrientationResults02 ──► final orientation workbook
                                      └──────── all four stages: Run_OrientationAll02 ────────┘
```
---
Repository structure
Module	Responsibility
`modBuildPanel02`	Builds the control panel and binds its action buttons.
`modCollector`	Collection-engine implementation retained in the project.
`modAnalysisWorkbookBuilderBuildAll02`	Entry point `Build_All`: coordinates the generation of the analysis workbooks.
`modAnalysisWorkbookBuilderData02`	Reads `Source02`, groups the records and creates the `Data` sheet.
`modAnalysisWorkbookBuilderAnalysis02`	Builds `Analysis01` and `Analysis02` from the `Data` table.
`modAnalysisWorkbookBuilderCompaire02`	Builds the `Compare` sheet.
`modAnalysisWorkbookBuilderCartTec02`	Builds the `Cart_Tec` student card.
`modAnalysisWorkbookBuilderCommon02`	Shared constants, text utilities, formatting helpers and data discovery for the six modules above.
`modCompaire`	Stand-alone term comparison (keeps its own `NewOutputPath` naming logic).
`modSplitTermResults`	Splits combined results into Term 01 and Term 02 workbooks.
`modAverage02`	Generates the annual-average workbook.
`modExamenAnalysis02`	Examination-results analysis.
`modTransfer02`	Transfers student data into `Term02Guidance.xlsx`.
`modStreamPreferences02`	Validates stream preferences and transfers them to the guidance workbook.
`modOrientation02`	Applies the orientation rules and prepares decisions and summary.
`modOrientationResults02`	Exports the final orientation results to a separate workbook.
`modOrientationAll02`	Orchestrates the four orientation stages (`Run_OrientationAll02`).
> **Naming note:** the spelling `Compaire` in `modCompaire` and `modAnalysisWorkbookBuilderCompaire02` is intentional: it matches the existing module and file names.
---
Inputs and outputs
Inputs
Keep these next to the main application workbook unless the configuration says otherwise.
Input	Used for
`Source02\`	Original second-term source workbooks (read-only input; never converted or moved).
`StreamPreferences02.xlsx`	Students' stream preferences.
`Term02Guidance.xlsx`	Intermediate workbook filled and processed by the orientation stages.
The sheet layouts and file-naming conventions of the sources must match what the VBA procedures expect.
Outputs
Analysis workbooks: created fresh on every run in the output folder beside the application workbook. If a file name already exists, a sequential suffix is added instead of overwriting. Each run records its status and writes a run log.
Split and annual-average workbooks: also use sequential file names.
Orientation results: saved in the folder `تحليل نتائج الفصل الثاني`, with one sheet `Litr` for literature-stream students and separate science-stream sheets (`Tecno1`, `Tecno2` and any other section found in the guidance data). Columns are auto-fitted and the sheets are protected without a password, so they can be unprotected for manual edits.
---
Getting started
Get the modules
```bash
   git clone https://github.com/SprstarSource/ResultsAnalysisTerm02.git
   ```
Import them into a workbook
Create a macro-enabled workbook (`.xlsm`) and open the VBA editor (`Alt + F11`).
Choose File ▸ Import File… and import the `.bas` modules.
Prepare the inputs: place the ministry exports in `Source02`, plus `StreamPreferences02.xlsx` and `Term02Guidance.xlsx` when you use the orientation workflow.
Generate the control panel: run the panel-building routine in `modBuildPanel02`.
Run an operation from the panel:
Second Term Analysis → `Build_All`
Term 02 Orientation → `Run_OrientationAll02`
Other operations can be run on their own, from their panel buttons or public entry points.
Enable macros if Excel shows a security warning and the workbook is trusted.
---
Design notes
Fresh output on every run. Workbooks are generated programmatically, without templates or reuse of earlier output files, so nothing stale or corrupted carries over.
Fast in-memory processing. Raw data is loaded into `Variant` arrays and joined with `Scripting.Dictionary` lookups instead of cell-by-cell worksheet access.
Self-contained modules, by choice. Some modules embed their own copy of the collection engine so that an operation does not depend on another module being present. The trade-off is duplicated logic that must be kept consistent when its rules change. For the full rationale see `ARCHITECTURE.txt`.
Dynamic discovery in generated workbooks. The analysis modules read subjects from the headers of the `Data` table rather than from hard-coded positions.
Interruption and error handling. The orientation workflow validates its inputs, reports consolidated results and handles `Esc` / `Ctrl + Break`; Excel settings are restored and the internal busy state reset after an interruption or failure.
---
Requirements and compatibility
Microsoft Excel for Windows with VBA enabled.
Microsoft 365 (or an Excel version supporting `MAXIFS`, `MINIFS`, `MODE.SNGL` and dynamic arrays), needed by the analysis workbooks.
A macro-enabled workbook containing the imported modules and the control panel.
The input folders and workbooks listed above.
Behaviour may vary with the Excel version, regional settings and the exact structure of the source files.
---
Validation status
The application has been run and tested by its developer on the developer's own computer, where the workflows behaved as expected. The repository holds only the exported source modules: it has no automated test suite and no runtime environment to reproduce the operations independently. A successful local run should not be read as a formal compatibility test across Excel versions or machine configurations.
---
Repository scope
This repository tracks the VBA source code and its history. It is not a place for student records or operational school data. Do not commit identifiable student information, source result workbooks, passwords, access tokens or any other confidential material.
Commit history documents how the modules and the architecture evolved.
---
Author
SprstarSource: github.com/SprstarSource
