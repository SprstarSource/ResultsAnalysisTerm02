ResultsAnalysisTerm02
Excel VBA system for second-term academic results processing, statistical analysis, comparison, and school-orientation workflows.
This repository contains the VBA modules that power an Excel-based application for processing secondary-school results for the second term. The system automates the preparation of analysis workbooks, term-result operations, and the student-orientation workflow.
Repository: https://github.com/SprstarSource/ResultsAnalysisTerm02
---
1. Main capabilities
A. Second-term analysis workbook generation
The analysis workflow reads the original source workbooks from `Source02` and generates new analysis workbooks. It does not require a pre-generated `Term02Results.xlsx` file.
The generated workbooks include:
A consolidated `Data` worksheet containing the collected student records.
`Analysis01` and `Analysis02` worksheets for first-term and second-term analysis.
A `Compare` worksheet comparing first-term and second-term performance.
A `Cart_Tec` worksheet designed to display an individual student card, including information intended for projection.
The analysis engine discovers subjects and relevant data from the generated `Data` table rather than relying on a fixed student-row count or fixed subject-column positions.
B. Term-result processing
The application includes dedicated operations to:
Split combined term results into separate first-term and second-term workbooks.
Generate annual averages from the available term results.
Compare student performance between terms.
The term-splitting operations are consolidated in `modSplitTermResults`.
C. School-orientation workflow
The orientation process is coordinated by `Run_OrientationAll02` and consists of four stages:
Data transfer: transfer eligible student data from `Source02` into `Term02Guidance.xlsx`.
Stream preferences: read students' choices from `StreamPreferences02.xlsx` and transfer them to the guidance workbook.
Orientation decisions: apply the configured orientation rules and produce the decision data and general summary.
Results export: create a separate workbook containing the final orientation results.
The workflow validates student matching before proceeding. Missing or ambiguous matches stop the relevant process; valid matches with empty preference fields can be reported without being treated as identity-matching failures.
When the complete orientation workflow is run through its coordinator, the stages operate in silent mode and produce a consolidated final status message. Individual modules retain their own reporting when run separately.
D. Examination analysis
`modExamenAnalysis02` provides the examination-analysis functionality used by the application.
---
2. Repository structure
The repository contains exported VBA modules (`.bas`). The main Excel application workbook (`.xlsm`), source workbooks, student data, and generated result files are not stored here.
Module	Responsibility
`modBuildPanel02`	Builds and manages the application's control panel and its operation buttons.
`modCollector`	Contains the collection-engine implementation retained in the project.
`modBiuldAnalysisfilesBuildAll02`	Main entry point for generating the second-term analysis workbooks; coordinates the data, analysis, comparison, and student-card modules.
`modBiuldAnalysisfilesData02`	Reads the `Source02` folder, collects and groups source records, and creates the `Data` worksheet in newly generated workbooks.
`modBiuldAnalysisfilesAnalysis02`	Builds the first-term and second-term analysis worksheets from the `Data` table.
`modBiuldAnalysisfilesCompare02`	Builds the comparison worksheet for first-term versus second-term results.
`modBiuldAnalysisfilesCartTec02`	Builds the `Cart_Tec` student-card worksheet for display or projection.
`modBiuldAnalysisfilesCommon02`	Provides shared constants, text utilities, formatting helpers, and data-discovery routines for the analysis-generation modules.
`modCompaire`	Performs the separate term-comparison operation using the second-term source data.
`modSplitTermResults`	Splits combined term-result source files into separate term-result workbooks.
`modAverage02`	Generates the annual-average workbook.
`modExamenAnalysis02`	Performs examination-results analysis.
`modTransfer02`	Transfers student data into `Term02Guidance.xlsx` for the orientation process.
`modStreamPreferences02`	Validates and transfers student stream preferences into the guidance workbook.
`modOrientation02`	Applies the orientation rules and prepares the orientation decisions and summary.
`modOrientationAll02`	Coordinates the complete four-stage orientation workflow.
`modOrientationResults02`	Exports the final orientation results into a separate workbook.
> **Naming note:** Some existing module names contain the spelling `Biuld`. The table intentionally preserves the actual filenames and VBA module identifiers in the repository.
---
3. Inputs and working files
The following files and folders are expected by the relevant workflows. They are working data and should be maintained locally alongside the main application workbook unless the application configuration specifies otherwise.
Input	Used for
`Source02`	Original second-term source workbooks. Used by the analysis-generation and related term-processing workflows.
`StreamPreferences02.xlsx`	Students' stream preferences for the orientation workflow.
`Term02Guidance.xlsx`	Intermediate guidance workbook populated and processed by the orientation stages.
The source workbooks in `Source02` are treated as input data. The analysis-generation process reads them and creates new output workbooks rather than converting or moving the source files.
The exact worksheet layouts and required source-file naming conventions must match the structures expected by the VBA procedures.
---
4. Outputs
Analysis generation
The main analysis-generation workflow creates new Excel workbooks in the configured output directory located alongside the main application workbook. It produces the analysis data and report worksheets described above.
The workflow preserves earlier generated files. If an output filename already exists, a sequential suffix is used rather than overwriting the existing workbook.
The process also records execution status and writes a run log in the output location.
Term processing
`modSplitTermResults` and `modAverage02` generate their output workbooks using sequential filenames so that previous outputs are not overwritten.
`modCompaire` retains its own `NewOutputPath` naming logic.
Orientation results
The final orientation workbook is saved in:
`تحليل نتائج الفصل الثاني`
The exported workbook contains:
`Litr` for literature-stream students.
Separate science-stream worksheets, such as `Tecno1`, `Tecno2`, and other sections found in the guidance data.
The export applies automatic column sizing and worksheet protection. The worksheets are protected without a password so that users can unprotect them when manual editing is required.
---
5. Running the application
Open the main macro-enabled Excel application workbook (`.xlsm`).
Enable macros if Excel displays a security warning and the workbook is trusted.
Ensure that the required input folders and workbooks are present.
Use the application's control panel to run the required operation.
For the complete second-term analysis-generation process, use the Second Term Analysis control-panel action, which invokes `Build_All`.
For the complete orientation process, use the Term 02 Orientation action, which invokes `Run_OrientationAll02`.
Individual operations can also be run separately through their corresponding control-panel actions or public VBA entry points, where applicable.
---
6. Technical notes and design decisions
Fresh output generation
The application follows a fresh-generation approach for its generated analysis workbooks. It creates new output workbooks instead of repeatedly clearing and reusing predefined output files.
Existing output files are preserved through sequential filename generation.
Independent collection logic
Some modules contain their own embedded copy of the data-collection engine. This is an intentional reliability and deployment decision: an operation should not become dependent on a separate shared source module being present or successfully generated at runtime.
The trade-off is that duplicated collection logic must be kept consistent when its rules are changed.
Dynamic data discovery
The analysis-generation modules use the `Data` table and discover subjects from its headers. This reduces dependence on fixed row counts and hard-coded subject-column positions.
Error handling and interruption
The orientation workflow includes validation, consolidated reporting, and handling for user interruption (including `Esc` and `Ctrl+Break`). The modules are designed to restore Excel application settings and reset their internal busy state after interruption or failure.
Worksheet protection and formatting
Generated worksheets may be protected against accidental changes. Protection is applied without a password where specified by the relevant export workflow. Automatic column sizing is used in generated outputs where implemented.
---
7. Requirements and compatibility
Microsoft Excel for Windows with VBA enabled.
A macro-enabled application workbook (`.xlsm`) containing the imported modules and control panel.
Microsoft 365 / a compatible Excel version for analysis workbooks that use modern worksheet functions such as `MAXIFS`, `MINIFS`, `MODE.SNGL`, and dynamic-array functionality.
The required local input folders and workbooks described in this document.
Compatibility may vary with Excel version, regional settings, workbook structure, and the exact source templates.
---
8. Validation status
The application has been run and tested by its developer on the developer's own computer, where the workflows were reported to operate as expected.
This repository contains the exported VBA source modules. It does not include an automated test suite or the complete runtime environment required to reproduce every operation independently.
A successful local run should therefore be distinguished from a formal compatibility test across different Excel versions and computer configurations.
---
9. Repository scope
This repository is intended to preserve, review, and track the VBA source code and its development history.
It is not a repository for student records or operational school data. Do not commit personal student information, source result workbooks containing identifiable records, passwords, access tokens, or other confidential information.
---
10. Development history
Git commits document changes to the source modules and the application's architecture. Review the repository's Commits history to follow the evolution of the workflow and its implementation.
---
Project: ResultsAnalysisTerm02  
Platform: Microsoft Excel / VBA  
Scope: Second-term results analysis and school-orientation processing
