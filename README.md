# Beam Solver – VBA

A simply supported beam solver written in VBA for Microsoft Excel. As of 2026-07-08, the cumulative loc is around 840, give or take.

## Features

- Point load analysis
- Support reaction calculation
- Shear force diagram
- Bending moment diagram
- Bending stress calculation
- Beam slope calculation
- Beam deflection calculation
- Global extrema summary
- Automatic Excel chart generation

## Roadmap

- Applied moments
- Distributed loads
- Influence lines
- Additional support types
- VBA automation tools
- Excel import/export automation
- AutoCAD/SolidWorks integration

## Repository Layout

src/
├── Analysis.bas
├── Buttons.bas
├── Constructors.bas
├── GitTools.bas
├── Globals.bas
├── MainModule.bas
├── Results.bas
├── Solvers.bas
├── BeamANHEILyzerV1.frm
└── SummaryForm.frm

## Version Control

The `.xlsm` workbook is intentionally excluded from version control because Office macro-enabled documents are binary files and are prone to poor security. Source code is exported directly from Excel using the included `GitTools.bas` module, allowing the project to be version-controlled as plain text.
