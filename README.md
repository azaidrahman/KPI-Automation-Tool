# CSM KPI Sheet Generator

## Overview
The CSM KPI Sheet Generator is a Python-based tool designed to automate the process of creating and managing Customer Success Manager (CSM) Key Performance Indicator (KPI) reports. This application streamlines the workflow for CSM teams by consolidating data from various sources and generating comprehensive Excel-based KPI sheets.

## Features

### File Upload: 
Supports uploading of job lists, budgets, payroll, and workbills files.

### Data Processing: 
Automatically processes and consolidates data from multiple sources.

### KPI Sheet Generation: 
Creates detailed Excel sheets with KPI data for each CSM.

### Export Options: 
Allows exporting of individual CSM sheets or all CSM data.

### User-Friendly GUI: 
Provides an intuitive graphical interface for easy operation.

## Main Components

- GUI (GUI.py): The main interface for user interaction.
- Execution Functions (execution_functions.py): Core logic for processing files and generating reports.
- Loading Functions (loading_functions.py): Handles reading and initial processing of input files.
- Processing Functions (processing_functions.py): Contains functions for detailed data processing and calculations.
- Export Functions (export_individual_sheets.py): Manages the export of individual CSM sheets.

## Usage

- Run main.py to start the application.
- Use the GUI to upload required files (job list, budget, payroll, and workbills).
- Click "Generate" to create the KPI sheet.
- Use export options to save individual CSM sheets or all data.

## Requirements

Python 3.x
Required libraries: pandas, openpyxl, tkinter

## Note
This tool is **strictly** designed for internal use and requires customization based on specific organizational data structures and KPI requirements.
