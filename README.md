# PDF-Processing-and-Data-Extraction-Tool
PDF Processing and Data Extraction Tool

中文版

PDF处理与数据提取工具
========================

简介
----

本项目是一个自动化工具，旨在优化和自动化从Excel文件中提取数据、管理PDF文件的特定业务流程。通过自动识别Excel文件中的关键信息，程序能够自动管理和处理相应的PDF文件，同时生成详细的日志文件以记录操作过程。

功能
----

- **数据提取**：从指定的Excel文件和工作表中提取含有特定表头的列的所有非空值。
- **房间号处理**：处理Excel文件中的房间号信息，并生成房间占用的统计信息。
- **PDF文件管理**：根据从Excel文件中提取的代码，自动识别并复制匹配的PDF文件到指定目录。
- **日志记录**：生成详细的日志文件，记录操作过程中的每一步操作结果。

如何运行
--------

1. 确保安装了所有必要的依赖项，通过`requirements.txt`安装：
pip install -r requirements.txt

2. 运行主脚本以开始处理：
python main_script.py

3.目录
c:\zbq\
pdf和团队excel都需要放在c:\zbq\目录

4.团队excel表格
护照和房间必须在表格中的第一行，任意列都可以

注意：请根据实际路径和文件名替换`main_script.py`，确保Excel文件和PDF文件的路径正确设置。

English Version

PDF Processing and Data Extraction Tool
=======================================

Introduction
------------

This project is an automation tool designed to optimize and automate specific business processes involving data extraction from Excel files and managing PDF files. By automatically identifying key information in Excel files, the program can manage and process corresponding PDF files while generating detailed log files to record the process.

Features
--------

- **Data Extraction**: Extract all non-empty values from columns with specific headers in designated Excel files and sheets.
- **Room Number Processing**: Process room number information from Excel files and generate statistics on room occupancy.
- **PDF File Management**: Automatically identify and copy matching PDF files to specified directories based on codes extracted from Excel files.
- **Logging**: Generate detailed log files documenting each step of the operation process.

How to Run
----------

1. Ensure all necessary dependencies are installed by installing from `requirements.txt`:
pip install -r requirements.txt

2. Execute the main script to start processing:
python main_script.py

3. Directory
c:\zbq\ Both PDF and team Excel files should be placed in the c:\zbq\ directory.

4. Team Excel Spreadsheet
Passport and room must be in the first row of the spreadsheet, in any column.（chinese only）

Note: Replace `main_script.py` with the actual path and filename, ensuring the paths to Excel and PDF files are correctly set.

这个README的框架为项目提供了一个清晰的介绍，包括项目的目的、主要功能、以及如何运行项目的基本指南。您可以根据项目的具体情况调整内容。
