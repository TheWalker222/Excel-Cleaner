Clean Excel

A lightweight Python application for automatically cleaning and standardizing Excel files.

The program provides a simple graphical interface where users can select one or multiple .xlsx files and clean them in one click.

Features

Clean multiple Excel files at once

Remove completely empty rows

Remove duplicate rows

Normalize text capitalization

Handle common capitalization exceptions

Detect and preserve URLs

Detect and normalize email addresses

Detect dates in multiple formats

Detect integers and decimal numbers

Normalize European and US number formats

Detect monetary values

Preserve values with leading zeros as text

Automatically resize Excel columns

Make header rows bold

Freeze the first row

Add Excel filters

Background processing with threading

Simple GUI built with CustomTkinter

Supported number formats

The cleaner can recognize several common number formats, including:

1500
1 500
1,500
1.500
1,500.00
1.500,00
15,50
15.50

Examples:

1,500.00 → 1500.00
1.500,00 → 1500.00
15,50    → 15.50

Note: Values such as 1,500 or 1.500 can be ambiguous depending on the country and formatting convention. The program uses predefined rules to determine whether a separator represents thousands or decimals.

Supported currencies

Currently supported currency symbols:

$
€
₹

Examples:

$1,500.00
1.500,00 €
₹15,000

Currency values are converted into numeric Excel values so they can be used in formulas and calculations.

Date detection

Supported date formats include:

DD.MM.YYYY
DD.MM.YY
DD/MM/YYYY
DD/MM/YY
DD-MM-YYYY
DD-MM-YY

Detected dates are stored as actual Excel dates and displayed using:

DD.MM.YYYY

Text normalization

Ordinary text is cleaned and converted to title case.

Example:

john smith → John Smith
helsinki   → Helsinki

Special exceptions are handled separately.

Examples:

usa        → USA
u.s.a      → USA
uk         → UK
api        → API
iphone     → iPhone
ipad       → iPad
macbook    → MacBook
mc'donalds → McDonald's

Additional exceptions can be added to the exceptions dictionary.

Email detection

The application recognizes email addresses such as:

john.smith@gmail.com
alex123@example.fi
name+work@mail.co.uk

Detected emails are preserved separately from ordinary text cleaning.

URL detection

The cleaner detects common URLs including:

https://example.com
http://example.org
example.fi
example.com/page

URLs are kept separate from text capitalization rules to avoid accidentally modifying valid links.

Duplicate removal

Duplicate rows can optionally be removed.

A row is considered a duplicate when all of its values match a row that has already appeared.

Example:

ID | Name | City
1  | Alex | Helsinki
2  | John | Espoo
1  | Alex | Helsinki

The second occurrence of:

1 | Alex | Helsinki

will be removed.

Empty row removal

Completely empty rows can optionally be removed from the workbook.

Rows containing actual data are preserved.

GUI

The graphical interface is built using:

tkinter

customtkinter

Users can:

Add one or multiple Excel files

Choose cleaning options

Start processing

See the current processing status

Remove files from the list using:

Double click

Delete

Backspace

Excel processing runs in a separate thread so the interface does not freeze while large files are being cleaned.

Technologies

Python

OpenPyXL

CustomTkinter

Tkinter

Regular Expressions (re)

Threading

Datetime

Installation

Clone or download the project and install the required dependencies:

pip install openpyxl customtkinter

Running the application

Run the main GUI file:

python main.py

The exact filename may differ depending on the project structure.

Project structure

Example:

CleanExcel/
│
├── main.py
├── CleanExcel.py
├── README.md
└── requirements.txt

main.py

Contains the graphical user interface and file-selection logic.

CleanExcel.py

Contains the Excel cleaning and data-processing logic.

Example workflow

Input:

ID     Name             Balance        Email
0012   john smith       $1,500.00      JOHN@EXAMPLE.COM
0013   ALEX JOHNSON     1.500,50 €     alex@example.com
0013   ALEX JOHNSON     1.500,50 €     alex@example.com

After cleaning:

ID     Name             Balance        Email
0012   John Smith       1500.00        john@example.com
0013   Alex Johnson     1500.50        alex@example.com

The duplicate row can also be removed automatically when the corresponding option is enabled.

Current limitations

Some data formats are inherently ambiguous.

For example:

1,500

may mean:

1500

in one locale, or:

1.5

in another.

The program currently uses heuristic rules rather than automatically detecting the source locale.

Text capitalization is also rule-based, so unusual company names, product names, abbreviations, and surnames may require additional entries in the exceptions dictionary.

Planned improvements

Possible future improvements include:

Select which columns should be cleaned

Remove duplicates based on selected columns

Detect Excel locale automatically

Add more currencies

Add configurable capitalization rules

Add detailed cleaning statistics

Show how many duplicates were removed

Show how many empty rows were removed

Add progress bar

Export a cleaning report

Preserve or copy more workbook formatting

Add optional AI-assisted detection for ambiguous values

Package the application as a standalone Windows .exe

Goal

The goal of this project is to automate repetitive Excel cleaning tasks while keeping the process fast, predictable, and easy to use.

Instead of manually fixing formatting, duplicate rows, empty records, inconsistent dates, numbers, and text, users can process their files through one simple interface.
