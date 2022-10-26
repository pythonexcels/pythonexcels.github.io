---
layout: post
title:  Python Excel Mini Cookbook
date:   2009-10-05
updated: 2019-09-21
categories: python
excerpt_separator: <!--end_excerpt-->
---

To get you started, I’ve created sample scripts to demonstrate some
common tasks you can do with Python and Excel. Each program below is a
self-contained example, just copy it, paste it, and run it.

<!--end_excerpt-->

Alternately, grab the collection of example scripts from the GitHub
repository. Once you have the scripts, examine, run, and modify each
script to understand it. You can copy the scripts as a zip file from
[https://github.com/pythonexcels/examples/archive/master.zip](https://github.com/pythonexcels/examples/archive/master.zip)
or clone the repository with the following command:

```shell
git clone https://github.com/pythonexcels/examples.git
```

A few things to note:

* These examples were tested in Excel versions 2016 and 2007, they should work
  fine in other versions as well.
* For really old versions of Excel, change .xlsx to .xls after in the
  workbook.SaveAs() statement
* If you’re new to this, I recommend typing these examples by hand
  into IDLE, IPython or the Python interpreter, then watching the
  effect in Excel as you enter the commands. To make Excel visible,
  add the line ``excel.Visible = True`` after the ``excel
  =win32.gencache.EnsureDispatch('Excel.Application')`` line in the
  script
* These are simple examples with no error checking. Make sure the
  output files doesn’t exist before running the script. If the script
  crashes, it may leave a copy of Excel running in the background.
  Open the Windows Task Manager and kill the background Excel process
  to recover.
* These examples contain no optimization. You typically wouldn’t use a
  for loop to iterate through data in individual cells, it’s provided
  here for illustration only.

## Table of Contents
- [Table of Contents](#table-of-contents)
- [Prerequisites](#prerequisites)
- [Recipes](#recipes)
	- [Open Excel, Add a Workbook](#open-excel-add-a-workbook)
	- [Open an Existing Workbook](#open-an-existing-workbook)
	- [Add a Worksheet](#add-a-worksheet)
	- [Ranges and Offsets](#ranges-and-offsets)
	- [Autofill Cell Contents](#autofill-cell-contents)
	- [Cell Color](#cell-color)
	- [Column Formatting](#column-formatting)
	- [Copying Data from Worksheet to Worksheet](#copying-data-from-worksheet-to-worksheet)
	- [Format Worksheet Cells](#format-worksheet-cells)
	- [Setting Row Height](#setting-row-height)
- [Source Files and Scripts](#source-files-and-scripts)


## Prerequisites
The following MUST be installed in order for the snippets to work:
- [**Python**](http://www.python.org)

- [**`pywin32`**](https://pypi.org/project/pywin32) (*Python module*)

- [**Microsoft Office Excel**](http://office.microsoft.com/excel) (*any other alternative won't work*)

## Recipes
### Open Excel, Add a Workbook

The following script simply invokes Excel, adds a workbook and saves the empty workbook.

[Go to example](https://github.com/pythonexcels/examples/blob/master/add_a_workbook.py)

```python
#
# Add a workbook and save to My Documents / Documents Library
# For really old versions of Excel, use the .xls file extension
#
import win32com.client as win32

excel = win32.gencache.EnsureDispatch('Excel.Application')
workbook = excel.Workbooks.Add()
workbook.SaveAs('add_a_workbook.xlsx')

excel.Application.Quit()
```

### Open an Existing Workbook

This script opens an existing workbook and displays it using
``excel.Visible =True``. The file workbook1.xlsx must already exist in
your local directory. You can also open spreadsheet files by
specifying the full path to the file as shown below. Using ``r'`` in
the statement ``r'C:\myfiles\excel\workbook2.xlsx'`` automatically
escapes the backslash characters and makes the file name a bit more
concise.

[Go to example](https://github.com/pythonexcels/examples/blob/master/open_an_existing_workbook.py)

```python
#
# Open an existing workbook
#
import win32com.client as win32

excel = win32.gencache.EnsureDispatch('Excel.Application')
workbook = excel.Workbooks.Open('workbook1.xlsx')
# Alternately, specify the full path to the workbook
# workbook = excel.Workbooks.Open(r'C:\myfiles\excel\workbook2.xlsx')
excel.Visible = True
```

### Add a Worksheet

This script creates a new workbook with three sheets, adds a fourth worksheet, names it *NewFirstSheet* and places it as first, then adds a fifth worksheet, names it *LastSheet* and places it as last.  
Then it saves the file to save to My Documents / Documents Library.

> **NOTE**  
> You can adjust the final position of the worksheet by making use of the `Before` and `After` parameters of the `workbook.Worksheets.Add` function.

[Go to example](https://github.com/pythonexcels/examples/blob/master/add_a_worksheet.py)

```python
#
# Add a workbook, add two worksheets,
# rename them and save
#
import win32com.client as win32

excel = win32.gencache.EnsureDispatch('Excel.Application')
workbook = excel.Workbooks.Add()

# Add a new sheet (before the first)
worksheet = workbook.Worksheets.Add()
worksheet.Name = "NewFirstSheet"

# Add a new sheet (after the last)
worksheet = workbook.Worksheets.Add(Before=None, After=workbook.Worksheets(workbook.Worksheets.Count))
worksheet.Name = "LastSheet"

workbook.SaveAs('add_a_worksheet.xlsx')
excel.Application.Quit()
```

### Ranges and Offsets

This script illustrates different techniques for addressing cells by
using the ``Cells()`` and ``Range()`` operators. Individual cells can
be addressed using ``Cells(row,column)``, where <em>row</em> is the
row number and <em>column</em> is the column number. Row and column
numbering begins at 1.

Single cells or groups of cells can be addressed using ``Range()``,
where the argument in the parenthesis can be a single cell name in
double quotes (for example, "A2"), a group with two cell names
separated by a colon and surrounded by double quotes (for example,
"A3:B4") or a group denoted with two ``Cells()`` identifiers (for
example, ``worksheet.Cells(1,1),worksheet.Cells(2,2)``). The ``Offset()`` method
provides a method to address a cell based on a reference to another cell.

[Go to example](https://github.com/pythonexcels/examples/blob/master/ranges_and_offsets.py)

```python
#
# Using ranges and offsets
#
import win32com.client as win32

excel = win32.gencache.EnsureDispatch('Excel.Application')
workbook = excel.Workbooks.Add()
worksheet = workbook.Worksheets("Sheet1")
worksheet.Cells(1,1).Value = "Cell A1"
worksheet.Cells(1,1).Offset(2,4).Value = "Cell D2"
worksheet.Range("A2").Value = "Cell A2"
worksheet.Range("A3:B4").Value = "A3:B4"
worksheet.Range("A6:B7,A9:B10").Value = "A6:B7,A9:B10"
workbook.SaveAs('ranges_and_offsets.xlsx')
excel.Application.Quit()
```

### Autofill Cell Contents

This script uses Excel’s autofill capability to examine data in cells
A1 and A2, then autofill the remaining column of cells through A10.
The script sets cell A1 to 1, sets cell A2 to 2, and autofills the
range A1:A10. As a result, cells A1:A10 are populated with 1, 2, 3, 4,
and so on up to 10.

[Go to example](https://github.com/pythonexcels/examples/blob/master/autofill_cells.py)

```python
#
# Autofill cell contents
#
import win32com.client as win32

excel = win32.gencache.EnsureDispatch('Excel.Application')
workbook = excel.Workbooks.Add()
worksheet = workbook.Worksheets("Sheet1")
worksheet.Range("A1").Value = 1
worksheet.Range("A2").Value = 2
worksheet.Range("A1:A2").AutoFill(worksheet.Range("A1:A10"),win32.constants.xlFillDefault)
workbook.SaveAs('autofill_cells.xlsx')
excel.Application.Quit()
```

### Cell Color

This script adds an interior (background) color to the cell with the
``Interior.ColorIndex`` method. Column A, rows 1 through 20 are filled
with a number and assigned that ``ColorIndex``.

[Go to example](https://github.com/pythonexcels/examples/blob/master/cell_color.py)

```python
#
# Add an interior color to cells
#
import win32com.client as win32

excel = win32.gencache.EnsureDispatch('Excel.Application')
workbook = excel.Workbooks.Add()
worksheet = workbook.Worksheets("Sheet1")
for i in range (1,21):
    worksheet.Cells(i,1).Value = i
    worksheet.Cells(i,1).Interior.ColorIndex = i
workbook.SaveAs('cell_color.xlsx')
excel.Application.Quit()
```

### Column Formatting

This script creates two columns of data, one narrow and one wide, then formats
the column width by setting the ``ColumnWidth`` property. You can also use the
``Columns.AutoFit()`` function to autofit all columns in the spreadsheet.

[Go to example](https://github.com/pythonexcels/examples/blob/master/column_widths.py)

```python
#
# Set column widths
#
import win32com.client as win32

excel = win32.gencache.EnsureDispatch('Excel.Application')
workbook = excel.Workbooks.Add()
worksheet = workbook.Worksheets("Sheet1")
worksheet.Range("A1:A10").Value = "A"
worksheet.Range("B1:B10").Value = "This is a very long line of text"
worksheet.Columns(1).ColumnWidth = 1
worksheet.Range("B:B").ColumnWidth = 27
# Alternately, you can autofit all columns in the worksheet
# worksheet.Columns.AutoFit()
workbook.SaveAs('column_widths.xlsx')
excel.Application.Quit()
```

### Copying Data from Worksheet to Worksheet

This script uses the ``FillAcrossSheets()`` method to copy data from one
location to all other worksheets in the workbook. Specifically, the data in the
range A1:J10 is copied from Sheet1 to sheets Sheet2 and Sheet3.

[Go to example](https://github.com/pythonexcels/examples/blob/master/copy_worksheet_to_worksheet.py)

```python
#
# Copy data and formatting from a range of one worksheet
# to all other worksheets in a workbook
#
import win32com.client as win32

excel = win32.gencache.EnsureDispatch('Excel.Application')
workbook = excel.Workbooks.Add()
worksheet = workbook.Worksheets("Sheet1")
worksheet.Range("A1:J10").Formula = "=row()*column()"
workbook.Worksheets.FillAcrossSheets(workbook.Worksheets("Sheet1").Range("A1:J10"))
workbook.SaveAs('copy_worksheet_to_worksheet.xlsx')
excel.Application.Quit()
```

### Format Worksheet Cells

This script creates two columns of data, then formats the font type and font
size used in the worksheet. Five different fonts and sizes are used, the numbers
are formatted using a monetary format.

[Go to example](https://github.com/pythonexcels/examples/blob/master/format_cells.py)

```python
#
# Format cell font name and size, format numbers in monetary format
#
import win32com.client as win32

excel = win32.gencache.EnsureDispatch('Excel.Application')
workbook = excel.Workbooks.Add()
worksheet = workbook.Worksheets("Sheet1")

for i,font in enumerate(["Arial","Courier New","Garamond","Georgia","Verdana"]):
    worksheet.Range(worksheet.Cells(i+1,1),worksheet.Cells(i+1,2)).Value = [font,i+i]
    worksheet.Range(worksheet.Cells(i+1,1),worksheet.Cells(i+1,2)).Font.Name = font
    worksheet.Range(worksheet.Cells(i+1,1),worksheet.Cells(i+1,2)).Font.Size = 12+i

worksheet.Range("A1:A5").HorizontalAlignment = win32.constants.xlRight
worksheet.Range("B1:B5").NumberFormat = "$###,##0.00"
worksheet.Columns.AutoFit()
workbook.SaveAs('format_cells.xlsx')
excel.Application.Quit()
```

### Setting Row Height

This script creates some sample data, then adjusts the row heights and
alignment of the data. Row height can be set with the ``RowHeight``
method. You can also use ``AutoFit()`` to automatically adjust the row
height based on cell contents.

[Go to example](https://github.com/pythonexcels/examples/blob/master/row_height.py)

```python
#
# Set row heights and align text within the cell
#
import win32com.client as win32

excel = win32.gencache.EnsureDispatch('Excel.Application')
workbook = excel.Workbooks.Add()
worksheet = workbook.Worksheets("Sheet1")
worksheet.Range("A1:A2").Value = "1 line"
worksheet.Range("B1:B2").Value = "Two\nlines"
worksheet.Range("C1:C2").Value = "Three\nlines\nhere"
worksheet.Range("D1:D2").Value = "This\nis\nfour\nlines"
worksheet.Rows(1).RowHeight = 60
worksheet.Range("2:2").RowHeight = 120
worksheet.Rows(1).VerticalAlignment = win32.constants.xlCenter
worksheet.Range("2:2").VerticalAlignment = win32.constants.xlCenter

# Alternately, you can autofit all rows in the worksheet
# worksheet.Rows.AutoFit()

workbook.SaveAs('row_height.xlsx')
excel.Application.Quit()
```

## Source Files and Scripts

Source for the program and data text file are available at [http://github.com/pythonexcels/examples](http://github.com/pythonexcels/examples)

Originally posted on October 5, 2009 / Updated September 21, 2019
