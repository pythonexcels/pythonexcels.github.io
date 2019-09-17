---
layout: post
title:  Python Excel Mini Cookbook
date:   2009-10-05
categories: python
excerpt_separator: <!--end_excerpt-->
---

To get you started, I’ve illustrated a number of common tasks you can do with
Python and Excel. Each program below is a self contained example, just copy it,
paste it and run it.

<!--end_excerpt-->

A few things to note:

* These examples were tested in Excel 2007, they should work fine in earlier and later versions as well after changing the extension of the file within the wb.SaveAs() statement from .xlsx to .xls
* If you’re new to this, I recommend typing these examples by hand into IDLE, IPython or the Python interpreter, then watching the effect in Excel as you enter the commands. To make Excel visible, add the line ``excel.Visible = True`` after the ``excel =win32.gencache.EnsureDispatch('Excel.Application')`` line in the script
* These are simple examples with no error checking. Make sure the output files doesn’t exist before running the script. If the script crashes, it may leave a copy of Excel running in the background. Open the Windows Task Manager and kill the background Excel process to recover.
* These examples contain no optimization. You typically wouldn’t use a for loop to iterate through data in individual cells, it’s provided here for illustration only.

## Open Excel, Add a Workbook

The following script simply invokes Excel, adds a workbook and saves the empty workbook. [https://github.com/pythonexcels/examples/blob/master/add_a_workbook.py](https://github.com/pythonexcels/examples/blob/master/add_a_workbook.py)

```
#
# Add a workbook and save (Excel 2007)
# For older versions of excel, use the .xls file extension
#
import win32com.client as win32
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Add()
wb.SaveAs('add_a_workbook.xlsx')
excel.Application.Quit()
```

## Open an Existing Workbook

This script opens an existing workbook and displays it using
``excel.Visible =True``. The file workbook1.xlsx must already exist in your “My
Documents” directory. You can also open spreadsheet files by specifying the full
path to the file as shown below. Using ``r'`` in the statement
``r'C:\myfiles\excel\workbook2.xlsx'`` automatically escapes the backslash
characters and makes the file name a bit more concise. [https://github.com/pythonexcels/examples/blob/master/open_an_existing_workbook.py](https://github.com/pythonexcels/examples/blob/master/open_an_existing_workbook.py)

```
#
# Open an existing workbook
#
import win32com.client as win32
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open('workbook1.xlsx')
# Alternately, specify the full path to the workbook
# wb = excel.Workbooks.Open(r'C:\myfiles\excel\workbook2.xlsx')
excel.Visible = True
```

## Add a Worksheet

This script creates a new workbook with three sheets, adds a fourth worksheet
and names it MyNewSheet. [https://github.com/pythonexcels/examples/blob/master/add_a_worksheet.py](https://github.com/pythonexcels/examples/blob/master/add_a_worksheet.py)

```
#
# Add a workbook, add a worksheet,
# name it 'MyNewSheet' and save
#
import win32com.client as win32
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Add()
ws = wb.Worksheets.Add()
ws.Name = "MyNewSheet"
wb.SaveAs('add_a_worksheet.xlsx')
excel.Application.Quit()
```

## Ranges and Offsets

This script illustrates different techniques for addressing cells by using the
``Cells()`` and ``Range()`` operators. Individual cells can be addressed using
``Cells(row,column)``, where row is the row number, column is the column number,
both start from 1. Groups of cells can be addressed using ``Range()``, where the
argument in the parenthesis can be a single cell denoted by its textual name (eg
"A2"), a group noted by a textual name with a colon (eg "A3:B4") or a group
denoted with two ``Cells()`` identifiers (eg ``ws.Cells(1,1),ws.Cells(2,2)``). The
Offsetmethod provides a way to address a cell based on a reference to another
cell. [https://github.com/pythonexcels/examples/blob/master/ranges_and_offsets.py](https://github.com/pythonexcels/examples/blob/master/ranges_and_offsets.py)

```
#
# Using ranges and offsets
#
import win32com.client as win32
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Add()
ws = wb.Worksheets("Sheet1")
ws.Cells(1,1).Value = "Cell A1"
ws.Cells(1,1).Offset(2,4).Value = "Cell D2"
ws.Range("A2").Value = "Cell A2"
ws.Range("A3:B4").Value = "A3:B4"
ws.Range("A6:B7,A9:B10").Value = "A6:B7,A9:B10"
wb.SaveAs('ranges_and_offsets.xlsx')
excel.Application.Quit()
```

## Autofill Cell Contents

This script uses Excel’s autofill capability to examine data in cells A1 and A2,
then autofill the remaining column of cells through A10. [https://github.com/pythonexcels/examples/blob/master/autofill_cells.py](https://github.com/pythonexcels/examples/blob/master/autofill_cells.py)

```
#
# Autofill cell contents
#
import win32com.client as win32
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Add()
ws = wb.Worksheets("Sheet1")
ws.Range("A1").Value = 1
ws.Range("A2").Value = 2
ws.Range("A1:A2").AutoFill(ws.Range("A1:A10"),win32.constants.xlFillDefault)
wb.SaveAs('autofill_cells.xlsx')
excel.Application.Quit()
```

## Cell Color

This script illustrates adding an interior color to the cell using
``Interior.ColorIndex``. Column A, rows 1 through 20 are filled with a number and
assigned that ``ColorIndex``. [https://github.com/pythonexcels/examples/blob/master/cell_color.py](https://github.com/pythonexcels/examples/blob/master/cell_color.py)

```
#
# Add an interior color to cells
#
import win32com.client as win32
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Add()
ws = wb.Worksheets("Sheet1")
for i in range (1,21):
    ws.Cells(i,1).Value = i
    ws.Cells(i,1).Interior.ColorIndex = i
wb.SaveAs('cell_color.xlsx')
excel.Application.Quit()
```

## Column Formatting

This script creates two columns of data, one narrow and one wide, then formats
the column width with the ``ColumnWidth`` property. You can also use the
``Columns.AutoFit()`` function to autofit all columns in the spreadsheet.
[https://github.com/pythonexcels/examples/blob/master/column_widths.py](https://github.com/pythonexcels/examples/blob/master/column_widths.py)

```
#
# Set column widths
#
import win32com.client as win32
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Add()
ws = wb.Worksheets("Sheet1")
ws.Range("A1:A10").Value = "A"
ws.Range("B1:B10").Value = "This is a very long line of text"
ws.Columns(1).ColumnWidth = 1
ws.Range("B:B").ColumnWidth = 27
# Alternately, you can autofit all columns in the worksheet
# ws.Columns.AutoFit()
wb.SaveAs('column_widths.xlsx')
excel.Application.Quit()
```

## Copying Data from Worksheet to Worksheet

This script uses the ``FillAcrossSheets()`` method to copy data from one
location to all other worksheets in the workbook. Specifically, the data in the
range A1:J10 is copied from Sheet1 to sheets Sheet2 and Sheet3.
[https://github.com/pythonexcels/examples/blob/master/copy_worksheet_to_worksheet.py](https://github.com/pythonexcels/examples/blob/master/copy_worksheet_to_worksheet.py)

```
#
# Copy data and formatting from a range of one worksheet
# to all other worksheets in a workbook
#
import win32com.client as win32
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Add()
ws = wb.Worksheets("Sheet1")
ws.Range("A1:J10").Formula = "=row()*column()"
wb.Worksheets.FillAcrossSheets(wb.Worksheets("Sheet1").Range("A1:J10"))
wb.SaveAs('copy_worksheet_to_worksheet.xlsx')
excel.Application.Quit()
```

## Format Worksheet Cells

This script creates two columns of data, then formats the font type and font
size used in the worksheet. Five different fonts and sizes are used, the numbers
are formatted using a monetary format. [https://github.com/pythonexcels/examples/blob/master/format_cells.py](https://github.com/pythonexcels/examples/blob/master/format_cells.py)

```
#
# Format cell font name and size, format numbers in monetary format
#
import win32com.client as win32
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Add()
ws = wb.Worksheets("Sheet1")

for i,font in enumerate(["Arial","Courier New","Garamond","Georgia","Verdana"]):
    ws.Range(ws.Cells(i+1,1),ws.Cells(i+1,2)).Value = [font,i+i]
    ws.Range(ws.Cells(i+1,1),ws.Cells(i+1,2)).Font.Name = font
    ws.Range(ws.Cells(i+1,1),ws.Cells(i+1,2)).Font.Size = 12+i

ws.Range("A1:A5").HorizontalAlignment = win32.constants.xlRight
ws.Range("B1:B5").NumberFormat = "$###,##0.00"
ws.Columns.AutoFit()
wb.SaveAs('format_cells.xlsx')
excel.Application.Quit()
```

## Setting Row Height

This script illustrates row height. Similar to column height, row height can be
set with the ``RowHeight`` method. You can also use ``AutoFit()`` to
automatically adjust the row height based on cell contents.
[https://github.com/pythonexcels/examples/blob/master/row_height.py](https://github.com/pythonexcels/examples/blob/master/row_height.py)

```
#
# Set row heights and align text within the cell
#
import win32com.client as win32
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Add()
ws = wb.Worksheets("Sheet1")
ws.Range("A1:A2").Value = "1 line"
ws.Range("B1:B2").Value = "Two\nlines"
ws.Range("C1:C2").Value = "Three\nlines\nhere"
ws.Range("D1:D2").Value = "This\nis\nfour\nlines"
ws.Rows(1).RowHeight = 60
ws.Range("2:2").RowHeight = 120
ws.Rows(1).VerticalAlignment = win32.constants.xlCenter
ws.Range("2:2").VerticalAlignment = win32.constants.xlCenter

# Alternately, you can autofit all rows in the worksheet
# ws.Rows.AutoFit()

wb.SaveAs('row_height.xlsx')
excel.Application.Quit()
```

## Prerequisites

Python (refer to [http://www.python.org](http://www.python.org))

Win32 Python module (refer to [http://sourceforge.net/projects/pywin32](http://sourceforge.net/projects/pywin32))

Microsoft Excel (refer to [http://office.microsoft.com/excel](http://office.microsoft.com/excel))

## Source Files and Scripts

Source for the program and data text file are available at[http://github.com/pythonexcels/examples](http://github.com/pythonexcels/examples)

Originally posted on October 5, 2009
