---
layout: post
title:  Mapping Excel VB Macros to Python
date:   2009-10-12
updated: 2019-09-27
categories: python
excerpt_separator: <!--end_excerpt-->
---

A handy feature in Excel is the ability to quickly record a sequence
of operations into a Visual Basic (VB) macro. It’s also fairly simple
to take a captured VB macro, change it slightly, and use it in your
Python scripts. I’ve used this capability dozens of times over the
years to automate spreadsheet calculations and pivot table generation.
I now have a good understanding of how to port the VB macro into
Python; let me share the technique with you.

<!--end_excerpt-->

In this post, I’ll capture a simple set of operations as a Visual Basic
macro, examine the macro, and port it to Python. I’m using the
MultiplicationTable.xlsx file as a starting point, it’s a simple 10×10
multiplication table that will be expanded and reformatted. You can create this
table yourself or download the file from
[https://github.com/pythonexcels/examples/raw/master/MultiplicationTable.xlsx](https://github.com/pythonexcels/examples/raw/master/MultiplicationTable.xlsx)

# Enabling the Developer Tab

The first step is to capture the macro in Excel by using the Record
Macro feature in Excel. The Record Macro button is located in the
Developer tab, which might be disabled in your application. To enable
the Developer tab,

1. Click the File tab

2. Click Options, then click “Customize Ribbon”

3. In the “Customize the Ribbon” section, select “Main Tabs”, enable
   the Developer option in the list, and click OK

![exceloptions](/assets/images/20190927_developer_tab.png)

In older versions of Excel,

1. Select “Excel Options” from the ribbon menu

2. Select “Popular” in the left column

3. Enable the “Show Developer tab in the Ribbon” option and click OK

![exceloptions](/assets/images/20091012_exceloptions.png)

# Recording a Macro

Now you are ready to record a macro. Starting with a simple
spreadsheet containing a table of data, click on the “Developer” tab,
then “Record Macro”.

![recordmacro](/assets/images/20091012_recordmacro.png)

The goal is to expand the existing table to a 15×15 table, adjust the
column width to make the table appear more proportional, and save the
new spreadsheet. Now that the macro is recording, the first step is to
select the last row of data and expand it by dragging it down an
additional five rows. First, select the data:

![selectrow](/assets/images/20091012_selectrow.png)

then drag the mouse to create five new rows of data.

![dragrow](/assets/images/20091012_dragrow.png)

Using the same technique, select the last column of data and create five
new columns.

![selectcolumn](/assets/images/20091012_selectcolumn.png)

![dragcolumn](/assets/images/20091012_dragcolumn.png)

Now you have a 15×15 multiplication table. To resize the columns, select the
headers for columns B through P, click the right mouse and select “Column
Width”.

![columnwidth](/assets/images/20091012_columnwidth.png)

Enter “4” as the new column width and click OK. The spreadsheet now looks like this:

![resizecolumns](/assets/images/20091012_resizecolumns.png)

Now stop capturing the macro by clicking the “Stop Recording” button.

![stoprecording](/assets/images/20091012_stoprecording.png)

To view the macro, click the Macros (View Macros) button.

![viewmacros](/assets/images/20091012_viewmacros.png)

Select the macro you just recorded and click Edit. The macro should be
named Macro1. If you were experimenting, you will have more than one
macro in the list and should select the highest numbered macro.

![editmacro](/assets/images/20091012_editmacro.png)

The tool opens your macro in the Microsoft Visual Basic Integrated
Development Environment (IDE). Your macro should look similar to the
following macro:

```
Sub Macro1()
'
' Macro1 Macro
'
'
    Range("B11:K11").Select
    Selection.AutoFill Destination:=Range("B11:K16"), Type:=xlFillDefault
    Range("B11:K16").Select
    Range("K2:K16").Select
    Selection.AutoFill Destination:=Range("K2:P16"), Type:=xlFillDefault
    Range("K2:P16").Select
    Columns("B:P").Select
    Selection.ColumnWidth = 4
End Sub
```

Don’t worry if there are some extra or redundant lines in your macro,
they can be removed as the script is ported.

Next, save your new spreadsheet as “MultiplicationTable15.xlsx”, but
don't close Excel. We’re now ready to start Python and port the VB
macro.

## Porting from Visual Basic to Python

To get started, open the Python Integrated Development Environment
(IDLE), and open the spreadsheet with the original 10×10
multiplication table by entering the following four commands. Make
sure the “MultiplicationTable.xlsx” spreadsheet is in your My
Documents folder.

```
import win32com.client as win32
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open('MultiplicationTable.xlsx')
excel.Visible = True
```

Your screen should now look like this:

![introfullscreen](/assets/images/20190927_intro_fullscreen.png)

In this example, the first command, ``import win32com.client as
win32``, imports the win32 module. The next statement, ``excel =
win32.gencache.EnsureDispatch('Excel.Application')``, attaches to a
running Excel process or opens a new Excel process if needed. The
command ``wb = excel.Workbooks.Open('MultiplicationTable.xlsx')``
opens the worksheet. In general, you’ll need to run
``excel.Workbooks.Open()`` or ``excel.Workbooks.Add()`` to open an
existing Excel file or create a new workbook. The command
``excel.Visible = True`` makes Excel visible on the screen.

Looking at the Macro1 macro, the first command is
``Range("B11:K11").Select``. The Range method is within the
context of the Worksheet, so you need to create an object that points
to the worksheet. The command ``ws = wb.Worksheets('Sheet1')`` will do
the trick.

![wsworksheet](/assets/images/20190927_wsworksheet.png)

> If you noticed, I made a typo when entering the command and typed
> ``Worksheet`` instead of ``Worksheets``. Don’t panic if you make a
> mistake as I did; in most cases you can simply retype the correct
> command and continue on.

After you’ve created the ``ws`` object to reference the worksheet,
append the `Range` command to ``ws.`` and try it. Note that in Python,
`Select` is a function and requires the open and close parenthesis
pair in order to operate correctly. This pattern may be used for every
``Range().Select`` line in the original VB macro.

Type `ws.Range("B11:K11").Select()` at the prompt in the IDLE window,
then bring the worksheet to the foreground. Confirm that range B11:K11
has been selected as shown in the following figure.

![wsb11k11select](/assets/images/20190927_b11k11select.png)

 The next task is to autofill the five rows below the existing table
by using the ``Selection.AutoFillDestination:=Range("B11:K16"),
Type:=xlFillDefault`` VB command. `Selection` is a method at the Excel
application level, so you need to prefix the command with ``excel.``
The arguments ``Destination:=Range("B11:K16"), Type:=xlFillDefault``
must be provided to the function with the keyword arguments
`Destination` and `Type` or by using positional notation as I’ve done
in this example.

### Specifying Visual Basic Constant Values in a Python Script

There are two ways to provide a constant value such as
`xlFillDefault`: 1) specify the constant name or 2) specify the value
of the constant. To use the constant name in a Python program, prefix
the name with ``win32.constants``. In this example, `xlFillDefault` in
Visual Basic becomes ``win32.constants.xlFillDefault`` in Python.

Alternatively, you can use the Visual Basic IDE to display the value
of the constant and use the value in your Python program. Click the constant
in the IDE and choose Quick Info from the context menu. The IDE
displays a tooltip with the constant value as shown below:

![vbobjectbrowser](/assets/images/20190927_constant_value.png)

I’ve seen many examples where the developer replaces the constant with
the actual value (0 in this case). My preference is to avoid replacing
Excel constants with numbers in my scripts; I believe that including
the constant names increases the clarity of the script.

## Finishing the Script

Combining these translations, the full Python command is
``excel.Selection.AutoFill(Destination=ws.Range("B11:K16"),
Type=win32.constants.xlFillDefault)``, or ``excel.Selection.AutoFill(
ws.Range("B11:K16"), win32.constants.xlFillDefault)`` as I’ve used in the example.

![idlefillrow](/assets/images/20190927_idlefillrow.png)

Occasionally you’ll make a mistake when capturing a macro and record
extra, unnecessary commands. At the same time, the macro recorder
might insert additional commands that aren’t needed in the Python
script. In this case, the command ``Range("B11:K16").Select`` isn’t
needed and can be ignored. The next two macro commands,
``Range("K2:K16").Select`` and
``Selection.AutoFillDestination:=Range("K2:P16"),
Type:=xlFillDefault``, are translated in the same way as the `Select`
and `AutoFill` commands discussed earlier.

![autofill](/assets/images/20190927_worksheetfilled.png)

The multiplication table is now expanded to the full 15×15 table as
follows:

![worksheetfilled](/assets/images/20091012_worksheetfilled.png)

The next section of the macro selects columns B through P and sets
their width to 4 characters. The statement ``Columns("B:P").Select`` is a
property of the worksheet, so prefix it with the ``ws.`` identifier
and add the parenthesis to make it a Python function call. In the next
statement, ``Selection`` is a property of excel, so insert the
`excel.` prefix. The translated statements are shown below.

![idlecolumnwidth](/assets/images/20190927_idlecolumnwidth.png)

The Excel spreadsheet is now complete; the multiplication table has
been expanded to 15×15 and the column widths have been set to four
characters. At this point, translation of the macro is complete, but
the modified spreadsheet must be saved. To write the file and quit
Excel, use the `SaveAs` and `Quit` methods as shown below.

![idlesavequit](/assets/images/20190927_idlesavequit.png)

For your reference, here is the complete Python script, also available at
[https://github.com/pythonexcels/examples/raw/master/make15x15.py](https://github.com/pythonexcels/examples/raw/master/make15x15.py)

```
#
# make15x15.py
# Expand an existing 10x10 multiplication table and resize columns
#
import win32com.client as win32
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open('MultiplicationTable.xlsx')
excel.Visible = True
ws = wb.Worksheets('Sheet1')
ws.Range("B11:K11").Select()
excel.Selection.AutoFill(ws.Range("B11:K16"),win32.constants.xlFillDefault)
ws.Range("K2:K16").Select()
excel.Selection.AutoFill(ws.Range("K2:P16"),win32.constants.xlFillDefault)
ws.Columns("B:P").Select()
excel.Selection.ColumnWidth = 4
wb.SaveAs('NewMultiplicationTable.xlsx')
excel.Application.Quit()
```

If this is the first time you’ve ported an Excel macro from VB to
Python, congratulations! Please note that in this example, things are
kept simple and there is absolutely no error checking or exception
handling used. Normally you would need to provide at least a minimal
level of error checking and exception handling in your script so that
common errors (missing input file, can’t invoke Excel, etc) are caught
and handled nicely. This script was validated in Excel 2017, but
should run without issue in older versions.

## Some Porting Guidelines

* Prefix the `Range().Select` statements with the object name pointing to the worksheet (`ws` in this example)
* Append `()` when porting a method from VB to Python
* Prefix the `Selection` statements with the object name for the Excel spreadsheet (`excel` in this example)
* Prefix the `Columns` statements with the object name for the worksheet (`ws` in this example)

## Porting Reference Table for this example

Note that I didn’t capture the `Workbooks.Open()` or
`Workbooks.SaveAs` lines in the VB script, it’s left as an exercise
for the reader to research those commands.

| Visual Basic                                                            | Python                                                                        |
|-------------------------------------------------------------------------|-------------------------------------------------------------------------------|
|                                                                         | `import win32com.client as win32`                                             |
|                                                                         | `excel = win32.gencache.EnsureDispatch(‘Excel.Application’)`                  |
|                                                                         | `wb = excel.Workbooks.Open(‘MultiplicationTable.xlsx’)`                       |
|                                                                         | `wb = excel.Workbooks.Open(‘MultiplicationTable.xlsx’)`                       |
|                                                                         | `excel.Visible = True`                                                        |
|                                                                         | `ws = wb.Worksheets(‘Sheet1’)`                                                |
| `Range("B11:K11").Select`                                               | `ws.Range("B11:K11").Select()`                                                |
| `Range("B11:K11").Select`                                               | `ws.Range("B11:K11").Select()`                                                |
| `Selection.AutoFill Destination:=Range("B11:K16"), Type:=xlFillDefault` | `excel.Selection.AutoFill(ws.Range("B11:K16"),win32.constants.xlFillDefault)` |
| `Range("K2:K16").Select`                                                | `ws.Range("K2:K16").Select()`                                                 |
| `Selection.AutoFill Destination:=Range("K2:P16"), Type:=xlFillDefault`  | `excel.Selection.AutoFill(ws.Range("K2:P16"),win32.constants.xlFillDefault)`  |
| `Columns("B:P").Select`                                                 | `ws.Columns("B:P").Select()`                                                  |
| `Selection.ColumnWidth = 4`                                             | `excel.Selection.ColumnWidth = 4`                                             |
|                                                                         | `excel.Application.Quit()`                                                    |

## Prerequisites

Python (refer to [http://www.python.org](http://www.python.org))

pywin32 Python module [https://pypi.org/project/pywin32](https://pypi.org/project/pywin32)

Microsoft Excel (refer to [http://office.microsoft.com/excel](http://office.microsoft.com/excel))

## Source Files and Scripts

Source for the program make15x15.py and data text file are available
at [http://github.com/pythonexcels/examples](http://github.com/pythonexcels/examples)

Originally posted on October 12, 2009 / Updated September 27, 2019
