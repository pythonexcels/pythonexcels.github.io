---
layout: post
title:  Automating Pivot Tables with Python
date:   2009-11-23
categories: python
excerpt_separator: <!--end_excerpt-->
---

In the [last post]({% post_url 2009-11-11-Introducing-Pivot-Tables
%}), I started with a raw data set and create four different pivot
tables to answer various questions about the data. This post
describes how to do the same thing in an automated fashion by using
Python and Excel.

<!--end_excerpt-->

Pivot tables in Excel are a powerful tool you can use for deriving
basic business intelligence from your data. As discussed in the last
post, there are occasions when you need to do interactive data
mining by changing column and row fields. But in my experience, it’s
nice to have my favorite reports built automatically, with the reports
available as soon as I open the spreadsheet. In this post, I’ll
develop the code to create a set of pivot tables automatically in the
worksheet.

The goal of this exercise is to automatically generate the pivot
tables described in the last post and save them to a new Excel file.

![Pivot Tables](/assets/images/20191002_reports.png)

I started with the file newABCDCatering.xls from the previous post and
recorded a macro to create the simple pivot table below. The table
shows Net Bookings by Sales Rep and Food Name for the last four
quarters as follows:

![Net Bookings](/assets/images/20191002_setup.png)

The recorded macro looks like this:

```
Sub Macro1()
'
' Macro1 Macro
'
    Selection.CurrentRegion.Select
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Sheet2!R1C1:R791C13", Version:=xlPivotTableVersion10).CreatePivotTable _
        TableDestination:="Sheet3!R3C1", TableName:="PivotTable1", DefaultVersion _
        :=xlPivotTableVersion10
    Sheets("Sheet3").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Fiscal Year")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Fiscal Quarter")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Sales Rep Name")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Food Name")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Net Booking"), "Sum of Net Booking", xlSum
End Sub
```

Looking at the VB macro, a general pattern should be apparent. First,
the pivot table is created with the
``ActiveWorkbook.PivotCaches.Create()`` method. Next, the columns and
rows are configured with a series of
``ActiveSheet.PivotTables("PivotTable1").PivotFields()`` methods.
Finally, the field used in the Values section of the table is
configured using the
``ActiveSheet.PivotTables("PivotTable1").AddDataField`` method. 



A pivot table has four basic areas where you can place a field from
the list:

| Report Filters | `.Orientation = xlPageField `   |
| Columns area   | `.Orientation = xlColumnField ` |
| Rows area      | `.Orientation = xlRowField `    |
| Values area    | `PivotTables().AddDataField()`  |


You can add multiple fields to each of these areas. In this example,
“Sales Rep Name” and “Food Name” were added to the Rows area. The
ordering of the fields changes the appearance of the table.

In [Mapping Excel VB Macros to Python]({% post_url
2009-10-12-Mapping-Excel-VB-Macros-to-Python %}), I covered a
technique for recording a Visual Basic (VB) macro and porting it to
Python. I could capture the VB macro and port it to Python
line-by-line with that approach, however, the Python script would
inherit a lot of redundancy. A better technique is to
extract some of the repetitive tasks into a function which can be
called with different parameters to build different pivot tables. The
following general-purpose function, `addpivot`, takes the table title,
filters, columns, rows, and data value to be summed and generates a
pivot table.

```
def addpivot(wb,sourcedata,title,filters=(),columns=(),
             rows=(),sumvalue=(),sortfield=""):
    """Build a pivot table using the provided source location data
    and specified fields
    """
    ...
    for fieldlist,fieldc in ((filters ,win32c.xlPageField),
                            (columns  ,win32c.xlColumnField),
                            (rows     ,win32c.xlRowField)):
        for i,val in enumerate(fieldlist):
            wb.ActiveSheet.PivotTables(tname).PivotFields(val).Orientation = fieldc
        wb.ActiveSheet.PivotTables(tname).PivotFields(val).Position = i+1
    ...
```

The actual values for filters, columns and rows in the `addpivot`
function are defined in the call to the function. For example, to
answer the question: “What were the total sales in each of the last
four quarters?”, the pivot table is built with the following call to
the `addpivot` function:

```
# What were the total sales in each of the last four quarters?
addpivot(wb,src,
         title="Sales by Quarter",
         filters=(),
         columns=(),
         rows=("Fiscal Quarter",),
         sumvalue="Sum of Net Booking",
         sortfield=())
```

which defines a pivot table using the row header “Fiscal Quarter” and
data value “Sum of Net Booking”. Note that the title parameter,
`title="Sales by Quarter"`, is used as the worksheet title and the tab
name.

![Title Tabs](/assets/images/20191002_titletabsbq.png)

The script does take some shortcuts. To keep things simple, this
script is limited to adding “Sum of” values only, and doesn’t handle
other Summarize Value functions such as Count, Min, Max, etc.

The complete script is shown below. Note that adding pivot tables
increases the size of the output Excel file, which can be mitigated by
disabling caching of pivot table data. Line 48 of the script contains
the command `newsheet.PivotTables(tname).SaveData = False`, which has
been commented out. Uncommenting this command will reduce the size of
the output Excel file, but will require that you refresh the pivot
table by clicking “Refresh Data” on the PivotTable toolbar.

```
#
# erpdatapivot.py:
# Load raw EPR data, clean up header info and
# build 5 pivot tables
#
import win32com.client as win32
win32c = win32.constants
import sys
import itertools
tablecount = itertools.count(1)

def addpivot(wb,sourcedata,title,filters=(),columns=(),
             rows=(),sumvalue=(),sortfield=""):
    """Build a pivot table using the provided source location data
    and specified fields
    """
    newsheet = wb.Sheets.Add()
    newsheet.Cells(1,1).Value = title
    newsheet.Cells(1,1).Font.Size = 16

    # Build the Pivot Table
    tname = "PivotTable%d"%next(tablecount)

    pc = wb.PivotCaches().Add(SourceType=win32c.xlDatabase,
                                 SourceData=sourcedata)
    pt = pc.CreatePivotTable(TableDestination="%s!R4C1"%newsheet.Name,
                             TableName=tname,
                             DefaultVersion=win32c.xlPivotTableVersion10)
    wb.Sheets(newsheet.Name).Select()
    wb.Sheets(newsheet.Name).Cells(3,1).Select()
    for fieldlist,fieldc in ((filters ,win32c.xlPageField),
                            (columns  ,win32c.xlColumnField),
                            (rows     ,win32c.xlRowField)):
        for i,val in enumerate(fieldlist):
            wb.ActiveSheet.PivotTables(tname).PivotFields(val).Orientation = fieldc
            wb.ActiveSheet.PivotTables(tname).PivotFields(val).Position = i+1

    wb.ActiveSheet.PivotTables(tname).AddDataField(
        wb.ActiveSheet.PivotTables(tname).PivotFields(sumvalue[7:]),
        sumvalue,
        win32c.xlSum)
    if len(sortfield) != 0:
        wb.ActiveSheet.PivotTables(tname).PivotFields(sortfield[0]).AutoSort(sortfield[1], sumvalue)
    newsheet.Name = title

    # Uncomment the next command to limit output file size, but make sure
    # to click Refresh Data on the PivotTable toolbar to update the table
    # newsheet.PivotTables(tname).SaveData = False

    return tname

def runexcel():
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    #excel.Visible = True
    try:
        wb = excel.Workbooks.Open('ABCDCatering.xls')
    except:
        print ("Failed to open spreadsheet ABCDCatering.xls")
        sys.exit(1)
    ws = wb.Sheets('Sheet1')
    xldata = ws.UsedRange.Value
    newdata = []
    for row in xldata:
        if len(row) == 13 and row[-1] is not None:
            newdata.append(list(row))
    lasthdr = "Col A"
    for i,field in enumerate(newdata[0]):
        if field is None:
            newdata[0][i] = lasthdr + " Name"
        else:
            lasthdr = newdata[0][i]
    rowcnt = len(newdata)
    colcnt = len(newdata[0])
    wsnew = wb.Sheets.Add()
    wsnew.Range(wsnew.Cells(1,1),wsnew.Cells(rowcnt,colcnt)).Value = newdata
    wsnew.Columns.AutoFit()

    src = "%s!R1C1:R%dC%d"%(wsnew.Name,rowcnt,colcnt)

    # What were the total sales in each of the last four quarters?
    addpivot(wb,src,
             title="Sales by Quarter",
             filters=(),
             columns=(),
             rows=("Fiscal Quarter",),
             sumvalue="Sum of Net Booking",
             sortfield=())

    # What are the sales for each food item in each quarter?
    addpivot(wb,src,
             title="Sales by Food Item",
             filters=(),
             columns=("Food Name",),
             rows=("Fiscal Quarter",),
             sumvalue="Sum of Net Booking",
             sortfield=())

    # Who were the top 10 customers for ABCD Catering in 2009?
    addpivot(wb,src,
             title="Top 10 Customers",
             filters=(),
             columns=(),
             rows=("Company Name",),
             sumvalue="Sum of Net Booking",
             sortfield=("Company Name",win32c.xlDescending))

    # Who was the highest producing sales rep for the year?
    addpivot(wb,src,
             title="Top Sales Reps",
             filters=(),
             columns=(),
             rows=("Sales Rep Name","Company Name"),
             sumvalue="Sum of Net Booking",
             sortfield=("Sales Rep Name",win32c.xlDescending))

    # What food item had the highest unit sales in Q4?
    ptname = addpivot(wb,src,
             title="Unit Sales by Food",
             filters=("Fiscal Quarter",),
             columns=(),
             rows=("Food Name",),
             sumvalue="Sum of Quantity",
             sortfield=("Food Name",win32c.xlDescending))
    wb.Sheets("Unit Sales by Food").PivotTables(ptname).PivotFields("Fiscal Quarter").CurrentPage = "2009-Q4"

    if int(float(excel.Version)) >= 12:
        wb.SaveAs('newABCDCatering.xlsx',win32c.xlOpenXMLWorkbook)
    else:
        wb.SaveAs('newABCDCatering.xls')
    excel.Application.Quit()

if __name__ == "__main__":
    runexcel()
```

## Prerequisites

Python (refer to [http://www.python.org](http://www.python.org))

pywin32 Python module (Refer to [https://pypi.org/project/pywin32](https://pypi.org/project/pywin32))

Microsoft Excel (refer to [http://office.microsoft.com/excel](http://office.microsoft.com/excel))

## Source Files and Scripts

Source for the program erpdatapivot.py and input spreadsheet file
ABCDCatering.xls are available at [http://github.com/pythonexcels/examples](http://github.com/pythonexcels/examples)

Originally posted on November 23, 2009 / Updated October 2, 2019
