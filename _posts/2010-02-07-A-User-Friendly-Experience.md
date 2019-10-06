---
layout: post
title:  A User-Friendly Experience
date:   2010-02-07
updated: 2019-10-03
categories: python
excerpt_separator: <!--end_excerpt-->
---

If you’re going to offer your Windows users a new application, you
should also provide a graphical user interface to help them use it.
Graceful error recovery and useful error messages also go a long way
toward making your application easy-to-use. This post will take the
Pivot Table generation script developed in [Extending Pivot Table
Data]({% post_url 2009-12-03-Extending-Pivot-Table-Data %}) and turn
it into a user-friendly application with an improved user experience.

<!--end_excerpt-->

The scripts developed previously could be run at the command line or
by double-clicking on the icon for the script like this.

![Command Line](/assets/images/20191003_command_line.png)

![Script Icon](/assets/images/20191003_desktop.png)

This works because the input file name, ABCDCatering.xls, is hard-coded
within the script. In a corporate environment, your users have
folders that contain dozens of randomly named spreadsheets. If your
user accidentally provides a corrupt spreadsheet, the program should
recover and continue processing the other files. The script developed
in the last post needs some enhancements to make it more user-friendly,
including:

* Provide support for multiple randomly named input spreadsheets
* Add some simple message boxes and drag-and-drop support
* Improve the error checking and error recovery to give the user feedback when something goes wrong

To keep things concise, this version of the script only allows the
user to run the program by dragging and dropping files onto the
program icon. Enhancing the script to support command-line
operation is left as an exercise for the reader. Let’s work through each
of the usability issues below:

## Providing Multiple File Support

As I mentioned, Windows users typically don’t interact with the
command prompt. Instead, programs are run by clicking on their icons,
either from the desktop, a folder, or the Start menu. A user specifies
spreadsheets or document files by opening them in the application or
dragging them onto the program icon on the desktop or in the Explorer
window.

To process multiple files, the program needs to process command-line
arguments provided by the `sys.argv` list in a Python program. Note
that the first argument, `sys.argv[0]`, is used for the script name.
In the script for this example, the runexcel function is modified to
accept `sys.argv` as an argument.

```
def runexcel(args):
    ...
    for fname in args[1:]:

if __name__ == "__main__":
    runexcel(sys.argv)
```

The for loop wraps the ``wb = excel.Workbooks.Open(fname)`` call, the
``wb.SaveAs()`` call, and everything in between so each workbook is
processed within the loop. After the loop finishes, the script checks
for errors and issues a warning message if needed.

## Enabling a Primitive GUI

Adding message boxes and providing basic drag-and-drop support adds a
level of familiarity for Windows users. Python supports many GUI
frameworks, see
[http://wiki.python.org/moin/GuiProgramming](http://wiki.python.org/moin/GuiProgramming)
for a comprehensive list. Building a complete graphic interface for
this script is beyond the scope of this article, and isn’t
necessary given the intent of this script. Instead, you can add
support for simple message boxes using the MessageBoxA function built
into Windows. The basic pattern for calling a message box using this
technique is to import ctypes and call ``ctypes.windll.user32.MessageBoxA``:

```
import ctypes
ctypes.windll.user32.MessageBoxW(None,"My message","My title",0)
```

This simple code produces a message box with the text “My Message”, an
OK button, and “My title” as the top banner. When Python runs the
``ctypes.windll.user32.MessageBoxW()`` statement, program execution
pauses until the user clicks the OK button.

![Messsage Box](/assets/images/20191003_message_box.png)

## Improving Error Checking

Several problems can happen when reading user spreadsheet data:

* The user can forget to specify an input file.
* The user provides the wrong spreadsheet or even a  non-spreadsheet file type.
* The spreadsheet might be corrupted.

You need to bulletproof your script and guard against potential issues, both
known and unknown.

Previous versions of the script made limited use of the try/except pattern to
catch errors as follows:

```
try:
    wb = excel.Workbooks.Open('ABCDCatering.xls')
except:
    print "Failed to open spreadsheet ABCDCatering.xls"
    sys.exit(1)
```

erppivotdragdrop.py provides additional checking and wraps more of the
program code in the try block. If an error occurs, the error can be handled more
cleanly with a warning message. The downside of using try/except is that you
lose the traceback message telling you where the error occurred. To get this
information back, use the traceback module and the
``traceback.print_exc()`` function. One usage is to call ``traceback.print_exc()`` in the
except block like this:

```
import traceback
try:
  a = 1/0
except:
  # Do error recovery
  traceback.print_exc()
```

Exceptions are now caught and handled while providing a detailed
traceback.

## Running the script

Let’s test out the script. First, copy the script to the desktop and
drag the ABCDCatering.xls spreadsheet onto the icon. Python starts
running in the command window and processes the spreadsheet. If
the script runs successfully, you’ll see a series of messages and the
“Finished” message box.

![Finished](/assets/images/20191003_noerror.png)

If a problem occurred, a message is displayed in the command window and
a message box is displayed.

![Error Message](/assets/images/20191003_witherror.png)

The completed script is too long to reproduce here, please view the
complete script at [https://github.com/pythonexcels/examples/blob/master/erppivotdragdrop.py](https://github.com/pythonexcels/examples/blob/master/erppivotdragdrop.py)

## Prerequisites

Python (refer to [http://www.python.org](http://www.python.org))

Win32 Python module (refer to [http://sourceforge.net/projects/pywin32](http://sourceforge.net/projects/pywin32))

Microsoft Excel (refer to [http://office.microsoft.com/excel](http://office.microsoft.com/excel))

## Source Files and Scripts

Source for the program erppivotextended.py and spreadsheet file ABCDCatering.xls
are available at [http://github.com/pythonexcels/examples](http://github.com/pythonexcels/examples)

Originally posted on February 7, 2010 / Updated October 3, 2019
