---
layout: post
title:  Installing Python
date:   2009-07-18
updated: 2019-09-20
categories: python
excerpt_separator: <!--end_excerpt-->
---

Getting started with Python is easy, you just need to download the Python
installation package and install onto your computer.

<!--end_excerpt-->

## Versions

As of this writing, Python 3.7 is the latest version, and all examples have been
tested using Python 3.7.3. The original version of these articles was developed
with Python 2.6, but I recommend that you use the latest Python 3 version if you
can.

## Which Python distribution should I use?

There are different ways to obtain Python for the Windows platform. Two of the
most popular distributions are from the Python foundation and Activestate
Software. For other alternatives, see the wiki page at
https://wiki.python.org/moin/PythonDistributions .

## Python.org

The site https://www.python.org, maintained by the Python Software Foundation,
is the main portal for information on Python. Python.org contains news,
documentation, information on the latest releases, download source for current
and previous versions of Python, and binary install files for Windows. The
latest version, as well as older versions of Python, can be found at
[https://www.python.org/downloads](https://www.python.org/downloads). This is
the Python distribution that I typically use, the examples that follow are based
on Python release 3.7.3.

## ActiveState Python

ActiveState Software provides software solutions for individuals and businesses,
including a complete download package containing Python executables and
documentation called ActivePython. The package is non-open source and is available
with an OEM license, which can be important in some corporate environments.
Please refer to
[https://www.activestate.com/activepython](https://www.activestate.com/activepython)
for more information. Python luminary Alex Martelli offers a concise description
of Why ActiveState in [this StackOverflow
post](https://stackoverflow.com/questions/1352528/why-does-activepython-exist).

## Build from Source

Source code for Python is available at
[https://www.python.org/downloads](https://www.python.org/downloads), and while
it is possible to create a working Python installation for Windows by compiling
from source, it’s beyond the scope of these exercises.

## Python Modules

One of the great things about Python is its excellent library support. Python is
“batteries included”: a large number of libraries are provided out-of-the-box,
right in the standard distribution. Other modules are available from a variety
of sources and chances are you can find a module that helps you solve a problem
with a simple web search. I’ll use a variety of third-party modules for the
exercises that follow, but you’ll at least need to start with the pywin32
module.

After installing Python, you can install the pywin32 module with the Python
`pip` install program. Open a Windows Command window and install pywin32 as
follows:

```
C:\>pip install pywin32
Collecting pywin32
Downloading https://files.pythonhosted.org/packages/...
ywin32-225-cp37-cp37m-win32.whl (8.4MB)
Installing collected packages: pywin32
Successfully installed pywin32-225
```

# Resources

I won’t be covering the basics of Python, but I suggest you check one of the many
other resources. The links listed on
[https://www.python.org/doc](https://www.python.org/doc) are an excellent
starting point. You can also find a number of other resources by searching for
“learn python”.

Originally posted on July 18, 2009 / Updated September 20, 2019

