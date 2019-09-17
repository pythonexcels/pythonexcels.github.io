---
layout: post
title:  Installing Python
date:   2009-07-18
categories: python
excerpt_separator: <!--end_excerpt-->
---

Getting started with Python is easy, you just need to download the python
installation package and install onto your computer.


<!--end_excerpt-->

## Versions

As of this writing, Python 3.1 is the latest version, though I won’t be using
the latest bleeding-edge build for my exercises. There’s nothing wrong with 3.1,
but third party package support for the packages I use can trail the latest
release by a version or two. If you plan on installing third party libraries not
provided in the standard Python release, you will want to use a slightly older
version. My exercises will be based on the latest Python 2.6.

## Which Python foundation should I use?

There are different ways to obtain Python for the Windows platform. Two of the
most popular distributions are from the Python foundation and Activestate
Software.

## Python.org

The site http://www.python.org, maintained by the Python Software Foundation, is
the main portal for information on Python. Python.org contains news,
documentation, information on the latest releases, download source for current
and previous versions of Python, and binary install files for Windows. The
latest version, as well as older versions of Python, can be found at
[http://www.python.org/download](http://www.python.org/download). This is the
Python foundation that I typically use, the examples that follow will be based
on Python 2.6.

## ActiveState Python

ActiveState Software provides software solutions for individuals and businesses,
including a complete download package containing Python executables and
documentation called ActivePython. The package is non open source and available
with an OEM license, which can be important in some corporate environments.
Please refer to
[http://www.activestate.com/activepython](http://www.activestate.com/activepython)
for more information. Python luminary Alex Martelli offers a concise description
of Why ActiveState in [this StackOverflow
post](http://stackoverflow.com/questions/1352528/why-does-activepython-exist).

## Build from Source

Source code for Python is available at
[http://www.python.org/download](http://www.python.org/download), and while it
is possible to create a working Python installation for Windows by compiling
from source, it’s beyond the scope of these exercises.

## Python Modules

One of the wonderful things about Python is the excellent library support for
the language. Python is “batteries included”: a large number of libraries are
provided out-of-the-box, right in the standard distribution. Other modules are
available from a variety of sources, chances are you can find a module that
helps you solve a problem with a simple web search. I’ll use a variety of third
party modules for the exercises that follow, but you’ll at least need to start
with the pywin module. You can find the appropriate pywin for your Python 2.6
installation at [http://sourceforge.net/projects/pywin32/files](http://sourceforge.net/projects/pywin32/files)

# Resources

I won’t be covering the basics of Python, I suggest you check one of the many
other resources. The links listed on
[http://www.python.org/doc](http://www.python.org/doc) are an excellent starting
point, many people also like the Dive into Python book at
[https://diveintopython.org](https://diveintopython.org). You can also find a
number of other resources by searching for "learn python".

Originally posted on July 18, 2009
