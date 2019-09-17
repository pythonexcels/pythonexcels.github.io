---
layout: post
title:  Another XLWT Example
date:   2009-09-19
categories: python
---


After completing the [last
example](./2009_09_10_Using_XLWT_to_Write_Spreadsheets_Without_Excel.html), I
wanted to try something a little more interesting with XLWT this time around.
This article describes a short script that uses Python and XLWT to download some
raw data from the web, parse it, and write a spreadsheet with a new column
derived from the data.

The data for this example comes from research done by David Harrison and Daniel
L. Rubinfeld in “Hedonic Housing Prices and the Demand for Clean Air”, published
in the Journal of Environmental Economics and Management, Volume 5, (1978), and
contains information on location, pricing, tax and other information from the
Boston housing market. I’ll be illustrating three things in this script:

* Downloading data directly from the web
* Parsing the data, removing extraneous information at the top of the file and writing the real data fields to a spreadsheet
* Adding a hyperlink for each record that links to a Google Map, based on the latitude and longitude data given

Note that it’s possible to do these steps using a web browser and an interactive
Excel session. You can easily download the file, import it into Excel, remove
the leading text, and make a formula to produce a hyperlink. But the beauty of
this script is that everything is done automatically, which can be very handy if
the source data is constantly updated.

Here’s the script that performs these operations:

```
#
# xlwt_bostonhousing.py
#
import sys
from urllib2 import urlopen
from xlwt import Workbook, easyxf, Formula

def doxl():
    '''Read the boston_corrected.txt file based on
       Harrison, David, and Daniel L. Rubinfeld, "Hedonic Housing Prices
       and the Demand for Clean Air," Journal of Environmental Economics
       and Management, Volume 5, (1978), write to an excel spreadsheet .
       '''
    #URL = 'http://stat.cmu.edu/datasets/boston_corrected.txt'
    URL = 'https://raw.github.com/pythonexcels/xlwt/master/boston_corrected.txt'
    try:
        fp = urlopen(URL)
    except:
        print 'Failed to download %s' % URL
        sys.exit(1)
    lines = fp.readlines()

    wb = Workbook()
    ws = wb.add_sheet('Housing Data')
    ulstyle = easyxf('font: underline single')
    r = 0
    for line in lines:
        tokens = line.strip().split('\t')
        if len(tokens) != 21:
            continue
        for c,t in enumerate(tokens):
            for dtype in (int,float):
                try:
                    t = dtype(t)
                except:
                    pass
                else:
                    break
            ws.write(r,c+1,t)
        if r == 0:
            hdr = tokens
            ws.write(r,0,'MAPLINK')
        else:
            d = dict(zip(hdr,tokens))
            link = 'HYPERLINK("http://maps.google.com/maps?q=%s'+\
                   ',+%s+(Observation+%s)&hl=en&ie=UTF8&z=14&'+\
                   'iwloc=A";"MAP")'
            link = link % (d['LAT'],d['LON'],d['OBS.'])
            ws.write(r,0,Formula(link),ulstyle)

        r += 1
    wb.save('bostonhousing.xls')
    print 'Wrote bostonhousing.xls'

if __name__ == "__main__":
    doxl()
```

As in the previous post, you must have xlwt installed (refer to the site
[http://www.python-excel.org](http://www.python-excel.org) for information on
downloading and installing). Looking at the important bits in the script above,
the following lines

```
#URL = 'http://stat.cmu.edu/datasets/boston_corrected.txt'
URL = 'https://raw.github.com/pythonexcels/xlwt/master/boston_corrected.txt'
try:
    fp = urlopen(URL)
except:
    print 'Failed to download %s' % URL
    sys.exit(1)
lines = fp.readlines()
```

open the URL for the boston_corrected.txt file, then reads the URL and returns a
list of strings (update 9/15/12: the original link is now broken, I’ve updated
the script to pull the data from a copy on github). The next section:

```
wb = Workbook()
ws = wb.add_sheet('Housing Data')
ulstyle = easyxf('font: underline single')
r = 0
```

creates a new Workbook object, then adds a sheet named “Housing Data” to the
workbook. The easyfx function provides a convenient way to add formatting to the
spreadsheet; in this example, the single underline format is used to denote a
hyperlink. In the next line, the variable r acts as a row counter.

The for loop below examines each row of data:

```
for line in lines:
    tokens = line.strip().split('\\t')
    if len(tokens) != 21:
        continue
    for c,t in enumerate(tokens):
        for dtype in (int,float):
            try:
                t = dtype(t)
            except:
                pass
            else:
                break
        ws.write(r,c+1,t)
```

Each line is “stripped” (leading and trailing white space characters are
removed), then split by tab characters. A data line contains 21 fields of
information, otherwise it is rejected. To properly format the data for the
spreadsheet, the datatype is set using try-except-else within the for* loop. The
loop only considers string, integer and float data, which is sufficient for this
input data. More complex input files may contain date information which would
require additional handling. The cell data with the correct type setting is
written to the spreadsheet using the ws.write statement.

The next section builds the hyperlink to a Google Map using the latitude and
longitude information within the input data.

```
    if r == 0:
        hdr = tokens
        ws.write(r,0,'MAPLINK')
    else:
        d = dict(zip(hdr,tokens))
        link = 'HYPERLINK("http://maps.google.com/maps?q=%s'+\
               ',+%s+(Observation+%s)&hl=en&ie=UTF8&z=14&'+\
               'iwloc=A";"MAP")'
        link = link % (d['LAT'],d['LON'],d['OBS.'])
        ws.write(r,0,Formula(link),ulstyle)

    r += 1
wb.save('bostonhousing.xls')
print 'Wrote bostonhousing.xls'
```

If this is the first row of data ``if r == 0``, it is assumed to be header data
and is saved in the hdr variable. Otherwise, the statement ``d = dict(zip(hdr,tokens))``
builds a dictionary, using the header information as keys.
This allows each field to be referenced by its column heading. The hyperlink is
built by specifying a URL containing ``http://maps.google.com/maps?``, with the
corresponding latitude ``d['LAT']`` and longitude ``d['LON']`` information from the
current line of data. (Forgive the link formatting, my original one line of code
is split across four lines to reduce the line width.) Finally, the hyperlink
data is written to the spreadsheet with ``ws.write``. The last two lines write the
spreadsheet and print a message.

The original raw data looks like this:

![rawdata](/assets/images/20090918_1.png)

The bostonhousing.xls spreadsheet containing the original data and the new map
hyperlink looks like this:

![bostonhousing](/assets/images/20090918_2.png)

Clicking on the MAP link brings up a Google Map showing the location according
to the latitude and longitude.

![googlemap](/assets/images/20090918_3.png)

For some reason, the location of latitude 42.255000 and longitude -70.955000
isn’t in the town of Nahant, but in the middle of Rock Island Cove near Quincy.
It’s left as an exercise for the reader to determine why this is so.

## Prerequisites

Python (refer to [http://www.python.org](http://www.python.org))

xlwt (refer to [http://www.python-excel.org](http://www.python-excel.org))

## Source Files and Scripts

Source for the program and data text file are available
at [http://github.com/pythonexcels/xlwt/tree/master](http://github.com/pythonexcels/xlwt/tree/master)

## References

[http://www.python-excel.org](http://www.python-excel.org)

This site contains pointers to the best information available about working with
Excel files in the Python programming language.

[http://groups.google.com/group/python-excel](http://groups.google.com/group/python-excel)

Google group for questions on xlrd, xlwt, xlutils and general questions on
interfacing to Excel with Python

Originally posted on September 19, 2009
