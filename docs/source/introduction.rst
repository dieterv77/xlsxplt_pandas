.. _intro:

Introduction
============

**xlsxplt_pandas** is a Python module for plotting the data contained in a `pandas <http://pypi.python.org/pypi/pandas>`_ DataFrame in Excel 2007+ XLSX.  It relies extensively on the excellent `Xlsxwriter <http://pypi.python.org/pypi/Xlsxwriter>`_ package to produce the actual XLSX files.

Installing
**********

The xlsxplt_pandas source code and is in the
`xlsxplt_pandas repository <http://github.com/dieterv77/xlsxplt_pandas>`_ on GitHub.
You can clone the repository and install from it as follows::

    $ git clone http://github.com/dieterv77/xlsxplt_pandas

    $ cd xlsxplt_pandas
    $ python setup.py install


   
Sample usage
************

If the installation went correctly you can create a small sample program like
the following to verify that the module works correctly:

.. code-block:: python
    
    import numpy as np
    import pandas
    import xlsxplt_pandas as pdplot

    wb = pdplot.getWorkbook('demo01.xlsx')
    data = pandas.DataFrame(np.random.randn(30,3), columns=['a', 'b', 'c'])
    pdplot.plotLineChart(data, wb, 'linechart')
    wb.close()

Running this script will create a file called ``demo01.xlsx`` which should look something like
the following:

.. image:: _images/demo01.png
