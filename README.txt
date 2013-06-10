==============
xlsxplt_pandas
==============

xlsxplt_pandas is a Python module to plot data contained in pandas
(http://pandas.pydata.org/) DataFrame objects to Excel
2010 files.  It relies heavily on the xlsxwriter module, which can be
found at https://xlsxwriter.readthedocs.org/.

The module is in a very preliminary stage and may change drastically
in the future.

Given data in a pandas.DataFrame object, the model supports the following:

* Write the data to a sheet
* Create a bar chart
* Create a line chart
* Create a scatter chart

Here is a small example::

    import pandas
    import numpy as np
    from xlsxplt_pandas import getWorkbook, plotBarChart
    
    data = pandas.DataFrame(np.random.randn(4,3))
    # Create an new Excel file and add a worksheet.
    workbook = getWorkbook('demo.xlsx')
    plotBarChart(data, workbook, 'mybarchart')

