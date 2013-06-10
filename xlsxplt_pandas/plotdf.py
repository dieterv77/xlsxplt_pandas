import datetime

import numpy as np
import pandas

from xlsxwriter.workbook import Workbook
from xlsxwriter.utility import xl_cell_to_rowcol, xl_rowcol_to_cell

def getWorkbook(fname):
    """Return a xlsxwriter Workbook by the given name"""
    return Workbook(fname)

def writeData(df, wb, sheetname, **kwargs):
    worksheet = wb.add_worksheet(sheetname)
    date_format = wb.add_format({'num_format': 'yyyy-mm-dd'}) 
    bold = wb.add_format({'bold': 1})

    if isinstance(df.columns, pandas.DatetimeIndex):
        worksheet.write_row('B1', df.columns, date_format)
    else:
        worksheet.write_row('B1', df.columns, bold)

    for idx, (name, data) in enumerate(df.iterrows()):
        if isinstance(name, datetime.date):
            worksheet.write(idx+1,0,name, date_format)
        else:
            worksheet.write(idx+1,0,name, bold)
        worksheet.write_row('B' + str(idx+2), data)

    return worksheet

def addSeries(df, chart, sheetname, **kwargs):
    if 'title' in kwargs:
        chart.set_title({'name': kwargs['title']})
    for idx, col in enumerate(df.columns):
        namecell = xl_rowcol_to_cell(0,idx+1)
        chart.add_series({
            'name':       '=%s!%s' % (sheetname, namecell),
            'categories': [sheetname, 1, 0, len(df.index)+1, 0],
            'values':     [sheetname, 1, idx+1, len(df.index)+1, idx+1],
        })

    # Set an Excel chart style.
    if 'style' in kwargs:
        chart.set_style(kwargs['style'])

def plotBarChart(df, wb, sheetname, **kwargs):
    """Plot bar chart of all data in DataFrame df"""
    worksheet = writeData(df, wb, sheetname, **kwargs)
    chart = wb.add_chart({'type': 'column'})
    addSeries(df, chart, sheetname, **kwargs)
    # Insert the chart into the worksheet (with an offset).
    cell = xl_rowcol_to_cell(2, len(df.columns) + 3) 
    worksheet.insert_chart(cell, chart, {'x_scale': 2.0, 'y_scale': 2.0})

def plotLineChart(df, wb, sheetname, **kwargs):
    """Plot line chart of all data in DataFrame df"""
    worksheet = writeData(df, wb, sheetname, **kwargs)
    chart = wb.add_chart({'type': 'line'})
    addSeries(df, chart, sheetname, **kwargs)
    # Insert the chart into the worksheet (with an offset).
    cell = xl_rowcol_to_cell(2, len(df.columns) + 3) 
    worksheet.insert_chart(cell, chart, {'x_scale': 2.0, 'y_scale': 2.0})

def addScatterSeries(df, pairs, chart, sheetname, **kwargs):
    if 'title' in kwargs:
        chart.set_title({'name': kwargs['title']})
    name2idx = dict((c,idx) for idx, c in enumerate(df.columns))
    for name, (col1, col2) in pairs.iteritems():
        idx1 = name2idx[col1]
        idx2 = name2idx[col2]
        chart.add_series({
            'name':       name,
            'categories': [sheetname, 1, idx1+1, len(df.index)+1, idx1+1],
            'values':     [sheetname, 1, idx2+1, len(df.index)+1, idx2+1],
        })

    # Set an Excel chart style.
    if 'style' in kwargs:
        chart.set_style(kwargs['style'])

def plotScatterChart(df, pairs, wb, sheetname, **kwargs):
    """pairs must be a dict mapping a name to pairs of column names
       in DataFrame df.
       This describes each series that will be plotted
    """
    worksheet = writeData(df, wb, sheetname, **kwargs)
    if 'subtype' in kwargs:
        chart = wb.add_chart({'type': 'scatter', 'subtype': kwargs['subtype']})
    else:
        chart = wb.add_chart({'type': 'scatter'})
    addScatterSeries(df, pairs, chart, sheetname, **kwargs)
    
    # Insert the chart into the worksheet (with an offset).
    cell = xl_rowcol_to_cell(2, len(df.columns) + 3) 
    worksheet.insert_chart(cell, chart, {'x_scale': 2.0, 'y_scale': 2.0})

if __name__ == "__main__":
    wb = Workbook('test.xlsx')
    df = pandas.DataFrame.from_csv('test_dates.csv')
    plotLineChart(df, wb, 'test_dates', style=42)
