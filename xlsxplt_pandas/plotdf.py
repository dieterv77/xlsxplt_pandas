import datetime

import pandas

from xlsxwriter.workbook import Workbook
from xlsxwriter.utility import xl_cell_to_rowcol, xl_rowcol_to_cell

def __getLocation(df, kwargs):
    if 'loc' in kwargs:
        cell = xl_rowcol_to_cell(*(kwargs['loc']))
    else:
        cell = xl_rowcol_to_cell(2, len(df.columns) + 3) 
    return cell

def getWorkbook(fname):
    """Return a xlsxwriter Workbook by the given name"""
    return Workbook(fname)

def writeData(df, wb, sheetname, **kwargs):
    """Write DataFrame to given sheetname in the given Workbook

    Parameters
    ----------
    df : pandas.DataFrame
        DataFrame with data
    wb : xlsxwriter.Workbook
    sheetname: : string
        Name of sheet to which data and plot should be written

    """
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
            'values':     [sheetname, 1, idx+1, len(df.index)+1, idx+1]
        })

    # Set an Excel chart style.
    if 'style' in kwargs:
        chart.set_style(kwargs['style'])

def plotBarChart(df, wb, sheetname, **kwargs):
    """Bar chart of columns in given DataFrame

    Parameters
    ----------
    df : pandas.DataFrame
        DataFrame with data
    wb : xlsxwriter.Workbook
    sheetname: : string
        Name of sheet to which data and plot should be written

    Other parameters
    ----------------
    subtype : string, optional
        Possible values: 'stacked', 'percent_stacked'
    title : string, optional
        Chart title
    style : int, optional
        Used to set the style of the chart to one of the 48 built-in styles available on the Design tab in Excel
    loc : (int, int) tuple, optional
        Row and column number where to locate the plot, if not specified the plot is placed to the right of the data

    """
    worksheet = writeData(df, wb, sheetname, **kwargs)
    params = {'type': 'bar'}
    if 'subtype' in kwargs:
        params['subtype'] = kwargs['subtype']
    chart = wb.add_chart(params)
    addSeries(df, chart, sheetname, **kwargs)
    # Insert the chart into the worksheet (with an offset).
    cell = __getLocation(df, kwargs)
    worksheet.insert_chart(cell, chart, {'x_scale': 2.0, 'y_scale': 2.0})

def plotColumnChart(df, wb, sheetname, **kwargs):
    """Column chart of columns in given DataFrame

    Parameters
    ----------
    df : pandas.DataFrame
        DataFrame with data
    wb : xlsxwriter.Workbook
    sheetname: : string
        Name of sheet to which data and plot should be written

    Other parameters
    ----------------
    subtype : string, optional
        Possible values: 'stacked', 'percent_stacked'
    title : string, optional
        Chart title
    style : int, optional
        Used to set the style of the chart to one of the 48 built-in styles available on the Design tab in Excel
    loc : (int, int) tuple, optional
        Row and column number where to locate the plot, if not specified the plot is placed to the right of the data

    """
    worksheet = writeData(df, wb, sheetname, **kwargs)
    params = {'type': 'column'}
    if 'subtype' in kwargs:
        params['subtype'] = kwargs['subtype']
    chart = wb.add_chart(params)
    addSeries(df, chart, sheetname, **kwargs)
    # Insert the chart into the worksheet (with an offset).
    cell = __getLocation(df, kwargs)
    worksheet.insert_chart(cell, chart, {'x_scale': 2.0, 'y_scale': 2.0})

def plotLineChart(df, wb, sheetname, **kwargs):
    """Line chart of columns in given DataFrame

    Parameters
    ----------
    df : pandas.DataFrame
        DataFrame with data
    wb : xlsxwriter.Workbook
    sheetname: : string
        Name of sheet to which data and plot should be written

    Other parameters
    ----------------
    subtype : string, optional
        Possible values: 'marker_only', 'straight_with_markers', 'straight', 'smooth_with_markers', 'smooth'
    title : string, optional
        Chart title
    style : int, optional
        Used to set the style of the chart to one of the 48 built-in styles available on the Design tab in Excel
    loc : (int, int) tuple, optional
        Row and column number where to locate the plot, if not specified the plot is placed to the right of the data

    """
    worksheet = writeData(df, wb, sheetname, **kwargs)
    params = {'type': 'line'}
    if 'subtype' in kwargs:
        params['subtype'] = kwargs['subtype']
    chart = wb.add_chart(params)
    addSeries(df, chart, sheetname, **kwargs)

    #Handle subtype here, since it is not actually an Xlsxwriter option for line charts
    if 'subtype' in kwargs:
        subtype = kwargs['subtype']
        if 'marker' in subtype:
            # Go through each series and define default values.
            for series in chart.series:
                # Set a marker type unless there is a user defined type.
                series['marker'] = {'type': 'automatic',
                                    'automatic': True,
                                    'defined': True,
                                    'line': {'defined': False},
                                    'fill': {'defined': False}
                                    }

        # Turn on smoothing if required
        if 'smooth' in subtype:
            for series in chart.series:
                series['smooth'] = True

        if subtype == 'marker_only':
            for series in chart.series:
                series['line'] = {'width': 2.25,
                                  'none': 1,
                                  'defined': True,
                                  }

    # Insert the chart into the worksheet (with an offset).
    cell = __getLocation(df, kwargs)
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
    """Scatter plot pairs of columns of given DataFrame

    Parameters
    ----------
    df : pandas.DataFrame
        DataFrame with data
    pairs : dict
        Dict mapping names to pairs (tuples) of columns names
        in df.  This describes each series that will be plotted
    wb : xlsxwriter.Workbook
    sheetname: : string
        Name of sheet to which data and plot should be written

    Other parameters
    ----------------
    subtype : string, optional
        Possible values: 'marker_only', 'straight_with_markers', 'straight', 'smooth_with_markers', 'smooth'
    title : string, optional
        Chart title
    style : int, optional
        Used to set the style of the chart to one of the 48 built-in styles available on the Design tab in Excel
    loc : (int, int) tuple, optional
        Row and column number where to locate the plot, if not specified the plot is placed to the right of the data

    """
    worksheet = writeData(df, wb, sheetname, **kwargs)
    params = {'type': 'scatter'}
    if 'subtype' in kwargs:
        params['subtype'] = kwargs['subtype']
    chart = wb.add_chart(params)
    addScatterSeries(df, pairs, chart, sheetname, **kwargs)
    
    # Insert the chart into the worksheet (with an offset).
    cell = __getLocation(df, kwargs)
    worksheet.insert_chart(cell, chart, {'x_scale': 2.0, 'y_scale': 2.0})

if __name__ == "__main__":
    wb = Workbook('test.xlsx')
    df = pandas.DataFrame.from_csv('test_dates.csv')
    plotLineChart(df, wb, 'test_dates', style=42)
