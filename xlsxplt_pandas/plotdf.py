import datetime
from collections import defaultdict

import numpy as np
import pandas

from xlsxwriter.workbook import Workbook
from xlsxwriter.utility import xl_cell_to_rowcol, xl_rowcol_to_cell

def __addReference(df, pairs, reffn):
    minval = min(df[x].min() for x,y in pairs.values())
    maxval = max(df[x].max() for x,y in pairs.values())
    minval = minval - 0.1 * abs(minval)
    maxval = maxval + 0.1 * abs(maxval)
    x = pandas.Series(np.linspace(minval, maxval, len(df.index)), index=df.index)
    y = x.apply(reffn)
    if 'refx' in df.columns or 'refy' in df.columns or 'Reference' in pairs:
        raise Exception('Unable to add reference columns, name conflict')
    df = df.copy()
    df['refx'] = x
    df['refy'] = y
    pairs['Reference'] = ('refx','refy')
    return df, pairs

def __sortDF(df, pairs):
    """For each pair, return a df that ensures that the x-values are in ascending order
       Note this will clobber the index.  Assumes y columns are not repeated
    """
    x2y = defaultdict(set)
    for x, y in pairs.values():
        x2y[x].add(y)
    final = []
    for x, ys in x2y.items():
        subdf = df[[x] + sorted(ys)].sort(x)
        subdf.index = list(range(len(subdf.index)))
        final.append(subdf)
    return pandas.concat(final,axis=1)

def __addAxisInfo(chart, kwargs):
    if 'x_title' in kwargs:
        chart.set_x_axis({'name': kwargs['x_title']})
    if 'x_lim' in kwargs:
        lim = kwargs['x_lim']
        chart.set_x_axis({'min': lim[0], 'max': lim[1]})
    if 'y_title' in kwargs:
        chart.set_y_axis({'name': kwargs['y_title']})
    if 'y_lim' in kwargs:
        lim = kwargs['y_lim']
        chart.set_y_axis({'min': lim[0], 'max': lim[1]})

def __getLocation(df, kwargs):
    if 'loc' in kwargs:
        cell = xl_rowcol_to_cell(*(kwargs['loc']))
    else:
        cell = xl_rowcol_to_cell(2, len(df.columns) + 3) 
    return cell

def __addQuotes(name):
    if not name.isalnum():
        return "'" + name + "'"
    return name

def getWorkbook(fname, options=None):
    """Return a xlsxwriter Workbook by the given name"""
    if options is not None:
        return Workbook(fname, options)
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
    secondaries = set()
    if 'secondary_y' in kwargs:
        secondaries = set(kwargs['secondary_y'])
    for idx, col in enumerate(df.columns):
        namecell = xl_rowcol_to_cell(0,idx+1)
        info = {
            # 'name':       '=%s!%s' % (sheetname, namecell),
            'name':       [__addQuotes(sheetname), 0, idx+1],
            'categories': [__addQuotes(sheetname), 1, 0, len(df.index), 0],
            'values':     [__addQuotes(sheetname), 1, idx+1, len(df.index), idx+1]
        }
        if col in secondaries:
            info['y2_axis'] = 1
        if 'gap' in kwargs:
            info['gap'] = kwargs['gap']
        chart.add_series(info)

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
    __addAxisInfo(chart, kwargs)
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
    __addAxisInfo(chart, kwargs)
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
    secondary_y : iterable, optional
        list of columns whose scale goes on the secondary y-axis

    """
    worksheet = writeData(df, wb, sheetname, **kwargs)
    params = {'type': 'line'}
    if 'subtype' in kwargs:
        params['subtype'] = kwargs['subtype']
    chart = wb.add_chart(params)
    __addAxisInfo(chart, kwargs)
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
    cols = sorted(x for x in pairs.keys() if x != 'Reference')
    if 'Reference' in pairs:
        cols = cols + ['Reference']
    for name in cols:
        (col1, col2) = pairs[name]
        idx1 = name2idx[col1]
        idx2 = name2idx[col2]
        params = {
            'name':       name,
            'categories': [sheetname, 1, idx1+1, len(df.index)+1, idx1+1],
            'values':     [sheetname, 1, idx2+1, len(df.index)+1, idx2+1],
        }
        if name == 'Reference':
            params['marker'] = {'type': 'none'}
            params['smooth'] = True
            params['line'] = {'dash_type': 'solid'}
        chart.add_series(params)

    # Set an Excel chart style.
    if 'style' in kwargs:
        chart.set_style(kwargs['style'])

def plotScatterChart(df, pairs, wb, sheetname, **kwargs):
    """Scatter plot pairs of columns of given DataFrame

    Parameters
    ----------
    Two ways to use this depending on first two arguments:
      1.  df: DataFrame
          pairs: Dict mapping name to tuples of size two, indicating the pair of columns to be scattered,
          if pairs is None, then it assumes there's only one pair in the DataFrame and will scatter them
      2.  df and pairs are pandas Series, to be scatter against each other
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
    sortonx : boolean, optional (default: False)
        Sort the pairs on the x values for nicer lines.  This will only include data to be plotted in the sheet.
    reference : callable, option (default: None)
        Pass a function to insert a reference series based on provided callable which should take a float argument
        and return a float

    """
    if isinstance(df, pandas.Series) and isinstance(pairs, pandas.Series):
        df = df.to_frame()
        df2 = pairs.to_frame()
        if df2.columns[0] == 0:
            df2.columns = [1] 
        df = pandas.concat([df, df2], axis=1)
        origlen = len(df.index)
        df = df.dropna()
        newlen = len(df.index)
        if newlen != origlen:
            print('Dropped', origlen - newlen, 'rows due to missing values')
        df.columns = [str(x) for x in df.columns]
        pairs = {'data': (df.columns[0], df.columns[1])} 
    elif isinstance(df, pandas.DataFrame):
        if pairs is None:
            if len(df.columns) != 2:
                raise Exception('Pairs cannot be None if DataFrame has more than 2 columns')
            pairs = {'data': (df.columns[0], df.columns[1])}
    if len(pairs) == 1:
        pair = list(pairs.values())[0]
        if 'x_title' not in kwargs:
            kwargs['x_title'] = pair[0]
        if 'y_title' not in kwargs:
            kwargs['y_title'] = pair[1]
    if 'sortonx' in kwargs and kwargs['sortonx']:
        df = __sortDF(df, pairs)
    pairs = pairs.copy()
    if 'reference' in kwargs and kwargs['reference'] is not None:
        df, pairs = __addReference(df, pairs, kwargs['reference'])

    worksheet = writeData(df, wb, sheetname, **kwargs)
    params = {'type': 'scatter'}
    if 'subtype' in kwargs:
        params['subtype'] = kwargs['subtype']
    chart = wb.add_chart(params)
    __addAxisInfo(chart, kwargs)
    addScatterSeries(df, pairs, chart, sheetname, **kwargs)
    
    # Insert the chart into the worksheet (with an offset).
    cell = __getLocation(df, kwargs)
    worksheet.insert_chart(cell, chart, {'x_scale': 2.0, 'y_scale': 2.0})

def plotHistogram(df, wb, sheetname, **kwargs):
    """Histogram chart of columns in given DataFrame

    Parameters
    ----------
    df : pandas.DataFrame
        DataFrame with data
    wb : xlsxwriter.Workbook
    sheetname: : string
        Name of sheet to which data and plot should be written

    Other parameters
    ----------------
    bins : None, optional
    title : string, optional
        Chart title
    style : int, optional
        Used to set the style of the chart to one of the 48 built-in styles available on the Design tab in Excel
    loc : (int, int) tuple, optional
        Row and column number where to locate the plot, if not specified the plot is placed to the right of the data

    """
    alldata = df.values.flatten()
    alldata = alldata[~np.isnan(alldata)]
    if 'bins' in kwargs:
        h, b = np.histogram(alldata, bins=kwargs['bin'])
    else:
        h, b = np.histogram(alldata)
    bins = b
    bindf = {}
    for colname, data in df.items():
        data = data.dropna().values
        h, b = np.histogram(data, bins=bins)
        bindf[colname] = pandas.Series(h, index=[x for x in b[:-1]])
    df = pandas.DataFrame(bindf)
    worksheet = writeData(df, wb, sheetname, **kwargs)
    params = {'type': 'column'}
    if 'subtype' in kwargs:
        params['subtype'] = kwargs['subtype']
    chart = wb.add_chart(params)
    __addAxisInfo(chart, kwargs)
    kwargs['gap'] = 0
    addSeries(df, chart, sheetname, **kwargs)
    # Insert the chart into the worksheet (with an offset).
    cell = __getLocation(df, kwargs)
    worksheet.insert_chart(cell, chart, {'x_scale': 2.0, 'y_scale': 2.0})

if __name__ == "__main__":
    wb = Workbook('test.xlsx')
    df = pandas.DataFrame.from_csv('test_dates.csv')
    plotLineChart(df, wb, 'test_dates', style=42)
