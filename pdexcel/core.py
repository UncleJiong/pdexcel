# -*- coding: utf-8 -*-
"""
Write DataFrame to an excel file.
Created on Thu Jul  5 10:43:35 2018

@author: Jimmy Wong
@Mail: twong@126.com
"""
import pandas as pd
from pandas import Series, DataFrame, MultiIndex
import xlsxwriter.utility as xu
import re
from matplotlib.pyplot import get_cmap
from copy import deepcopy
from warnings import warn
from numbers import Number
from math import ceil
from xlsxwriter.format import Format

range_parts2 = re.compile(r'(\$?)([A-Z]{0,3})(\$?)(\d*)')

def _add_suffix_(name=None, prefix=None, keys={}):
    'Rename if name is already existed'
    if not (name is None):
        if name not in keys:
            return name
        else:
            prefix = name
    base = prefix
    for num in range(1, 10001):
        prefix = base+str(num)
        if prefix not in keys:
            return prefix
        elif num == 10000:
            raise ValueError('The number of name is out of range.')
    
def num2row(rownum):
    row = chr(ord('A') + rownum)
    return row
    
_fillnone_ = lambda x,y: [y[i] if x[i] is None else x[i] for i in range(len(x))]

def range_to_xl(row_start=None, row_end=None, col_start=None, col_end=None,
                sheetname=None):
    '''Convert zero indexed row and col cell references to a excel range string.
    
    Parameters
    --------
    Int.
    
    Examples
    --------
    >>> range_to_xl(0, 1, 2, 3)
    'C1:D2'
    >>> range_to_xl(0, None, 2, None)
    'C1'
    '''
    row_start = '' if row_start is None else str(row_start+1)
    row_end = '' if row_end is None else str(row_end+1)
    col_start = '' if col_start is None else num2row(col_start)
    col_end =  ''if col_end is None else num2row(col_end)
    if sheetname is None:
        sheetname = ''
    else:
        sheetname = "'%s'!"%sheetname
    start = col_start + row_start
    end = col_end + row_end
    if not start and not end:
        raise ValueError('All agruements is null')
    elif start and end:
        return sheetname + start + ':' + end
    elif start:
        return sheetname + start
    else:
        return sheetname + end
    
def range_to_py(range_str):
    '''Convert excel style range references to a zero indexed row and col cell.
    
    Parameters
    --------
    String.
    
    Examples
    --------
    >>> range_to_py('AA2:B3')
    [[1, 2, 26, 1]]
    >>> range_to_py("'Sheet1'!A2:BA3,'Sheet2'!A:B")
    [[1, 2, 0, 52], [None, None, 0, 1]]
    '''
    locs = []
    for rg in range_str.split(','):
        if '!' in range_str:
            rg = rg[rg.rfind('!')+1:]
        rg = rg.split(':')
        rg_py = [None, None, None, None]
        for i, rg_xls in enumerate(rg):
            rg_xls = re.findall(range_parts2, rg_xls)[0]
            if rg_xls[1]:
                if not rg_xls[3]:
                    rg_py[2+i] = xu.xl_cell_to_rowcol(rg_xls[1]+'0')[1]
                else:
                    rg_py[0+i], rg_py[2+i] = xu.xl_cell_to_rowcol(rg_xls[1]+rg_xls[3])
            else:
                rg_py[0+i] = int(rg_xls[3]) - 1
        locs.append(rg_py)
    return locs     
  
def convert_cell_args(method):
    """
    Decorator function to convert A1 notation in cell method calls
    to the default row/col notation.
    """
    def cell_wrapper(self, *args, **kwargs):
        try:
            # First arg is an int object, default to row/col notation.
            if len(args):
                int(args[0])
        except ValueError:
            # First arg isn't an int object, convert to A1 notation.
            new_args = list(range_to_py(args[0]))[0]
            new_args.extend(args[1:])
            args = new_args
        return method(self, *args, **kwargs)
    return cell_wrapper

class ColorMap:
    def __init__(self, cmap=None):
        if cmap is None:
            self.colors = ['#4F81BD', '#C0504D', '#9BBB59', '#8064A2', '#4BACC6', 
                     '#F79646', '#2C4D75', '#772C2A', '#5F7530', '#4D3B62']
        elif isinstance(cmap, str):
            cm = get_cmap(cmap)
            self.colors = ['#'+''.join([hex(int(x*250))[2:].rjust(2,'0') for x in l])
                      for l in cm.colors]
        else:
            self.colors = cmap
        self.loc = 0
        self.nb_colors = len(self.colors)

    def __iter__(self):
        return self

    def __next__(self):
        color = self.colors[self.loc]
        self.loc = (self.loc+1) % self.nb_colors
        return color
    
class Chart(object):
    '''
    Add a chart to worksheet.
    
    Parameters
    -----------
    table: Table object
        Data for chart.
    loc_chart: Tuple
        Zero indexed of chart: (row, None, column, None).
    index_col: boolean, or str, default True
        Use default index if True,
        Don't use index if False,
        Set a column as chart index if pass a column name.
    chart_col: str or list
        The columns for the value of chart.
    chart_type: str. Default `line`
        'area','bar','column','line','pie','doughnut','scatter',
        'stock','radar'
    chart_style: None OR int
        Style_id.
        Like the options in the function xlsxwriter.chart.chart.set_style,
        See https://xlsxwriter.readthedocs.io/chart.html#chart.set_style()
    legend_pars: None, False OR dict
        See chart.set_legend,
        http://xlsxwriter.readthedocs.io/working_with_charts.html
    series_pars: Dict
        The paraments for adding series,
        like the options in the function xlsxwriter.chart.chart.add_series,
        See https://xlsxwriter.readthedocs.io/chart.html#chart.add_series()
    colormap : str or colormap list, default None
        Colormap to select colors from. If string, load colormap with that
        name from matplotlib.
    '''
    def __init__(self, table, chart_name, legend_pars=None, *args, **kwargs):
        self.table = table
        self.sheet = self.table.sheet
        self.workchart = self.insert_chart(*args, **kwargs)
        self.chart_name = chart_name
        # legend
        if legend_pars is None:
            legend_pars = {}
        self.set_legend(**legend_pars)
        self.set_grid()
    
    def set_colors(self, series_pars, cm=None, set_all=False):
        '''
        Set color using colormap if the key "fill" or "line" has been set on the 
        argument `series_pars`.
        '''
        cond1 = 'fill' in series_pars and 'color' not in series_pars['fill']
        cond2 = 'line' in series_pars and 'color' not in series_pars['line']
        cond3 = set_all and 'fill' not in series_pars and 'line' not in series_pars
        if cond1 or cond2 or cond3:
            color = next(cm)
            if cond1:
                series_pars['fill']['color'] = color
            if cond2:
                series_pars['line']['color'] = color
            if cond3:
                series_pars['fill'] = {'color':color}
                series_pars['line'] = {'color':color}
        return series_pars
        
    def insert_chart(self, loc_chart=None, index_col=True, chart_col=None, 
                 chart_type=None, title=None, chart_style=None, y_axis_name=None,
                 x_scale=None, y_scale=None, x_offset=None, y_offset=None,
                 legend_pars=None, series_pars=None, colormap=None, **kwargs):
        '''Add a chart to worksheet.
        
        Parameters
        -----------
        loc_chart: Tuple.
            Zero indexed indices of chart: (row, None, column, None).
        index_col: boolean, or str, default True.
            Use default index if True,
            Don't use index if False,
            Set a column as chart index if pass a column name.
        chart_col: str or list.
            The columns for the value of chart.
        title: str or dict.
            Pass a title name or a dictionary paraments of title.
        chart_type: str. Default `line`.
            'area','bar','column','line','pie','doughnut','scatter',
            'stock','radar'
        chart_style: None OR int, Style_id
            Set the chart style type.
            See https://xlsxwriter.readthedocs.io/chart.html#chart.set_style()
        legend_pars: None, False OR dict.
            See chart.set_legend,
            http://xlsxwriter.readthedocs.io/working_with_charts.html
        series_pars: Dict.
            The paraments for adding series,
            like the options in the function xlsxwriter.chart.chart.add_series,
            See https://xlsxwriter.readthedocs.io/chart.html#chart.add_series()
        colormap : str or colormap list, default None
            Colormap to select colors from. If string, load colormap with that
            name from matplotlib.
        '''

        if colormap is None:
            cm = ColorMap()
            set_all = False
        else:
            cm = ColorMap(colormap)
            set_all = True     
        
        if series_pars is None:
            series_pars = {}
        else:
            series_pars = deepcopy(series_pars)
        if chart_type is None:
            chart_type = 'line'
        if x_scale is None:
            x_scale = 1.5
        if y_scale is None:
            y_scale = 1.5
        option_add = dict(x_scale=x_scale, y_scale=y_scale, x_offset=x_offset, y_offset=y_offset)
        option_add = {k:v for k,v in option_add.items() if not (v is None)}
        
        kwargs['type'] = kwargs.setdefault('type', chart_type)
        chart1 = self.table.workbook.add_chart(kwargs)
        # Set chart's index.
        if index_col is False:
            series_pars_baisc = {}
        elif index_col is True:
            loc_idx = self.table.get_loc_index(if_sheetname=True)
            if loc_idx:
                series_pars_baisc = {'categories': '='+loc_idx}
            else:
                series_pars_baisc = {}
        elif isinstance(index_col, str):
            loc_idx = self.table.get_loc_value(index_col, if_sheetname=True)
            series_pars_baisc = {'categories': '='+loc_idx}
        series_pars_baisc.update(series_pars)
        
        # Add selected series into chart.
        if chart_col is None:
            chart_col = self.table.df.columns
        elif isinstance(chart_col, str):
            chart_col = [chart_col]
        for colname in chart_col:
            series_pars_ = deepcopy(series_pars_baisc)
            pars_name = self.table.get_loc_value(colname,if_sheetname=1, if_value=0)
            pars_values = self.table.get_loc_value(colname,if_sheetname=1, if_columnname=0)
            series_pars_.update({'name': '='+pars_name,
                                 'values': '='+pars_values})
            series_pars_ = self.set_colors(series_pars_, cm, set_all)
            chart1.add_series(series_pars_)

        # Add a chart title and axis labels.
        if not title is None:
            if isinstance(title, str):
                title = {'name': title}
            title.setdefault('name_font', {'size':14})
            title['name_font'].setdefault('size', 14)
            chart1.set_title(title)
        # Set axis names.
        if not y_axis_name is None:
            if isinstance(y_axis_name, str):
                y_axis_name = {'name': y_axis_name}
            y_axis_name.setdefault('name_font', {'size':14})
            y_axis_name['name_font'].y_axis_name('size', 14)
            chart1.set_y_axis(y_axis_name)
            x_axis_name = y_axis_name.copy()
            x_axis_name = {'name': self.table.df.index.name}
        else:
            x_axis_name = {'name': self.table.df.index.name}
        chart1.set_x_axis(x_axis_name)

        # Set an Excel chart style.
        if chart_style is not None:
            chart1.set_style(chart_style)
        
        chart1.set_x_axis({'label_position': 'low'})
        # Insert chart into the sheet.
#        if chart_row == None:
#            chart_row = self.chart_row
#        self.chart_row = chart_row + 15
#        par_pos = '%s%s'%(self.num2row(self.cols+2), chart_row)
        if loc_chart is None:
            loc = 'A1'
        else:
            loc = range_to_xl(*loc_chart)
            
        self.table.sheet.worksheet.insert_chart(loc, chart1, option_add)
        self.chart_info = self.sheet.worksheet.charts
        info_keys = ['row', 'col', 'chart', 'x_offset', 'y_offset', 'x_scale', 'y_scale']
        c = dict(zip(info_keys, self.sheet.worksheet.charts[-1]))
        self.chart_info = c
        self.height = (c['y_scale'] + c['y_offset'] / 288) * 14.4
        self.width = (c['x_scale'] + c['x_offset'] / 480) * 7.5
        
        row_start = c['row']
        row_end = row_start + ceil(self.height)
        col_start = c['col']
        col_end = col_start + ceil(self.width)
        self.loc = (row_start, row_end, col_start, col_end)
        return c['chart']
    
    def set_table(self, drop_legend=True, **kwargs):
        kwargs.setdefault('show_keys', True)
        self.workchart.set_table(kwargs)
        if drop_legend:
            self.workchart.set_legend({'position': 'none'})
    
    def set_legend(self, none=None, position='bottom', font=None, 
                   delete_series=None, layout=None, **kwargs):
        """
        Set the chart legend options.
        
        Parameters
        -----------
        none: boolean or None.
            In Excel chart legends are on by default. Turns off the chart legend if passed `True`.
        position: str, default bottom
            Set the position of the chart legend. The available positions are
                - bottom
                - left
                - right
                - overlay_left
                - overlay_right
                - none
        font: dict
            See the Chart Fonts section for more details on font properties.
                - border: Set the border properties of the legend such as \
                color and style. See Chart formatting: Border.
                - fill: Set the solid fill properties of the legend such as \
                color. See Chart formatting: Solid Fill.
                - pattern: Set the pattern fill properties of the legend. See \
                Chart formatting: Pattern Fill.
                - gradient: Set the gradient fill properties of the legend. \
                See Chart formatting: Gradient Fill.
                - delete_series: This allows you to remove one or more series \
                from the legend (the series will still display on the chart). \
                This property takes a list as an argument and the series are \
                zero indexed:
        delete_series: list
            This allows you to remove one or more series from the legend \
            (the series will still display on the chart). This property takes \
            a list as an argument and the series are zero indexed.
        layout: dict
            Set the (x, y) position of the legend in chart relative units:
                
            >>> chart.set_legend({'layout': {
            ...                              'x':      0.80,
            ...                              'y':      0.37,
            ...                              'width':  0.12,
            ...                              'height': 0.25,
            ...                             }
            ...                   })
    
        total_func: dict
            eg: {'cola':{'total_string': 'Total'},'colb':{'total_function': 'sum}}
        tbl_pars: dict
            Options to pass to `xlsxwriter.worksheet.add_table`.
        merge_format: str
            Fomart for table_name.
        kwargs: keywords
            Options to pass to `DataFrame.to_excel`.
        """
        if none is not None:
            kwargs['none'] = none
        if position is None:
            kwargs['position'] = 'none'
        else:
            kwargs['position'] = position
        if font is not None:
            kwargs['font'] = font
        if delete_series is not None:
            kwargs['delete_series'] = delete_series
        if layout is not None:
            kwargs['layout'] = layout
        return self.workchart.set_legend(kwargs)
            
          
    def set_grid(self, visible=True, color='#D9D9D9', width=None, 
                 transparency=None, axis='y', **kwargs):
        '''
        Set the major gridlines for the axis.
        
        Parameters
        -----------
        transparency: int
            The transparency property sets the transparency of the line color \
            in the integer range 1 - 100.The color must be set for transparency\
            to work, it doesn’t work with an automatic/default color.
        sheetname: str
            Name of sheet which will contain DataFrame.
        df: DataFrame, optional
            Data to excel.
        index: boolean, default True, optional
            Write index.
        tbl_style: Str, optional
            Table Styles in excel. Default 'Table Style Medium 6'.
            You can find the style names on Microsoft Excel.
        Additional keyword arguments will be passed as keywords to Table object.
        '''
        kwargs['color'] = color
        if width is not None:
            kwargs['width'] = width
        if transparency is not None:
            kwargs['transparency'] = transparency
        
        par = {'visible': visible, 'line':kwargs}
        if axis == 'y':
            self.workchart.set_y_axis({'major_gridlines': par})
        elif axis == 'x':
            self.workchart.set_x_axis({'major_gridlines': par})
        else:
            raise ValueError('Argurment `axis` should be `x` or `y`, passed \
                             `%s`'%str(axis))
            
class Table(object):
    """
    Add a new table to the Excel worksheet.
    
    Parameters
    -----------
    df: DataFrame, optional
        Data to WorkSheet.
    table_name: str
        Reserved field.
    index: boolean, default True, optional
        Write index.
    loc_table: tuple: (row, None, column, None), default None
        Location to insert.
    tbl_style: str, default "Table Style Medium 6"
        The name of table style in excel.
    total_func: dict
        eg: {'cola':{'total_string': 'Total'},'colb':{'total_function': 'sum}}
    tbl_pars: dict
        Options to pass to `xlsxwriter.worksheet.add_table`.
    merge_format: str
        Fomart for table_name.
    kwargs: keywords
        Options to pass to `DataFrame.to_excel`.
    """
    def __init__(self, df, table_name=None, index=True, loc_table=None, 
                 sheet=None, show_tablename=True, **kwargs):
        if not isinstance(df, pd.DataFrame):
            raise TypeError("The argument 'df' should be a DataFrame object")
        self.table_name = table_name
        self.show_tablename = show_tablename
        self.index = index
        self.sheet = sheet
        self.df = self._check_df_(df.copy())
        if self.sheet is None:
            self.startloc_tbl = self.startloc_chart = loc_table
            self.sheetname = None
        else:
            self.sheetname = self.sheet.sheetname
            if loc_table is None:
                self.startloc_tbl = (0, None, 0, None)
            else:
                if len(loc_table) == 2:
                    loc_table = (loc_table[0], None, loc_table[1], None)
                self.startloc_tbl = loc_table
            self.startloc_chart = self.sheet.loc_chart_next
            self.workbook = self.sheet.workbook
            self.worksheet = self.sheet.worksheet
        self._set_loc_(self.startloc_tbl)
        
        if 'merge_format' in kwargs:
            merge_format = kwargs.pop('merge_format')
        else:
            merge_format = None
        if self.show_tablename:
            self._insert_title_(merge_format)
        self._set_heading_()
        if not (self.sheet is None):
            self._insert_table_(**kwargs)
        self.charts = {}
        self.format_header()
        
    def _check_df_(self, df):
        if self.index:
            if isinstance(df.index, MultiIndex):
                df.index.names = [' ' if idx is None else idx for idx in df.index.names]
            elif df.index.name is None:
                df.index.name = ' '
        if isinstance(df, Series):
            df = DataFrame(df)
        elif isinstance(df, DataFrame):
            if isinstance(df.columns, MultiIndex):
                raise ValueError("'Table' object doesn't support multiindex for"
                                 " column")
        else:
            raise TypeError("'Table' object doesn't support %s"%type(df))
        return df
    
    def _insert_title_(self, merge_format=None):
        # 'fg_color': '#D7E4BC'
        fm = {'bold': True, 'align': 'center', 'valign': 'vcenter'}
        if not (merge_format is None):
            fm.update(merge_format)
        loc_header = deepcopy(self.loc_tbl[0])
        loc_header[0] -= 1
        loc_header[1] = loc_header[0]
        loc_xls = range_to_xl(*loc_header)
        fm = self.workbook.add_format(fm)
        self.worksheet.merge_range(loc_xls, self.table_name, fm)

    def _set_heading_(self):
        self.columns_df = list(self.df.columns)
        nb_idx = len(self.loc_index)
        if nb_idx == 0:
            self.headings = list(map(str, self.columns_df))
        else:
            if nb_idx == 1:
                if self.df.index.name is None:
                    self.df.index.name = ' '
                head_idx = [self.df.index.name]
            else:
                head_idx = [' '*i if n is None else n 
                            for i, n in enumerate(self.df.index.names)]
                self.df.index.names = head_idx
            self.headings = list(map(str, head_idx + self.columns_df))
 
        
    def _format_table_(self, tbl_style=None, total_func=None, tbl_pars=None,
                       **kwargs):
        """Multi-index is unsupported."""
        if isinstance(self.df.index, MultiIndex):
            warn('Multi-Index is unsupported for Excel Table Formatting.')
        else:
            if tbl_style is None:
                tbl_style = 'Table Style Medium 6'
            if tbl_pars is None:
                tbl_pars = {}
#            if not (self.table_name is None):
#                tbl_pars.setdefault('name', self.table_name)
            colrg = self.get_loc_table()
            tbl_pars.update({'autofilter': False, 'style': tbl_style})
            if total_func is None:
                dicheader = [{'header': col_} for col_ in self.headings]
            else:
                func_key = lambda col_: list(total_func[col_].keys())[0]
                func_val = lambda col_: list(total_func[col_].values())[0]
                dicheader = [{'header': col_, func_key: func_val}
                             for col_ in self.headings]
                tbl_pars['total_row'] = True
            tbl_pars['columns'] = dicheader
            self.tbl_pars = (colrg, tbl_pars)
            self.table = self.sheet.worksheet.add_table(colrg, tbl_pars)  
            
    def format_header(self, dicformat=None):
        if dicformat is None:
            self.header_format = self.workbook.add_format({})
            self.header_format.set_align('center')
            self.header_format.set_align('vcenter')
            self.header_format.set_text_wrap()
#            dicformat = {'text_wrap': True, 'align': 'center', 'valign': 'vcenter'}
        else:
            self.header_format = self.workbook.add_format(dicformat)
        header_row = self.loc_tbl[0][0]
        header_size = self.worksheet.row_sizes.get(header_row, 15)
        self.worksheet.set_row(header_row, header_size, self.header_format)

#    def _df2excel_(self, **kwargs):
#        self.df.to_excel(self.sheet.pdwriter, self.sheetname, 
#                         index=self.index, startrow=self.startloc_tbl[0],
#                         startcol=self.startloc_tbl[2], **kwargs)

    def _insert_table_(self, tbl_style=None, total_func=None, tbl_pars=None, 
                       **kwargs):
        """
        Parameters
        -----------
        tbl_style: str, default "Table Style Medium 6"
            The name of table style in excel.
        total_func: dict
            {'cola':{'total_string': 'Total'},'colb':{'total_function': 'sum}}
        tbl_pars: dict
            Options to pass to `xlsxwriter.worksheet.add_table`.
        kwargs: keywords
            Options to pass to `DataFrame.to_excel`.
        """
        self.df.to_excel(self.sheet.pdwriter, self.sheetname, 
                         index=self.index, startrow=self.startloc_tbl[0],
                         startcol=self.startloc_tbl[2], **kwargs)
        self._format_table_(tbl_style, total_func, tbl_pars)

    def _set_loc_(self, startloc=None):
        if startloc is None:
            startloc = [0, None, 0, None]
        if self.show_tablename:
            startloc[0] += 1
        start_row, _, start_col, _ = startloc
        len_val, len_key = self.df.shape
        end_row = start_row + len_val

        if self.index:
            if isinstance(self.df.index, MultiIndex):
                len_index = len(self.df.index.levshape)
            else:
                len_index = 1
            self.loc_index = [start_row,end_row,start_col,start_col+len_index-1]
        else:
            len_index = 0
            self.loc_index = None

        end_col = start_col + len_index + len_key - 1

        # The location of index.
        self.loc_index = []
        for i in range(len_index):
            loc_ = [start_row, end_row, start_col+i, start_col+i]
            self.loc_index.append(loc_)

        # The location of values.
        self.loc_val = []
        loc_col_base = start_col+len_index
        for i, col in enumerate(self.df.columns):
            loc_ = [start_row, end_row, loc_col_base+i, loc_col_base+i]
            self.loc_val.append(loc_)

        # The location of table, index, values.
        self.loc_tbl = [[start_row, end_row, start_col, end_col],
                        [start_row, end_row, start_col,start_col+len_index-1],
                        [start_row, end_row, loc_col_base, end_col]]

    def get_columns(self, column_name):
        index = [i for i, col in enumerate(self.df.columns) if col==column_name]
        return index

    def _range_to_xl_(self, location, if_sheetname=False):
        sheetname = self.sheetname if if_sheetname else None 
        return range_to_xl(*location, sheetname)

#    def _range_to_xl_multi(self, locations, if_sheetname=False):
#        loc = [self._range_to_xl_(*l, if_sheetname=if_sheetname) for l in locations]
#        return ','.join(loc)

    def get_loc_table(self, if_sheetname=False):
        return self._range_to_xl_(self.loc_tbl[0], if_sheetname=if_sheetname)

    def get_loc_index(self, level=None, if_sheetname=False, if_header=True):
        def _getloc_(loc):
            header = 1 if if_header else 0
            loc[0] += header
            loc = self._range_to_xl_(loc, if_sheetname=if_sheetname)
            return loc
        if level is None:
            return _getloc_(self.loc_tbl[1])
        elif isinstance(level, int):
            return _getloc_(self.loc_index[level])
        else:
            locs = [_getloc_(self.loc_index[l]) for l in level]
            return ','.join(locs)

    def get_loc_value(self, column_name=None, column_i=None, if_columnname=True,
                      if_value=True, if_sheetname=False):
        def select_col_val(loc):
            loc = list(loc)
            if if_columnname and not if_value:
                loc[1] = loc[0]
            elif not if_columnname and if_value:
                loc[0] += 1
            elif not if_columnname and not if_value:
                msg = 'At least one argument to be True between if_columnname & if_value.'
                raise ValueError(msg)
            return tuple(loc)
    
        if column_name is None and column_i is None:
            locs = [self._range_to_xl_(select_col_val(self.loc_tbl[2]))]
        else:
            if not (column_name is None):
                try:
                    if_iter = column_name not in self.df.columns
                except TypeError:
                    if_iter = True
                if not if_iter:
                    idx = self.get_columns(column_name)
                else:
                    idx = []
                    for col in column_name:
                        idx.extend(self.get_columns(col))
            else:
                if isinstance(column_i, int):
                    idx = [column_i]
                else:
                    idx = list(column_i)
            locs = [self._range_to_xl_(select_col_val(self.loc_val[i]),
                                     if_sheetname=if_sheetname) for i in idx]
        return ','.join(locs)     

    def format_cells(self, dicformat=None, **kwargs):
        if dicformat is None:
            dicformat = {}
        dicformat.setdefault('valign', 'vcenter')
        dicformat.update(kwargs)
        fm = self.workbook.add_format(dicformat)
        row_start, row_end = self.loc_tbl[2][:2]
        for row in range(row_start, row_end+1):
            self.sheet.worksheet.set_row(row, None, fm)

    def _width_suggested_(self, values, loc_col, maxwidth):
        func = lambda x: len(str(x).encode('gbk'))
        width = max(map(func, values)) + 2
        width = min(width, maxwidth)
        if loc_col in self.sheet.width:
            width = max(self.sheet.width[loc_col], width)
        return width

    def _dicwidth_suggested_(self, maxwidth = 40, if_index=True, if_value=True):
        new_width = {}
        if self.index and if_index:
            idx_start, idx_end = self.loc_tbl[1][2:]
            for i, j in enumerate(range(idx_start, idx_end+1)):
                width_ = self.df.index.get_level_values(i).tolist()
                width_ = self._width_suggested_(width_, j ,maxwidth)
                new_width[j] = max(15, width_)
        if if_value:
            idx_start, idx_end = self.loc_tbl[1][2:]
            for i, j in enumerate(range(idx_start, idx_end+1)):
                new_width[j] = self.df.iloc[:, i].tolist()
                width_ = self.df.iloc[:, i].tolist() + [self.df.columns[i]]
                new_width[j] = self._width_suggested_(width_, j ,maxwidth)
        return new_width

    def _dicwidth_giving_(self, width, if_index=True, if_value=True):
        new_width = {}
        if self.index and if_index:
            idx_start, idx_end = self.loc_tbl[1][2:]
            for i, j in enumerate(range(idx_start, idx_end+1)):
                new_width[j] = width
        if if_value:
            idx_start, idx_end = self.loc_tbl[2][2:]
            for i, j in enumerate(range(idx_start, idx_end+1)):
                new_width[j] = width
        return new_width

    def format_width(self, width=None, maxwidth=40):
        '''
        Parameters
        -----------
        Set sheet's column width.
        width: None, Int, float or 'auto'.
            if pass a number object, set the giving width.
            if pass 'auto', set width depending on table values.
        '''
        if not (width is None):
            if isinstance(width, Number):
                new_width = self._dicwidth_giving_(width, if_index=False)
                new_width_b = self._dicwidth_suggested_(maxwidth, if_value=False)
                new_width.update(new_width_b)
            elif isinstance(width, str) and width.lower() == 'auto':
                new_width = self._dicwidth_suggested_(maxwidth)
            else:
                raise ValueError('Wrong input: %s'%str(width))
            for loc_col, width in new_width.items():
                self.sheet.worksheet.set_column(loc_col, loc_col, width)
            self.sheet.width.update(new_width)

    def add_chart(self, loc_chart = None, chart_name = None, index_col=True, 
                  chart_col=None, chart_type=None, chart_style=None, 
                  y_axis_name=None,legend_pars=None, series_pars=None, 
                  colormap=None, row_space=1, **kwargs):
        if loc_chart is None:
            loc_chart = self.sheet.loc_chart_next
        else:
            loc_chart = _fillnone_(loc_chart, self.sheet.loc_chart_next)
        if chart_name and 'title' not in kwargs:
            kwargs['title'] = chart_name
        chart_name = _add_suffix_(chart_name, 'chart', self.charts)
        chart = Chart(self, loc_chart=loc_chart, chart_name = chart_name, 
                      index_col=index_col, chart_col=chart_col, 
                      chart_type=chart_type, chart_style=chart_style, 
                      y_axis_name=y_axis_name, legend_pars=legend_pars, 
                      series_pars=series_pars, colormap=colormap, **kwargs)
        self.charts[chart_name] = chart
        
        loc_new_chart = chart.loc
        self.sheet._set_next_location_(row_space, loc_new_chart=loc_new_chart)
        return chart

#    def __getitem__(self, item):
#        pass

class Sheet(object):
    '''
    Add a new worksheet to the Excel workbook.
    
    Parameters
    -----------
    pdwriter: ExcelWriter object
        Pandas's ExcelWriter object.
    sheetname: str
        Name of sheet which will contain DataFrame.
    df: DataFrame, optional
        Data to excel.
    index: boolean, default True, optional
        Write index.
    tbl_style: Str, optional
        Table Styles in excel. Default 'Table Style Medium 6'.
        You can find the style names on Microsoft Excel.
    Additional keyword arguments will be passed as keywords to Table object.
    '''
    def __init__(self, pdwriter, sheetname=None, df=None, table_name=None, 
                 index=True, tbl_style=None, width='auto', dicformat=None, 
                 **kwargs):
        if isinstance(pdwriter, ExcelWriter):
            self.pdwriter = pdwriter.pdwriter
        elif isinstance(pdwriter, pd.ExcelWriter):
            self.pdwriter = pdwriter
        elif isinstance(pdwriter, Sheet):
            self.pdwriter = pdwriter.pdwriter
        elif isinstance(pdwriter, (Table, Chart)):
            self.pdwriter = pdwriter.sheet.pdwriter
        elif isinstance(pdwriter, str) and pdwriter.endswith(('xls','xlsx')):
            self.pdwriter = ExcelWriter(pdwriter).pdwriter
        else:
            err = "The argument 'pdwriter' should be an ExcelWriter object"
            raise TypeError(err)

        self.workbook = self.pdwriter.book
        
        self.sheetname = self._check_name_(sheetname)
        self.worksheet = self._add_worksheet_()
        self.worktable = None

        self._height_chart_default = 12

        self.loc_table_next = [0, 0, 0, 0]
        self.loc_chart_next = [0, 0, 2, 2]
        self.tables = {}
        self.width = {}

        if df is not None:
            self.df = df.copy()
            self.worktable = self.add_table(df, table_name, index, 
                                            tbl_style=tbl_style, width=width, 
                                            dicformat=dicformat, **kwargs)

    def _add_worksheet_(self):
        if self.pdwriter.engine == 'xlwt':
            add_worksheet = self.workbook.add_sheet
        elif self.pdwriter.engine == 'xlsxwriter':
            add_worksheet = self.workbook.add_worksheet
        worksheet = add_worksheet(self.sheetname)
        self.pdwriter.sheets[self.sheetname] = worksheet
        return worksheet

    def _check_name_(self, sheetname=None):
        '''
        Check if sheetname is duplicate.
        Generate a sheetname if sheetname is None.
        '''
        if sheetname is None:
            return _add_suffix_(prefix='Sheet', keys=self.pdwriter.sheets)
        else:
            #sheetname2 = xu.quote_sheetname(sheetname)
            sheetname2 = sheetname
            if sheetname2 != sheetname:
                warn('Change sheet name into excel format.')
            if sheetname2 in self.pdwriter.sheets.keys():
                raise ValueError("Sheetname '%s' already exists"%sheetname2)
            else:
                return sheetname2

    def _set_next_location_(self, row_space=1, loc_new_tbl=None, loc_new_chart=None):
        row_space += 1
        if not (loc_new_tbl is None):
            #Set table's row.
            self.loc_table_next[0] = max(loc_new_tbl[1]+row_space, self.loc_table_next[0])
            #The right side of table.
            self.loc_table_next[3] = max(self.loc_table_next[3], loc_new_tbl[3])
        if not (loc_new_chart is None):
            #Set chart's row.
            # When all location values are None, use default chart height.
            if all(map(lambda x:x is None,loc_new_chart)):
                self.loc_chart_next[0] += (self._height_chart_default + row_space)
            else:
                self.loc_chart_next[0] = max(loc_new_chart[1]+row_space, self.loc_chart_next[0])
        # Reset chart's column to make sure the chart is on the right of tables.
        self.loc_chart_next[2] = max(self.loc_chart_next[2], self.loc_table_next[3] + row_space)

    def add_table(self, df, table_name=None, index=True, start_loc=None,
                  tbl_style=None, total_func=None, tbl_pars=None, row_space=1,
                  width='auto', dicformat=None, dicformat_header=None, 
                  show_tablename=None, **kwargs):
        if start_loc is None:
            start_loc = self.loc_table_next
        else:
             start_loc = _fillnone_(start_loc, self.loc_table_next)
        tbl_name = _add_suffix_(table_name, 'table', self.tables)
        if table_name is None and show_tablename is None:
            show_tablename = False
        else:
            show_tablename = True
        table = Table(df, tbl_name, index, start_loc, sheet=self, 
                      tbl_style=tbl_style, total_func=total_func, 
                      tbl_pars=tbl_pars, show_tablename=show_tablename,
                      **kwargs)
        # Autofit width.
        table.format_width(width=width)
        table.format_cells(dicformat=dicformat)
        table.format_header(dicformat_header)
        self._set_next_location_(row_space, table.loc_tbl[0])
        self.tables[tbl_name] = table
        return table

    def format_cells(self, start, end, dicformat=None, width=None, axis=0):
        '''
        Format rows or columns.
        
        Parameters
        -----------
        dicformat: dict
            [eg] {'num_format':'0.00%'}
        start,end: int.
            Index of row/column.
        axis: {0,1, 'row', 'column'}
            0 for row, 1 for column.
        '''
        if dicformat is None:
            fm = None
        else:
            fm = self.add_format(dicformat)
        for i in range(start, end+1):
            if axis in (0, 'row'):
                self.worksheet.set_row(i, width, fm)
            elif axis in (1, 'column'):
                self.worksheet.set_column(i, i, width, fm)

    #@convert_cell_args
    def format_cell(self, row_start, row_end, column_start, column_end, cellformat):
        '''
        Format a cell.

        Parameters
        -----------
        cellformat: xlsxwriter.format.Format or dict object
            [e.g.] {'num_format':'0.00%'}
        '''
        for row in range(row_start, row_end+1):
            for column in range(column_start, column_end+1):
                cell = self.worksheet.table[row][column]
                if not isinstance(cellformat, Format):
                    cellformat = self.add_format(cellformat)
                self.worksheet.table[row][column] = cell._replace(format=cellformat)

    def add_format(self, cellformat):
        return self.workbook.add_format(cellformat)
    
    
class ExcelWriter(object):
    '''Class for writing DataFrame objects into excel sheets.
    
    Parameters
    -----------
    file_path : string
        Path to xls or xlsx file.
    sheetname: str.
        Sheetname for the table on excel file.
    date_format : string, default None
        Format string for dates written into Excel files (e.g. 'YYYY-MM-DD')
    datetime_format : string, default None
        Format string for datetime objects written into Excel files
        (e.g. 'YYYY-MM-DD HH:MM:SS')
    '''
    def __init__(self, file_path, date_format=None, datetime_format=None, 
                 sheetname=None, kwargs_sheet=None, *args, **kwargs):
        if isinstance(file_path, pd.ExcelWriter):
            self.pdwriter = file_path
        elif isinstance(file_path, Sheet):
            self.pdwriter = file_path.pdwriter
        elif isinstance(file_path, (Table, Chart)):
            self.pdwriter = file_path.sheet.pdwriter
        elif isinstance(file_path, ExcelWriter):
            self.pdwriter = file_path.pdwriter
        else:
            self.pdwriter = pd.ExcelWriter(file_path, engine='xlsxwriter', 
                                           date_format=date_format, 
                                           datetime_format=datetime_format,
                                           *args, **kwargs)
        if sheetname is not None:
            kwargs_sheet = {} if kwargs_sheet is None else kwargs_sheet
            self.sheetname = self.add_sheet(sheetname, **kwargs_sheet)
        self.sheets = self.pdwriter.sheets

    def add_sheet(self, *args, **kwargs):
        '''
        Add a new worksheet to the Excel workbook.
        
        Parameters
        -----------
        sheetname: str
            Name of sheet which will contain DataFrame.
        df: DataFrame, optional
            Data to excel.
        index: boolean, default True, optional
            Write index.
        tbl_style: Str, optional
            Table Styles in excel. Default 'Table Style Medium 6'.
            You can find the style names on Microsoft Excel.
        Additional keyword arguments will be passed as keywords to Table object.
        '''
        return Sheet(self.pdwriter, *args, **kwargs)
        
    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, exc_tb):
        self.close()

    def close(self):
        self.pdwriter.close()

    def save(self):
        return self.pdwriter.save()

def to_excel(writer, tables, sheet=None, chart=True, chart_type='line', 
             kwargs_sheet=None, kwargs_table=None, kwargs_cell=None, 
             kwargs_chart=None, **kwargs):
    '''
    Write data to the Excel.
    
    Parameters
    -----------
    tables: DataFrame, dict of DataFrame, list of DataFrame
        Table to the excel.
    sheet: str, Sheet object
        Create a new sheet if pass a sheetname.
        Add the table to the existed sheet if pass a Sheet object.
    df: DataFrame, optional
        Data to excel.
    chart: boolean, default True
        Chart in excel if passed True.
    kwargs_sheet: dict
        Options to pass to ExcelWriter.add_sheet.
    kwargs_table: dict
        Options to pass to Sheet.add_table.
    kwargs_cell: dict
        Options to pass to Table.format_cells.
    kwargs_chart: dict
        Options to pass to Table.add_chart.
    Additional keyword arguments will be passed as keywords to Table object.
    
    Return
    -----------
    Sheet object.
    '''

    n2dic = lambda x,type=dict:type() if x is None else x
    kwargs_sheet = n2dic(kwargs_sheet)
    kwargs_table = n2dic(kwargs_table)
    kwargs_chart = n2dic(kwargs_chart)
    kwargs_table.setdefault('dicformat', kwargs_cell)

    if sheet is None and isinstance(tables, str):
        sheet = tables
    if isinstance(tables, (Series, DataFrame)):
        tables = ((None, tables), )
    elif isinstance(tables, (tuple,list,set)):
        tables = tuple(zip([None]*len(tables), tables))
    elif isinstance(tables, dict):
        tables = tuple(tables.items())

    if not isinstance(sheet, Sheet):
        sheet = writer.add_sheet(sheetname=sheet, **kwargs_sheet)
    kwargs.update(kwargs_table)
    for tblname, df in tables:
        if len(df) > 0:
            table = sheet.add_table(df, table_name=tblname, **kwargs)
            if chart:
                table.add_chart(title=tblname, chart_type=chart_type, **kwargs_chart)
    return sheet


if __name__ == "__main__":
    import numpy as np
    dfa = pd.DataFrame(np.random.rand(8, 4), columns=list('ABCD'))
    dfb = pd.DataFrame(np.random.rand(10, 2), columns=list('AB'))

    "============= 简单用法 ============="
    with ExcelWriter('demo.xlsx') as writer2:
        # 表格插入Sheet1, 对A、C字段数据绘折线图
        sheet1 = to_excel(writer2, dfa, kwargs_chart=dict(chart_col=['A', 'C']))
        # 表格插入Sheet1, 不绘图, 数值以百分比格式保存
        sheet1 = to_excel(writer2, dfb, sheet1, chart=False,
                          kwargs_cell={'num_format':'0.00%'})
        # 表格插入Sheet2, 绘直方图
        sheet2 = to_excel(writer2, dfa, chart_type='column')

    "============= 一般用法 ============="
    # 创建Excel文件
    writer = ExcelWriter('demo2.xlsx')

    # 创建工作表Sheet1
    sheet1 = writer.add_sheet(sheetname='Sheet1')
    # 添加表格1
    table11 = sheet1.add_table(dfa, table_name='Table_1')
    # 添加表格2, 数值格式百分比两位小数，列宽设置为8.
    table12 = sheet1.add_table(dfb, dicformat={'num_format':'0.00%'}, width=6)
    # 插入图表, 默认为折线图
    chart11 = table11.add_chart()
    # 插入直方图, 标题'Column Chart'
    chart12 = table12.add_chart(chart_name='Column Chart', chart_type='column')
    # 插入 A、C两列数据的折线图, 高度为默认的2倍
    chart13 = table12.add_chart(chart_col=['A'], y_scale=2)
    
    # 创建工作表Sheet2
    sheet2 = writer.add_sheet()
    # 插入指定风格的表格
    table2 = sheet2.add_table(dfa, tbl_style='Table Style Light 11')
    # 插入指定风格的图表
    chart2 = table2.add_chart(chart_style=37)

    # 退出并保存文件
    writer.close()

    