'''excelbuilder'''
#core
from dataclasses import dataclass
import logging as log
import os
#modules
from numpy.dtypes import DateTime64DType
from pandas import DataFrame, ExcelWriter, Timestamp
from pandas.core.dtypes.dtypes import DatetimeTZDtype
from tzlocal import get_localzone
from xlsxwriter import __name__ as XLSXWRITER
from xlsxwriter.utility import xl_col_to_name

@dataclass
class Hyperlink:
    '''Hyperlink'''
    title: str = None
    url: str = None

class ExcelBuilder():
    '''ExcelBuilder'''

    def clean_data_for_report(self, channels_df, videos_df, views_df):
        '''
        Cleans the DataFrames for the report by creating new focused DataFrames.
        The new Dataframes only the have columns we want on the report.
        This also turns title/url columns into Hyperlink columns for rendering.
        '''
        ch_df = DataFrame(channels_df, columns=['channel_title', 'channel_url', 'videos'])
        vd_df = DataFrame(videos_df,
            columns=['channel_title', 'channel_url', 'video_title', 'video_url', 'views'])
        vw_df = DataFrame(views_df, columns=['video_title', 'video_url', 'view'])

        #turn title/url columns into hyperlinks for rendering
        ch_df = self.create_hyperlink(ch_df, 'Channel', 'channel_title', 'channel_url')
        vd_df = self.create_hyperlink(vd_df, 'Channel', 'channel_title', 'channel_url')
        vd_df = self.create_hyperlink(vd_df, 'Video', 'video_title', 'video_url')
        vw_df = self.create_hyperlink(vw_df, 'Video', 'video_title', 'video_url')
        return (ch_df, vd_df, vw_df)

    @staticmethod
    def create_hyperlink(a_df, col_label, title_col, url_col):
        '''
        Searches by column names to see if it can create a Hyperlink object for
        channels and videos. This Hyperlink object then can be turned into a 
        clickable hyperlink on the Excel spreadsheet when it is rendered.
        '''
        if title_col in a_df.columns and url_col in a_df.columns:
            idx = a_df.columns.get_loc(title_col)
            a_df.insert(idx, col_label,
                a_df.apply(lambda row: Hyperlink(row[title_col], row[url_col]),
                axis=1))
            a_df.drop(columns=[title_col, url_col], inplace=True)
        return a_df

    def export_spreadsheet(self, filename, dfs):
        '''export_spreadsheet'''
        ch_df, vd_df, vw_df = self.clean_data_for_report(
            dfs['channels_df'], dfs['videos_df'], dfs['views_df'])
        mv_df = dfs['monthlyviews_df']
        channel_widths = [45, 6]
        video_widths = [45, 45, 6]
        views_widths = [45, 19]
        mviews_widths = [8, 6]
        with ExcelWriter(f"{filename}", engine=XLSXWRITER) as writer: # pylint: disable=abstract-class-instantiated
            self.export_sheet(writer.book, 'Channels', channel_widths, ch_df)
            self.export_sheet(writer.book, 'Videos', video_widths, vd_df)
            self.export_sheet(writer.book, 'Views', views_widths, vw_df)
            self.export_sheet(writer.book, "Monthly", mviews_widths, mv_df)
            self.add_graph(writer.book, "Monthly", mv_df)
            home =os.path.expanduser('~')
            log.info('Exported %s', str(filename).replace(home,"~"))

    def add_graph(self, book, sheet_name, a_df):
        '''
        Adds a graph to a sheet.
        '''
        #get sheet by sheet_name
        sheet = None
        if sheet_name in book.sheetnames:
            sheet = book.get_worksheet_by_name(sheet_name)
        else:
            sheet = book.add_worksheet(sheet_name)
        sheet.active = True

        #chart - need to be able to pass in labels later
        chart = book.add_chart({'type': 'column'})
        chart.set_title({'name': 'Views per Month'})
        chart.set_x_axis({'name': 'Month', 'num_format': 'yyyy-MM', 'date_axis': False})
        chart.set_y_axis({'name': 'Count', 'major_unit': 1})
        chart.set_size({ 'width': (10 * a_df.shape[0])})
        v = xl_col_to_name(1) #B
        c = xl_col_to_name(0) #A
        chart.add_series({'values': f'=\'{sheet_name}\'!${v}$2:${v}${a_df.shape[0] + 1}',
                                'categories': f'==\'{sheet_name}\'!${c}$2:${c}${a_df.shape[0] + 1}',
                                'name': sheet_name,
                                'fill': {'color': 'red'},
                                'border': {'color': 'black'}})
        sheet.insert_chart(f'{xl_col_to_name(a_df.shape[1] + 1)}1', chart)

    def export_sheet(self, book, sheet_name, widths, a_df):
        '''export_sheet'''
        #book settings
        book.remove_timezone = True
        bolded = book.add_format({"bold": True})
        vw_fmt = book.add_format({'num_format': 'yyyy-MM-dd hh:mm AM/PM'})
        mo_fmt = book.add_format({'num_format': 'yyyy-MM'})

        #get sheet by sheet_name
        sheet = None
        if sheet_name in book.sheetnames:
            sheet = book.get_worksheet_by_name(sheet_name)
        else:
            sheet = book.add_worksheet(sheet_name)
        sheet.active = True

        #sheet settings
        sheet.set_row(0, None, bolded)
        sheet.add_write_handler(Hyperlink, self.write_hyperlink)
        sheet.add_write_handler(Timestamp, self.write_local_datetime)
        for idx, col in enumerate(a_df.columns):
            width = 12
            if len(widths) > idx:
                width = widths[idx]

            if isinstance(a_df.dtypes[col], (DatetimeTZDtype, DateTime64DType)):
                fmt = vw_fmt
                if col == 'month':
                    fmt = mo_fmt
                sheet.set_column(idx, idx, width, fmt)
            else:
                sheet.set_column(idx, idx, width)

        #title row
        sheet.write_row(0, 0, [c.replace('_',' ').title() for c in a_df.columns])

        #data rows
        for idx, row_data in enumerate(a_df.itertuples(index=False)):
            sheet.write_row(idx + 1, 0, row_data, None)

    def write_hyperlink(self, worksheet, row, col, link:Hyperlink, _):
        '''write_xlsx_hyperlink'''
        return worksheet.write_url(row, col, url=link.url, string=link.title)

    def write_local_datetime(self, worksheet, row, col, ts, _):
        '''
        Excel doesn't like datetimes with timezones.
        Convert the timezone aware Timestamps to local timezone
        for Excel and the end user.
        '''
        local_datetime = ts
        if ts.tzinfo is not None:
            ts.astimezone(get_localzone())
        return worksheet.write_datetime(row, col, local_datetime)
