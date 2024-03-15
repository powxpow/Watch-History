'''excelbuilder'''
#core
from dataclasses import dataclass
import logging as log
import os
#modules
from pandas import DataFrame, ExcelWriter, Timestamp
from pandas.core.dtypes.dtypes import DatetimeTZDtype
from tzlocal import get_localzone
from xlsxwriter import __name__ as XLSXWRITER

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

    def export_spreadsheet(self, channels_df, videos_df, views_df, filename):
        '''export_spreadsheet'''
        ch_df, vd_df, vw_df = self.clean_data_for_report(channels_df, videos_df, views_df)
        channel_widths = [45, 6]
        video_widths = [45, 45, 6]
        views_widths = [45, 19]
        with ExcelWriter(f"{filename}", engine=XLSXWRITER) as writer: # pylint: disable=abstract-class-instantiated
            self.export_sheet(writer.book, 'Channels', channel_widths, ch_df)
            self.export_sheet(writer.book, 'Videos', video_widths, vd_df)
            self.export_sheet(writer.book, 'Views', views_widths, vw_df)
            home =os.path.expanduser('~')
            log.info('Exported %s', str(filename).replace(home,"~"))

    def export_sheet(self, book, sheet_name, widths, a_df):
        '''export_sheet'''
        #book settings
        book.remove_timezone = True
        bolded = book.add_format({"bold": True})
        vw_fmt = book.add_format({'num_format': 'yyyy-MM-dd hh:mm AM/PM'})

        #get sheet by sheet_name
        sheet = None
        if sheet_name in book.sheetnames:
            sheet = book.get_worksheet_by_name(sheet_name)
        else:
            sheet = book.add_worksheet(sheet_name)

        #sheet settings
        sheet.active = True
        sheet.set_row(0, None, bolded)
        sheet.add_write_handler(Hyperlink, self.write_hyperlink)
        sheet.add_write_handler(Timestamp, self.write_local_datetime)
        for idx, col in enumerate(a_df.columns):
            width = 12
            if len(widths) > idx:
                width = widths[idx]

            if isinstance(a_df.dtypes[col], DatetimeTZDtype):
                sheet.set_column(idx, idx, width, vw_fmt)
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
        local_datetime = ts.astimezone(get_localzone())
        return worksheet.write_datetime(row, col, local_datetime)
