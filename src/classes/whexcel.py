'''excelbuilder'''
#core
from dataclasses import dataclass
from datetime import datetime
import logging as log
import os
#modules
from pandas import DataFrame, ExcelWriter, Timestamp
from xlsxwriter import __name__ as XLSXWRITER

@dataclass
class Hyperlink:
    '''Hyperlink'''
    title: str = None
    url: str = None

class ExcelBuilder():
    '''ExcelBuilder'''

    @staticmethod
    def clean_data_for_report(channels_df, videos_df, views_df):
        '''
        Cleans the DataFrames for the report by returning a new DataFrames 
        that only include needed columns for reporting. For example: 
        video_id and channel_id are not included for the end user's spreadsheet.
        '''
        ch_df = DataFrame(channels_df, columns=['channel_title', 'channel_url', 'videos'])
        vd_df = DataFrame(videos_df,
            columns=['channel_title', 'channel_url', 'video_title', 'video_url', 'views'])
        vw_df = DataFrame(views_df, columns=['video_title', 'video_url', 'view'])
        return (ch_df, vd_df, vw_df)

    @staticmethod
    def create_hyperlinks(a_df):
        '''
        Searches by column names to see if it can create a Hyperlink object for
        channels and videos. This Hyperlink object then can be turned into a 
        clickable hyperlink on the Excel spreadsheet when it is rendered.
        '''
        if 'channel_url' in a_df.columns and 'channel_title' in a_df.columns:
            idx = a_df.columns.get_loc("channel_title")
            a_df.insert(idx, "Channel",
                a_df.apply(lambda row: Hyperlink(row["channel_title"], row["channel_url"]),
                axis=1))
            a_df.drop(columns=['channel_title', 'channel_url'], inplace=True)
        if 'video_url' in a_df.columns and 'video_title' in a_df.columns:
            idx = a_df.columns.get_loc("video_title")
            a_df.insert(idx, "Video",
                a_df.apply(lambda row: Hyperlink(row["video_title"], row["video_url"]),
                axis=1))
            a_df.drop(columns=['video_title', 'video_url'], inplace=True)
        return a_df

    def export_spreadsheet(self, channels_df, videos_df, views_df, filename):
        '''export_spreadsheet'''
        ch_df, vd_df, vw_df = self.clean_data_for_report(channels_df, videos_df, views_df)
        ch_df = self.create_hyperlinks(ch_df)
        vd_df = self.create_hyperlinks(vd_df)
        vw_df = self.create_hyperlinks(vw_df)
        channel_widths = [45, 6]
        video_widths = [45, 45, 6]
        views_widths = [45, 19]
        with ExcelWriter(f"{filename}", engine=XLSXWRITER) as writer: # pylint: disable=abstract-class-instantiated
            self.export_sheet(writer, 'Channels', channel_widths, ch_df)
            self.export_sheet(writer, 'Videos', video_widths, vd_df)
            self.export_sheet(writer, 'Views', views_widths, vw_df)
            home =os.path.expanduser('~')
            log.info('Exported %s', str(filename).replace(home,"~"))

    def export_sheet(self, writer, sheet_name, widths, a_df):
        '''export_sheet'''
        #book settings
        writer.book.remove_timezone = True
        bolded = writer.book.add_format({"bold": True})
        vw_fmt = writer.book.add_format({'num_format': 'yyyy-MM-dd hh:mm AM/PM'})

        #get sheet by sheet_name
        sheet = None
        if sheet_name in writer.sheets:
            sheet = writer.sheets[sheet_name]
        else:
            sheet = writer.book.add_worksheet(sheet_name)

        #sheet settings
        sheet.active = True
        sheet.set_row(0, None, bolded)
        sheet.add_write_handler(Hyperlink, self.write_hyperlink)
        for idx, col in enumerate(a_df.columns):
            width = 12
            if len(widths) > idx:
                width = widths[idx]

            if isinstance(a_df[col][0], datetime) or isinstance(a_df[col][0], Timestamp):
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
