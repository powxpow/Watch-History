'''watch_history_core.py'''
import argparse
from dataclasses import dataclass
from datetime import datetime
import json
import logging as log
import os
from pathlib import Path
import sys
from zipfile import ZipFile
from dateutil import parser as dateutil_parser
from htmlement import parse as html_parse
from pandas import DataFrame, ExcelWriter, Timestamp
from xlsxwriter import __name__ as XLSXWRITER

@dataclass
class ViewRecord:
    '''ViewRecord'''
    channel_title: str
    channel_url: str
    channel_id: str
    video_title: str
    video_url: str
    video_id: str
    view: datetime

@dataclass
class Hyperlink:
    '''Hyperlink'''
    title: str = None
    url: str = None

class Feedback(log.StreamHandler):
    '''Feedback'''
    signal = None

    def __init__(self, stream=None, signal=None):
        self.signal = signal
        super().__init__(stream)

    def emit(self, record:log.LogRecord):
        '''emit'''
        if self.signal is not None:
            if record.levelno == log.INFO:
                self.signal.emit(f'{record.message}')
            else:
                self.signal.emit(f'{record.levelname} {record.message}')
        else:
            super().emit(record)

class WatchHistory():
    '''WatchHistory'''
    def __init__(self, log_handler=None):
        if log_handler is not None:
            log.getLogger().addHandler(log_handler)

    def check_parameters(self, source, out_dir):
        '''check_parameters'''
        out = None
        src = Path(source)

        if src.is_file():
            log.info("Found '%s'", src.name)
            out_file = src.name.replace(src.suffix, '.xlsx')
            out = Path(out_dir, out_file)
        else:
            log.error("Could not find '%s'.", src)
            src = None
        return src, out

    def process_yt_history(self, src, out):
        '''process_yt_history'''
        if isinstance(src, str):
            src = Path(src)
        log.info('Processing %s', src.name)
        views_df = self.create_views_df_from_source(src)
        if views_df is not None:
            log.info('Creating video records')
            videos_df = self.create_videos_df(views_df)
            log.info('%7d views of already watched videos', views_df.shape[0] - videos_df.shape[0])
            log.info('Creating channel records')
            channels_df = self.create_channels_df(videos_df)
            log.info('%7d channels', channels_df.shape[0])

            exporter = ExcelBuilder()
            exporter.export_spreadsheet(channels_df, videos_df, views_df, out)
        log.info("Done")

    #PROCESS DATA SECTION
    def create_views_df_from_source(self, src):
        '''create_views_df_from_source'''
        views_df = None

        match Path(src).suffix[1:].lower():
            case 'zip':
                with ZipFile(src) as azip:
                    for file in azip.filelist:
                        found_html = file.filename.endswith('watch-history.html')
                        found_json = file.filename.endswith('watch-history.json')
                        if found_html:
                            with azip.open(file) as doc:
                                views_df = self.create_views_df_html(doc)
                        elif found_json:
                            with azip.open(file) as doc:
                                data = json.load(doc)
                            views_df = self.create_views_df_json(data)
            case 'html':
                with open(src, 'r', encoding='UTF-8') as doc:
                    views_df = self.create_views_df_html(doc)
            case 'json':
                with open(src, 'r', encoding='UTF-8') as doc:
                    data = json.load(doc)
                views_df = self.create_views_df_json(data)
            case _:
                log.error('Unable to process %s: unrecognized file type', src)

        return views_df

    @staticmethod
    def create_views_df_json(data):
        '''create_views_df_json'''
        total = len(data)
        views = []
        data_views = [rec for rec in data if 'subtitles' in rec]
        survey_count = 0

        for rec in data_views:
            channel = rec['subtitles'][0]
            if 'url' in channel:
                #get ids
                ch_url = channel.get('url')
                ch_id  = ch_url.split("/channel/", 1)[1] if '/channel/' in ch_url else ch_url
                vd_url = rec.get('titleUrl')
                vd_id = vd_url.split("?v=",1)[1] if '?v=' in vd_url else vd_url

                view_record = ViewRecord(
                    channel_title=channel.get('name'),
                    channel_url=ch_url,
                    channel_id=ch_id,
                    video_title=rec.get('title').replace('Watched ',''),
                    video_url=vd_url,
                    video_id=vd_id,
                    view=dateutil_parser.isoparse(rec.pop('time'))
                )
                views.append(view_record)
            else:
                #print(rec)
                survey_count += 1

        views_df = DataFrame(views)
        log.info('%7d total records processed', total)
        log.info('%7d ads ignored', total - views_df.shape[0])
        log.info('%7d survey ads encountered', survey_count)
        log.info('%7d views', views_df.shape[0])
        return views_df

    #div.outer-cell
    #   div.mdl-grid
    #       header-cell: p "YouTube"
    #       content-cell
    #           0 a=survey, 1 a=ad, 2 a=video
    #       content-cell - (no data) and class ends with "mdl-typography--text-right"
    #       content-cell
    #           b "Products" (always followed by "YouTube")
    #           optional b "Details" followed by "From Google Ads" (ads and surveys)
    #           b "Why is this here?" followed by explanation and 1 a link
    @staticmethod
    def create_views_df_html(doc):
        '''create_views_df_html'''
        views = []
        idx = 0
        root = html_parse(doc, encoding="UTF-8")

        for outer_cell in root.iterfind(".//div[@class]"):
            if outer_cell.get('class').startswith('outer-cell'):
                idx += 1
                div = outer_cell.find('.//div[1]/div[@class][2]')
                channel_alink = div.find(".//a[2]")
                if channel_alink is not None:
                    #now process the video view
                    video_alink = div.find(".//a[1]")
                    view_date = div.find(".//br[2]").tail.replace('\u202f', ' ')
                    #get ids
                    ch_url = channel_alink.get('href')
                    ch_id = ch_url.split("/channel/", 1)[1] if "/channel/" in ch_url else ch_url
                    vd_url = video_alink.get('href')
                    vd_id = vd_url.split("?v=",1)[1] if '?v=' in vd_url else vd_url

                    view_record = ViewRecord(
                        channel_title=channel_alink.text,
                        channel_url=ch_url,
                        channel_id=ch_id,
                        video_title=video_alink.text,
                        video_url=vd_url,
                        video_id=vd_id,
                        view=dateutil_parser.parse(view_date, fuzzy=True)
                    )
                    views.append(view_record)

        views_df = DataFrame(views)
        log.info('%7d total records processed', idx)
        log.info('%7d ads ignored', idx - views_df.shape[0])
        log.info('%7d views', views_df.shape[0])
        return views_df

    @staticmethod
    def create_videos_df(views_df):
        '''create_videos_df'''
        videos_df = DataFrame(
            views_df,
            columns=['channel_id', 'channel_title', 'channel_url',
                     'video_id', 'video_title', 'video_url']).drop_duplicates()
        # video_url is more unique than video_id
        # to get the right counts 'music.youtube' needs to be counted separately from 'www.youtube'
        videos_df.loc[:, 'views'] = videos_df.loc[:, 'video_url'].map(
            views_df.loc[:, 'video_url'].value_counts())
        videos_df = videos_df.sort_values(by='views', ascending=False)
        return videos_df

    @staticmethod
    def create_channels_df(videos_df):
        '''create_channels_df'''
        channels_df = DataFrame(
            videos_df,
            columns=['channel_id', 'channel_title', 'channel_url']).drop_duplicates()
        channels_df.loc[:, 'videos'] = channels_df.loc[:, 'channel_url'].map(
            videos_df['channel_url'].value_counts())
        channels_df = channels_df.sort_values(by='videos', ascending=False)
        return channels_df

class ExcelBuilder():
    '''ExcelBuilder'''

    @staticmethod
    def create_hyperlinks(a_df):
        '''create_hyperlinks'''
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
        ch_df = self.create_hyperlinks(DataFrame(channels_df,
            columns=['channel_title', 'channel_url', 'videos']))
        vd_df = self.create_hyperlinks(DataFrame(videos_df,
            columns=['channel_title', 'channel_url', 'video_title', 'video_url', 'views']))
        vw_df = self.create_hyperlinks(DataFrame(views_df,
            columns=['video_title', 'video_url', 'view']))
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
        vw_fmt = writer.book.add_format({'num_format': 'dd/mm/yyyy hh:mm AM/PM'})

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

#READ DATA SECTION
class MainHandler:
    '''MainHandler'''
    @staticmethod
    def get_parameters():
        '''get_parameters'''
        args = MainHandler.get_args()
        source_prompt = "Enter Google Takeout file path"
        src_default = "Google Takeout/takeout-20240226T234658Z-001.zip"
        source = args.source_file or MainHandler.get_from_user(source_prompt, src_default)
        output_dir_prompt = "Enter Output directory"
        out_dir = args.output_dir or MainHandler.get_from_user(output_dir_prompt, ".")
        return source, out_dir

    @staticmethod
    def get_args():
        '''get_args'''
        desc = "Process Google Takeout file and output spreadsheet to directory."
        parser = argparse.ArgumentParser(description=desc)
        parser.add_argument("source_file", nargs="?", help="Google Takeout file")
        parser.add_argument("output_dir", nargs="?", help="Output directory")
        args = parser.parse_args()
        return args

    @staticmethod
    def get_from_user(prompt, default=None):
        '''get_from_user'''
        if default:
            prompt = f"{prompt} (default: {default}): "
        else:
            prompt = f"{prompt}: "
        user_input = input(prompt)
        return user_input or default

def main():
    '''main'''
    log_handler = Feedback(sys.stdout)
    log_fmt = '%(asctime)s %(levelname)s\t%(message)s'
    log.basicConfig(level=log.INFO, format=log_fmt, handlers=[log_handler])

    src, out_dir = MainHandler.get_parameters()

    watch_history = WatchHistory()
    src_file, out_file = watch_history.check_parameters(src, out_dir)
    if src_file is not None:
        watch_history.process_yt_history(src_file, out_file)

if __name__ == '__main__':
    main()
