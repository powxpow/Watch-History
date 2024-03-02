'''
    Watch History Data Handler: Reads the watch history file and 
    converts it into three DataFrames:
    - Views, which is as close to the original document as possible
    - Videos, which collapses the views into a count
    - Channels, which collapses the videos into a count
    The DataFrames can then be used by other parts of the program.
'''
#core
from dataclasses import dataclass
from datetime import datetime
import json
from json import JSONDecodeError
import logging as log
from pathlib import Path
from zipfile import ZipFile
#modules
from dateutil import parser as dateutil_parser
from htmlement import parse as html_parse
from pandas import DataFrame

@dataclass
class ViewRecord:
    '''
    ViewRecord dataclass to structure the data coming from Google.
    Since Google exports in two formats (JSON and HTML), we want to
    standardize the information we are getting from them.
    '''
    channel_title: str
    channel_url: str
    channel_id: str
    video_title: str
    video_url: str
    video_id: str
    view: datetime

class WatchHistoryDataHandler():
    '''WatchHistoryDataHandler'''
    def create_views_df_from_source(self, source_file):
        '''create_views_df_from_source'''
        views_df = None
        src = Path(source_file)

        match src.suffix[1:].lower():
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
                                try:
                                    data = json.load(doc)
                                    views_df = self.create_views_df_json(data)
                                except JSONDecodeError as jerr:
                                    log.error("JSON %s", jerr.msg)
            case 'html':
                with open(src, 'r', encoding='UTF-8') as doc:
                    views_df = self.create_views_df_html(doc)
            case 'json':
                with open(src, 'r', encoding='UTF-8') as doc:
                    try:
                        data = json.load(doc)
                        views_df = self.create_views_df_json(data)
                    except JSONDecodeError as jerr:
                        log.error("JSON %s", jerr.msg)
            case _:
                log.error('Unable to process %s: unrecognized file type', src)

        return views_df

    @staticmethod
    def create_views_df_json(data):
        '''create_views_df_json'''
        views_df = None
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

        if len(views) > 0:
            views_df = DataFrame(views)
            log.info('%7d total records processed', total)
            log.info('%7d ads ignored, %d were surveys',
                total - views_df.shape[0], survey_count)
            log.info('%7d views', views_df.shape[0])

        return views_df

    #div.outer-cell
    #   div.mdl-grid
    #       header-cell: p "YouTube"
    #       content-cell
    #           0 a=survey, 1 a=ad, 2 a=video
    #       content-cell - (no data) and class ends with "mdl-typography--text-right"
    #       content-cell
    #           b "Products" (followed by "YouTube","YouTube Music")
    #           optional b "Details" followed by "From Google Ads" (ads and surveys)
    #           b "Why is this here?" followed by explanation and 1 a link
    @staticmethod
    def create_views_df_html(doc):
        '''create_views_df_html'''
        views_df = None
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

        if len(views) > 0:
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
