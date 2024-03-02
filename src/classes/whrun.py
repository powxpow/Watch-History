'''Watch History Run: 
   Does some handling of the path before calling Watch History Data.
   If creating the Views DataFrame is succesful, it then calls Watch
   History Data to create the other DataFrames (Videos, Channels).
   Finally, if all the data was created properly, it will call the 
   Spreadsheet Renderer (currently only whexcel) to create the spreadsheet.'''
#core
import logging as log
from pathlib import Path
#modules
from pandas import DataFrame

class WatchHistoryRun():
    '''WatchHistoryRun'''
    whdf = None
    spreadsheet = None
    def __init__(self, log_handler=None, data_handler=None, spreadsheet=None):
        if log_handler is not None:
            log.getLogger().addHandler(log_handler)
        self.whdf = data_handler
        self.ss = spreadsheet

    @staticmethod
    def get_source_path(source_file):
        '''get_source_path'''
        src = None
        spath = source_file
        if isinstance(source_file, str):
            spath = Path(source_file)
        if isinstance(spath, Path):
            if spath.exists():
                if not spath.is_file():
                    log.error("%s is not a file", source_file)
                else:
                    src = spath
                    log.info("Found '%s'", src.name)
            else:
                log.error("Unable to locate '%s'", source_file)
        return src

    def run(self, source_file, dest_file):
        '''run'''
        src = self.get_source_path(source_file)
        if src is not None:
            wh = self.whdf
            views_df = wh.create_views_df_from_source(src)
            channels_df= None
            videos_df = None
            if views_df is not None:
                log.info('Creating video records')
                videos_df = wh.create_videos_df(views_df)
                log.info('%7d total videos', videos_df.shape[0])
                log.info('%7d views of already watched videos',
                    views_df.shape[0] - videos_df.shape[0])
                log.info('Creating channel records')
                channels_df = wh.create_channels_df(videos_df)
                log.info('%7d channels', channels_df.shape[0])
            if isinstance(channels_df, DataFrame) and self.ss is not None:
                self.ss.export_spreadsheet(channels_df, videos_df, views_df, dest_file)
            else:
                log.info("No Watched History data to work with.")
                log.info("Done.")
