"""
SignalHook is to allow the logging functions
to also be sent to the PySide UI
"""
import logging as log


class SignalHook(log.StreamHandler):
    """
    SignalHook is to allow the logging functions
    to also be sent to the PySide UI
    """
    signal = None

    def __init__(self, stream=None, signal=None):
        self.signal = signal
        super().__init__(stream)

    def emit(self, record: log.LogRecord):
        """
            emit: extend the StreamHandler.emit() 
            to also send the signal to the PySide UI

            :param record: a log record
            :type loging.LogRecord
        """
        if self.signal is not None:
            if record.levelno == log.INFO:
                self.signal.emit(f'{record.message}')
            else:
                self.signal.emit(f'{record.levelname} {record.message}')
        else:
            super().emit(record)
