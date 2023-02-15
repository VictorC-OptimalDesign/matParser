#!/usr/bin/env python


# === IMPORTS ==================================================================

import math
import os
import typing

from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum


# === CLASSES ==================================================================

@dataclass
class LogEntry:
    line : int = field(init=False, default=None)
    time : datetime = field(init=False, default=None)
    type : str = field(init=False, default=None)
    system : str = field(init=False, default=None)
    data : str = field(init=False, default=None)
    extra : typing.List[str] = field(init=False, default_factory=list)


class CSV:
    _EXTENSION : str = '.csv'
    _SEPARATOR : str = ','
    
    class Offsets(Enum):
        TIME = 0
        TYPE = 1
        SYSTEM = 2
        DATA = 3
    
    def __init__(self, fileName : str):
        self.fileName : str = fileName
        self.filePath : str = None
        self.name : str = None
        self.entries : typing.List[LogEntry] = []
        self._read()
        
    def _read(self):
        if (self.fileName):
            baseName : str = os.path.basename(self.fileName)
            self.name = baseName.replace(self._EXTENSION, '')
            self.filePath = os.path.join(os.getcwd(), self.fileName)
            with open(self.filePath, 'r') as file:
                readLines = file.readlines()
            file.close
            
            for i, line in enumerate(readLines):
                fields : typing.List[str] = [x.strip() for x in line.split(self._SEPARATOR)]
                fields = [i for i in fields if i]
                self._processFields(i, fields)
                pass
                
    def _processFields(self, line : int, fields : typing.List[str]):
        _DATETIME_FORMAT : str = '%Y-%m-%d %H:%M:%S.%f'
        numberOfFields : int = len(fields)
        if (numberOfFields > self.Offsets.TIME.value):
            try:
                timeString : str = fields[0]
                time : datetime = datetime.strptime(timeString, _DATETIME_FORMAT)
                self.entries.append(self._makeLogEntry(line, time, fields))
                pass
            except:
                numberOfEntries : int = len(self.entries)
                if (numberOfEntries > 0):
                    self.entries[numberOfEntries - 1].extra.append(self._SEPARATOR.join(fields))
                pass
                    
    def _makeLogEntry(self, line : int, time : datetime, fields : typing.List[str]) -> LogEntry:
        entry : LogEntry = LogEntry()
        entry.time = time
        
        numberOfFields : int = len(fields)
        if (numberOfFields > self.Offsets.TYPE.value):
            entry.type = fields[self.Offsets.TYPE.value]
        if (numberOfFields > self.Offsets.SYSTEM.value):
            entry.system = fields[self.Offsets.SYSTEM.value]
        if (numberOfFields > self.Offsets.DATA.value):
            subFields : typing.List[str] = fields[(self.Offsets.DATA.value):(numberOfFields-1)]
            entry.data = self._SEPARATOR.join(subFields)
            
        return entry
    