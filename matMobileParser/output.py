#!/usr/bin/env python


# === IMPORTS ==================================================================

import typing

from input import CSV, LogEntry, SearchQuery, SearchQueries
from xlsxwriter import Workbook

 
# === CLASSES ==================================================================

class XLSX:
    _FILE_NAME : str = 'matMobile'
    _EXTENSION : str = '.xlsx'
    _SHEET_NAME_LOG : str = 'log'
    _SHEET_NAME_STATS : str = 'stats'
    _DATE_NUMBER_FORMAT : str = 'yyyy-mm-dd'
    _TIME_NUMBER_FORMAT : str = 'hh:mm:ss.000000'
    _TOP_ALIGN_FORMAT : str = 'top'
    
    
    _HEADERS : typing.Tuple[str] = (
        'line',
        'date',
        'time',
        'search',
        'type',
        'system',
        'data',
        'extra'
    )
    
    
    def __init__(self):
        self.size : int = 0
        self.wb : Workbook = Workbook(self._FILE_NAME + self._EXTENSION)
        self.logSheets : typing.List[Workbook.worksheet_class] = []
        self.dateFormat = self.wb.add_format({'num_format' : self._DATE_NUMBER_FORMAT})
        self.timeFormat = self.wb.add_format({'num_format' : self._TIME_NUMBER_FORMAT})
        self.dataFormat = self.wb.add_format()
        self.extraFormat = self.wb.add_format()
        self.dataFormat.set_align(self._TOP_ALIGN_FORMAT)
        self.extraFormat.set_align(self._TOP_ALIGN_FORMAT)
        
        
    def writeData(self, csv : CSV, search : SearchQueries = None):
        ws : Workbook.worksheet_class = self.wb.add_worksheet(csv.name)
        
        # Write the header.
        row : int = 0
        for col, header in enumerate(self._HEADERS):
            ws.write(row, col, header)
        row += 1
        
        # Write the data.
        for entry in csv.entries:
            matchingQuery : SearchQuery = self._findMatchingSearchQuery(entry, search)
            if (matchingQuery != None):
                ws.write(row, self._HEADERS.index('line'), entry.line)
                ws.write(row, self._HEADERS.index('date'), entry.time, self.dateFormat)
                ws.write(row, self._HEADERS.index('time'), entry.time, self.timeFormat)
                ws.write(row, self._HEADERS.index('search'), str(matchingQuery))
                if (entry.type != None):
                    ws.write(row, self._HEADERS.index('type'), entry.type)
                if (entry.system != None):
                    ws.write(row, self._HEADERS.index('system'), entry.system)
                if (entry.data != None):
                    ws.write(row, self._HEADERS.index('data'), entry.data, self.dataFormat)
                extraLength = len(entry.extra)
                if (extraLength > 0):
                    col = self._HEADERS.index('extra')
                    for extraEntry in entry.extra:
                        ws.write(row, col, extraEntry, self.extraFormat)
                        col += 1
                row += 1
                
        ws.autofit()
        ws.set_column(self._HEADERS.index('time'), self._HEADERS.index('time'), len(self._TIME_NUMBER_FORMAT))
        self.logSheets.append(ws)
        
        
    def _findMatchingSearchQuery(self, entry : LogEntry, search : SearchQueries) -> SearchQuery:
        matchingQuery : SearchQuery = None
        found : bool = True
        if (search != None and len(search.queries) > 0):
            for query in search.queries:
                found = True
                if (query.type != None):
                    found = found and (entry.type != None) and (entry.type.find(query.type) >= 0)
                if (query.system != None):
                    found = found and (entry.system != None) and (entry.system.find(query.system) >= 0)
                if (query.data != None):
                    found = found and (entry.data != None) and (entry.data.find(query.data) >= 0)
                if (query.extra != None):
                    foundExtra : bool = False
                    for extraEntry in entry.extra:
                        foundExtra = extraEntry.find(query.extra)
                        if (foundExtra):
                            break
                    found = found and foundExtra
                if found:
                    matchingQuery = query
                    break
            pass
        return matchingQuery
    
    
    def finalize(self):
        self.wb.close()
