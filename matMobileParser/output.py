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
    
    
    _HEADERS : typing.Tuple[str] = (
        'line',
        'time',
        'type',
        'system',
        'data',
        'extra'
    )
    
    
    def __init__(self):
        self.size : int = 0
        self.wb : Workbook = Workbook(self._FILE_NAME + self._EXTENSION)
        self.logSheets : typing.List[Workbook.worksheet_class] = []
        
        
    def writeData(self, csv : CSV, search : SearchQueries = None):
        ws : Workbook.worksheet_class = self.wb.add_worksheet(csv.name)
        
        # Write the header.
        row : int = 0
        for col, header in enumerate(self._HEADERS):
            ws.write(row, col, header)
        row += 1
        
        # Write the data.
        for entry in csv.entries:
            if (self._searchEntry(entry, search)):
                ws.write(row, 0, entry.line)
                ws.write(row, 1, entry.time)
                if (entry.type != None):
                    ws.write(row, 2, entry.type)
                if (entry.system != None):
                    ws.write(row, 3, entry.system)
                if (entry.data != None):
                    ws.write(row, 4, entry.data)
                extraLength = len(entry.extra)
                if (extraLength > 0):
                    col = 5
                    for extraEntry in entry.extra:
                        ws.write(row, col, extraEntry)
                        col += 1
                row += 1
        self.logSheets.append(ws)
        
        
    def _searchEntry(self, entry : LogEntry, search : SearchQueries) -> bool:
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
                    break
            pass
        return found
    
    
    def finalize(self):
        self.wb.close()
