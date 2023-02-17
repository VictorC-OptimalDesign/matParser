#!/usr/bin/env python


# === IMPORTS ==================================================================

import glob
import os
import typing

from input import CSV, SearchQuery, SearchQueries
from output import XLSX


# === CONSTANTS ================================================================

_VERSION_MAJOR : int = 0
_VERSION_MINOR : int = 0
_VERSION_UPDATE : int = 0

_VERSION : str = '{0}.{1}.{2}'.format(_VERSION_MAJOR, _VERSION_MINOR, _VERSION_UPDATE)


_COLORS : typing.Tuple[str] = (
    '#99E4A3', # [0] green (pastel)
    '#F6A09D', # [1] red (pastel)
    '#CFBCFD', # [2] purple (pastel)
    '#DABB9D', # [3] brown (pastel)
    '#F4B1E3', # [4] magenta (pastel)
    '#CFCFCF', # [5] grey (pastel)
    '#FEFDA7', # [6] yellow (pastel)
    '#C1F2EF', # [7] aqua (pastel)
    '#A7CAF2', # [8] blue (pastel)
    '#F7B486', # [9] orange (pastel)
    '#62A76A', # [10] green (deep)
    '#BB5054', # [11] red (deep)
    '#8073B1', # [12] purple (deep)
    '#907861', # [13] brown (deep)
    '#D38DC2', # [14] magenta (deep)
    '#8C8C8C', # [15] grey (deep)
    '#C9B878', # [16] yellow (deep)
    '#72B5CC', # [17] aqua (deep)
    '#5373AE', # [18] blue (deep)
    '#D58457', # [19] orange (deep)
)

# === PRIVATE FUNCTIONS ========================================================

def _createSearchQueries() -> SearchQueries:
    search : SearchQueries = SearchQueries()
    search.queries.append(SearchQuery(None, None, 'Display Shot', None, _COLORS[0]))
    search.queries.append(SearchQuery(None, None, 'Lifetime Bow Odometer', None, _COLORS[1]))
    search.queries.append(SearchQuery(None, None, 'Fetching', None, _COLORS[2]))
    search.queries.append(SearchQuery(None, None, 'Shot data transfer completed', None, _COLORS[3]))
    search.queries.append(SearchQuery(None, None, 'Factory Reset', None, _COLORS[4]))
    search.queries.append(SearchQuery(None, None, 'Algorithm run result:', None, _COLORS[5]))
    search.queries.append(SearchQuery(None, None, 'Uploading shot id', None, _COLORS[6]))
    return search

def _process():
    _FILE_SEARCH_PATTERN : str = './**/*.csv'
    print('{0}()'.format(_process.__name__))
    search : SearchQueries = _createSearchQueries()
    xlsx : XLSX = XLSX(search)
    for fileName in glob.glob(_FILE_SEARCH_PATTERN, recursive=True):
        print('processing {0}...'.format(fileName))
        csv : CSV = CSV(fileName)
        xlsx.writeData(csv)
        pass
    xlsx.finalize()


# === MAIN =====================================================================

if __name__ == "__main__":
    print('{0} version {1}'.format(os.path.basename(__file__), _VERSION))
    _process()
else:
    print("ERROR: {0} needs to be the calling python module!".format(os.path.basename(__file__), _VERSION))
