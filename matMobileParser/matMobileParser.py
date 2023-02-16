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
    '#FFFFA6', # [0] yellow
    '#FFE994', # [1] yellow-orange
    '#FFB66C', # [2] orange
    '#FFAA95', # [3] coral
    '#FFA6A6', # [4] pinkish
    '#EC9BA4', # [5] mauve
    '#BF819E', # [6] lilac
    '#B7B3CA', # [7] steel blue
    '#B4C7DC', # [8] blue
    '#B3CAC7', # [9] blue-green
    '#AFD095', # [10] green
    '#E8F2A1', # [11] limon (lime-lemon)
)

# === PRIVATE FUNCTIONS ========================================================

def _createSearchQueries() -> SearchQueries:
    search : SearchQueries = SearchQueries()
    search.queries.append(SearchQuery(None, None, 'Display Shot', None, _COLORS[0]))
    search.queries.append(SearchQuery(None, None, 'Lifetime Bow Odometer', None, _COLORS[2]))
    search.queries.append(SearchQuery(None, None, 'Fetching', None, _COLORS[4]))
    search.queries.append(SearchQuery(None, None, 'Shot data transfer completed', None, _COLORS[5]))
    search.queries.append(SearchQuery(None, None, 'Factory Reset', None, _COLORS[6]))
    search.queries.append(SearchQuery(None, None, 'Algorithm run result:', None, _COLORS[8]))
    search.queries.append(SearchQuery(None, None, 'Uploading shot id', None, _COLORS[10]))
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
