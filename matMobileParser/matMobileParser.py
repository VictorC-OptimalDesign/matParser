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


# === PRIVATE FUNCTIONS ========================================================

def _createSearchQueries() -> SearchQueries:
    search : SearchQueries = SearchQueries()
    search.queries.append(SearchQuery(None, None, 'Display Shot', None))
    search.queries.append(SearchQuery(None, None, 'Lifetime Bow Odometer', None))
    search.queries.append(SearchQuery(None, None, 'Fetching', None))
    search.queries.append(SearchQuery(None, None, 'Shot data transfer completed', None))
    search.queries.append(SearchQuery(None, None, 'Factory Reset', None))
    search.queries.append(SearchQuery(None, None, 'Algorithm run result:', None))
    search.queries.append(SearchQuery(None, None, 'Uploading shot id', None))
    return search

def _process():
    _FILE_SEARCH_PATTERN : str = './**/*.csv'
    print('{0}()'.format(_process.__name__))
    xlsx : XLSX = XLSX()
    search : SearchQueries = _createSearchQueries()
    for fileName in glob.glob(_FILE_SEARCH_PATTERN, recursive=True):
        print('processing {0}...'.format(fileName))
        csv : CSV = CSV(fileName)
        xlsx.writeData(csv, search)
        pass
    xlsx.finalize()


# === MAIN =====================================================================

if __name__ == "__main__":
    print('{0} version {1}'.format(os.path.basename(__file__), _VERSION))
    _process()
else:
    print("ERROR: {0} needs to be the calling python module!".format(os.path.basename(__file__), _VERSION))
