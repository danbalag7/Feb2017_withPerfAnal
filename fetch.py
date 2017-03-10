"""
Class that fetches CSV files

"""

from webClient import WebClient
from util import DateRange
import dateparser
import logging
import zipfile
import os

class Fetch:
    date_time_format = '%Y-%m-%d'
    username = "empower"
    password =  "emgo123"

    wc = WebClient(username, password)
    output_categories = {
        "com_kaco_app_blade_details": "blade",
        "com_kaco_app_details": "app_details",
        "com_kaco_app_panels": "app_panel"
    }

    def __init__(self):
        pass

    def get_systems(self):
        online, offline = self.wc.get_systems()
        return (online, offline)

    def get_dates(self, system_id):
        dts = self.wc.get_valid_dates(system_id)
        date_ranges = list()
        date_ranges_str = list()

        # this is a basic state machine to turn a list of dates into a range.
        for dt in dts:
            date = dateparser.parse(dt[0], date_formats=[self.date_time_format])
            if len(date_ranges) == 0:
                date_ranges.append(DateRange(date, date))
            else:
                if (date - date_ranges[-1].end).days == 1:
                    date_ranges[-1].end = date
                else:
                    date_ranges.append(DateRange(date, date))
        for range in date_ranges:
            date_ranges_str.append("%s - %s" % (range.start.strftime(self.date_time_format),
                                                range.end.strftime(self.date_time_format)))
        return date_ranges_str


    def get_csvs(self, system_id, start, end):
        csvs = []

        i = 1
        for category, name in self.output_categories.iteritems():
            print("\nFetching CSV %d of %d" % (i, len(self.output_categories)))
            i += 1

            filename = "%s.zip" % name
            self.wc.get_csv(system_id, category, start, end, filename)
            # self._test_zip(filename)
            csvs.append(self._zip_extract(system_id, filename, name))
        return csvs

    # this function is not being called because of errors with zlib.
    def _test_zip(self, path):
        """ Tests the zip file to make sure it contains useful data """
        logging.debug("Testing: %s" % path)
        with open(path, 'r') as test:
            zip_file = zipfile.ZipFile(test)
            file = zip_file.namelist()[0]
            ret = zip_file.read(file)
            if 'Invalid Parameters' in ret:
                raise ValueError('No useful data exists')

    def _zip_extract(self, system_id, path, name):
        csvName = ""
        with open(path, 'rb') as zf:
            z = zipfile.ZipFile(zf)
            csvName = z.namelist()[0]
            z.extract(csvName, os.path.join(str(system_id), name))
        os.remove(path)
        return os.path.join(str(system_id), name, csvName)