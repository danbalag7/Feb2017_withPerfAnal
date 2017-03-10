#!/usr/bin/env python
import urllib

import logging
import mechanize
from BeautifulSoup import BeautifulSoup
from collections import OrderedDict
import json
import sys

class WebClient:
    login_url = 'https://installations.tigoenergy.com/base/login/login'
    systems_list_table_url = 'https://installations.tigoenergy.com/main/ajax/sys_table_data.php'
    systems_valid_dates_url = 'https://installations.tigoenergy.com/base/main/summary/energy'
    power_data_url = 'https://installations.tigoenergy.com/base/main/details/powerdata'
    serial_number_url = 'https://installations.tigoenergy.com/base/main/status/uv/sysid/'
    raw_data_freq = -1
    zero_time = '00-00-00'

    browser = None

    def __init__(self, login, password):
        """ Logs into tigoenergy website using empower credentials"""
        self.browser = mechanize.Browser()
        self.browser.open(self.login_url)
        self.browser.select_form(nr=0)
        self.browser.form['Users[login]'] = login
        self.browser.form['Users[password]'] = password
        self.browser.submit()

    def logout(self):
        self.browser.open('https://installations.tigoenergy.com/base/login/logout')
        self.browser.close()

    def get_csv(self, system_id, src, start, end, file_name):
        get_data = {
            'sysid': system_id,
            'from': "%s_%s" % (start, self.zero_time),
            'to': "%s_%s" % (end, self.zero_time),
            'freq': self.raw_data_freq,
            'src': src,
            'uid': self._get_serial_number(system_id),
            'power': 1,
            'lmus': 1,
            'z': 1,
        }

        # https://installations.tigoenergy.com/base/main/details/powerdata?rfl=1&sysid=27369
        # &from=2016-09-07_00-00-00
        # &to=2016-09-09_00-00-00
        # &uid=04C05B8075DD
        # &src=com_kaco_app_blade_details
        # &power=1
        # &lmus=1
        # &z=1
        # &freq=-1

        try:
            logging.debug("Fetching: %s?%s" % (self.power_data_url, urllib.urlencode(get_data)))
            self.browser.retrieve("%s?%s" % (self.power_data_url, urllib.urlencode(get_data)),
                                  file_name, reporthook=self._download_progress)[0]

        except Exception as err:
            logging.error(err)

    # var count is the current amount downloaded.
    # var blockSize is the maximum size of the bytes being transferred.
    # var totalSize is the total size of the file.
    def _download_progress(self, count, blockSize, totalSize):
        """ Prints progress of download """
        percent = int(count * blockSize * 100 / totalSize)  # total percentage of work done
        if percent > 100:
            percent = 100
        sys.stdout.write('\r%2d%%' % percent)
        sys.stdout.flush()

    def get_systems(self):
        """ Returns a list of systems (online, offline) """


        # Ordered Dict to hold systems. (SystemID -> Address)
        # parsed row looks like: [[], [u'28150'], [], [u'3,000 W'], [u'Test77D9'],
        # [u'3020 Kenneth St, Santa Clara, CA 95051, United States'], [], []]
        online_systems = OrderedDict()
        offline_systems = OrderedDict()
        table_addr_col_index = 5
        table_sysid_col_index = 1
        table_name_col_index = 4


        post_data = {
            'limit': '50',
            'orderby': 'id',
            'filter': 'a:0:{}',
            'asc': '1',
            'status_access': '2'
        }

        response = self.browser.open(self.systems_list_table_url, data=urllib.urlencode(post_data))
        soup = BeautifulSoup(response.read())
        rows = (soup.findChildren('table')[0]).findChildren(['tr'])
        for row in rows:
            # parse the row into columns (a list)
            row_parsed = [col.findAll(text=True) for col in row.findAll('td')]

            # remove entries within columns which just have a '\n' in them.
            # parsed row looks like: [[], [u'28150'], [], [u'3,000 W'], [u'Test77D9'],
            # [u'3020 Kenneth St, Santa Clara, CA 95051, United States'], [], []]
            row_parsed = [filter(lambda x: x != '\n', r) for r in row_parsed]
            logging.debug("Parsed row: %s", row_parsed)
            if len(row_parsed) < table_sysid_col_index: continue
            sys_id = row_parsed[table_sysid_col_index][0]
            addr = row_parsed[table_addr_col_index][0]
            sys_name = row_parsed[table_name_col_index][0]
            logging.debug("Parsed system ID: %s, Address: %s, SysName: %s", sys_id, addr, sys_name)

            # online_systems is a dictionary, so it is being referred to by a key
            if self._is_system_online(row):
                # online_systems[sys_id] = addr
                online_systems[sys_id] = sys_name
            else:
                # offline_systems[sys_id] = addr
                offline_systems[sys_id] = sys_name

        return (online_systems, offline_systems)

    def get_valid_dates(self, system_id):
        get_data = {
            'sysid': system_id,
        }

        response = self.browser.open("%s?%s" % (self.systems_valid_dates_url, urllib.urlencode(get_data)))
        json_dates = str(BeautifulSoup(response.read()))
        return json.loads(json_dates)

    def _get_serial_number(self, system_id):
        serial_number_table_index = 5
        serial_number_title_str = 'Serial Number'

        """
        The table looks like:

        [[u'Serial Number'], [u'04C05B8075DD']]
        [[u'Timezone'], [u'\n', u'America/Los_Angeles', u'\n']]
        [[u'Last check-in'], [u'2017-Jan-21 18:57:28']]
        [[u'Last update'], [u' 00:01:07 ago']]
        [[u'Software Build'], [u'2.8.3-nd-bin']]

        It can be found in (for e.g.)
        https://installations.tigoenergy.com/base/main/status/uv/sysid/27369
        """

        url = self.serial_number_url + str(system_id)
        response = self.browser.open(url)
        soup = BeautifulSoup(response.read())
        table = soup.findAll('table')[serial_number_table_index]
        for row in table.findChildren('tr'):
            row_parsed = [col.findAll(text=True) for col in row.findAll('td')]
            row_parsed = filter(lambda x: x != [], row_parsed)
            if len(row_parsed) > 1 and (row_parsed[0] == [serial_number_title_str]):
                logging.debug("Found serial number: %s" % row_parsed[0])
                return row_parsed[1][0]
        raise Exception


    # if we find the image html 'ok.png', the system is online
    def _is_system_online(self, row):
        encodedRow = unicode(row).encode('utf-8')
        return True if ('ok.png' in encodedRow) else False