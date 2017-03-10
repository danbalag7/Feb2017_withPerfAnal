#!/bin/python

"""
Parser for Tygo-generated logs.

This script is designed to work in *nix environment, which can be created
on a Windows machine using Cygwin (https://www.cygwin.com/). The only
requirenmet is that you have Python 2.7+. If *nix-like environment is not
possible, the script should be able to execute on any machine with a standard
Python installation.

In *nix environment don't forget about executable permissions:
$ chmod 755 parser.py

... or you can execute the script like so:
$ python parser.py --help

To view command line options:
$ ./parser.py --help

To parse a CSV log file:
$ ./parser.py log.csv

The script will generate a single CSV file for all inverters and separate
one for each blade.
"""

# pylint: disable=W0311,W0702

from __future__ import print_function
from argparse import ArgumentParser
import os
import re

class Entry(object):
  """ Log entry abstract class """
  def __init__(self, _id, tokens):
    self._id = _id
    self._tokens = tokens
    # The module 're' provides regular expression matching operations
    self._inverter_re = re.compile("^[A-Z]+[0-9]+$")

  def is_blade(self):
    """ Returns True if entry is for a blade """
    # print (self._id)
    return len(self._id) == 1

  def is_inverter(self):
    # print (self._inverter_re)
    """ Returns True if entry is for a inverter """
    return self._inverter_re.match(self._id)

  def is_cloud_connect(self):
    """ Return True if entry is for cloud connect """
    return self._id == "CC"

  def get_id(self):
    """ Returns entry's blade's or inverter's ID """
    # print (self._id)
    return self._id

  def get_field_count(self):
    """ Returns number of fields within this entry """
    # print (len(self._tokens))
    return len(self._tokens)

  def __hash__(self):
    # What does this do????
    return hash(self._id)

  def __eq__(self, other):
    if other is None:
      return False
    return self._id == other.get_id()

  def __str__(self):
    return ",".join(self._tokens)

class Header(Entry):
  """ Log entry for a header """
  def __init__(self, tokens):
    Entry.__init__(self, _id=tokens[3].split("_")[1], tokens=tokens)

class Data(Entry):
  """ Log entry for a data set """

  def __init__(self, tokens, _id):
    Entry.__init__(self, _id=_id, tokens=tokens)

  def get_epoch(self):
    """ Return data entry's epoch-based timestamp """
    return self._tokens[1]

  def get_report_timestamp(self):
    """ Return epoch + datetime combination """
    if self.is_blade():
      # Corresponds to the DATETIME in Blade.csv
      return self._tokens[0]
    else:
      # Corresponds to the Report Timestamp in Inverter.csv
      for i in tokens:
        if "_ReportTimestamp" in self._tokens[i]:
          return self._tokens[i]


class LogStore(object):
  """ Log output abstract class """

  def __init__(self):
    self._headers = {}
    self._logs = {}

  def is_empty(self):
    """ Returns True if log store is empty """
    return True if len(self._headers) == 0 and len(self._logs ) == 0 else False

  def add_entry(self, entry):
    """ Add a log entry to output """

    # entry.get_id() = A or A1, A2 etc.
    log = self._logs.get(entry.get_id(), {})
    if isinstance(entry, Header):
      if entry.get_id() not in self._logs:
        self._headers[entry.get_id()] = entry
        self._logs[entry.get_id()] = log
    else:
      if entry.get_id() not in self._logs:
        print("ERROR: Log is missing.")
      elif entry.get_report_timestamp() in log:
        # in case of duplicate entries keep the first one
        pass
      # elif entry.get_report_timestamp() == "2000-01-01T00:00:00.000Z":
      #   # ignore entries with 2000-01-01T00:00:00.000Z timestamp
      #   pass
      else:
        log[entry.get_report_timestamp()] = entry

    # # Segregating header row of blade / inverter to access each column name if required.
    # header_list = []
    # header_list.append(str(self._headers.get(entry.get_id())))
    # headers_split = [x.split(',') for x in header_list]

    # print (self._headers.keys())
    # print (self._logs.keys())

    # - Prints header of Blade / Inverter
    # - Blade Headers:
    # - [['DATETIME', 'TimeStamp', 'GMT', 'LMU_A_Faults1', 'LMU_A_Faults2', 'LMU_A_SGCtrl_State',
    # 'LMU_A_SGCtrl_StopReason', 'LMU_A_Timestamp', 'LMU_A_WarningFlags', 'LMU_A_Apparent_Power', 'LMU_A_Freq',
    # 'LMU_A_Ia_RMS', 'LMU_A_Reactive_Power', 'LMU_A_Real_Power', 'LMU_A_SGCtrl_State_Int',
    # 'LMU_A_SGCtrl_StopReason_Int', 'LMU_A_Temp', 'LMU_A_Timestamp_Sec', 'LMU_A_Va_RMS', 'LMU_A_Vain_RMS']]
    # - Inverter Headers:
    # - [['DATETIME', 'TimeStamp', 'GMT', 'LMU_A9_Build', 'LMU_A9_Faults', 'LMU_A9_FaultsPersist', 'LMU_A9_PlcBuild',
    # 'LMU_A9_PlcVersion', 'LMU_A9_ReportTimestamp', 'LMU_A9_SafetyVersion', 'LMU_A9_State', 'LMU_A9_StopReason',
    # 'LMU_A9_Timestamp', 'LMU_A9_Version','LMU_A9_WarningFlags', 'LMU_A9_EnergyKwh', 'LMU_A9_EnergyWh',
    # 'LMU_A9_FbMsgRx', 'LMU_A9_FbMsgTx', 'LMU_A9_FxMsgTx', 'LMU_A9_Idc','LMU_A9_IoutQAvg', 'LMU_A9_Pdc',
    # 'LMU_A9_Power', 'LMU_A9_ReportTimestamp_Sec', 'LMU_A9_State_Int', 'LMU_A9_StopReason_Int', 'LMU_A9_Temp',
    # 'LMU_A9_Timestamp_Sec', 'LMU_A9_Vdc', 'LMU_A9_VoutQAvg']]

    # - Checking which column is used in get_report_timestamp():
    # if entry.is_blade():
    #     blade_report_header = headers_split[0][0]
    # if entry.is_inverter():
    #     inv_report_header = headers_split[0][17]

  def dump_to_file(self):
    """ Output-log-to-file abstract method """
    pass

def tryint(string):
  """ return integer if s can be converted to one, otherwise return s itself """
  try:
    return int(string)
  except:
    return string

def natural_key(string):
  """ tokenize string key into list of strings and ints """
  return [tryint(c) for c in re.split("([0-9]+)", string)]

class InverterLogStore(LogStore):
  """ Inverter log output """

  def dump_to_file(self, dir="."):
    """ Print all inverter logs to one file """

    if self.is_empty():
      return

    _file = open(os.path.join(dir, "Inverters.csv"), "w")
    # _file_timestamps = open(os.path.join(dir, "Timestamps.csv"), "w")


    # - print headers
    # - collect timestamps for all entries for sorting
    # - as a 1st previous entry use a blank entry
    timestamps = []
    ids = sorted(self._headers.keys(), key=natural_key) # sort inverter IDs
    ids = filter(None, ids)                             # remove empty strings
    prev_entries = {}
    # print (ids)
    for _id in ids:
      print(self._headers[_id], end=",", file=_file)
      # print ("self._logs[_id].values()", self._logs[_id].values())
      # print ("self._logs[_id].keys()", self._logs[_id].keys())

      timestamps += self._logs[_id].keys()
      # - prev_entries produces 30 commas, repeated in 'n' rows; n = # inverters
      prev_entries[_id] = "," * (self._headers[_id].get_field_count() - 1)
    print("", file=_file)   # newline

    # remove duplicate timestamps:
    # print ("unsorted", len(timestamps))
    # print (timestamps, file = _file_timestamps)
    timestamps = list(set(timestamps))
    # iterate over sorted timestamps for *all* entries
    timestamps.sort()
    # print ("SORTED TS", len(timestamps))
    # print (timestamps, file = _file_timestamps)


    for time in timestamps:
      # - print entries for all IDs that are "within" this time
      # - print previous entry if ID doesn't have one for this time
      for _id in ids:
        log_entry = self._logs[_id].get(time, prev_entries[_id])
        prev_entries[_id] = log_entry
        print(log_entry, end=",", file=_file)
      print("", file=_file) # newline

    _file.close()

class BladeLogStore(LogStore):
  """ Blade log output """

  def dump_to_file(self, dir="."):
    """ Print blade logs to individual files """

    if self.is_empty():
      return

    # open files and write headers for each blade
    fds = {}
    for _id, header in self._headers.iteritems():
      path = os.path.join(dir, "Blade_" + header.get_id() + ".csv")
      fds[_id] = open(path, "w")
      print(header, file=fds[_id])

    for _id, log in self._logs.iteritems():
      for timestamp in sorted(log):
        print(log[timestamp], file=fds[_id])

    # close files
    for _id, _file in fds.iteritems():
      _file.close()

def header_tokens(_tokens):
  """ Returns True if tokens are for a header """
  return _tokens[0] == "DATETIME"

def parse(filePath, outdir="."):

  # print (filePath)
  """ Main """
  with open(filePath) as _file:
    content = _file.readlines()

  invlog = InverterLogStore()
  bladelog = BladeLogStore()
  header = None

  for line in content:
    tokens = line.rstrip().split(",")

    if header_tokens(tokens):
      header = Header(tokens)
      if header.is_blade():
        bladelog.add_entry(header)
      elif header.is_inverter():
        invlog.add_entry(header)

    elif header is not None:
      # NOTE: assume last header pertains to all consecutive data entries
      data = Data(tokens, header.get_id())
      if data.is_blade():
        bladelog.add_entry(data)
      elif data.is_inverter():
        invlog.add_entry(data)

  invlog.dump_to_file(outdir)
  bladelog.dump_to_file(outdir)

if __name__ == "__main__":
    arg_parser = ArgumentParser(description="Parser for Tygo-generated logs.")
    arg_parser.add_argument("log_path", nargs=1, help="log file to parse")
    args = arg_parser.parse_args()
    parse(args.log_path[0])


# simply commenting