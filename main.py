#!/usr/bin/env python

from fetch import Fetch
import logging
import json
import argparse
import tygoLogParser
import perfAnalysis
from faults import find_faults

# Accesses tables in website and downloads raw csv files
fetcher = Fetch()
# Accesses the "argparse" module: creating an ArgumentParser object, which holds all info necessary to parse CL into Python data types
parser = argparse.ArgumentParser()


def main(args):
    logging.basicConfig(level=logging.DEBUG, filename='web2read.log')
    if args.list:
        list_systems()
        exit(0)

    if args.id is not None:
        if args.start is None or args.end is None:
            # we want to list valid date ranges.
            list_dates(args.id)
            exit(0)

    # we want to download the CSV files associated.
    if args.id is not None:
        csvs = download_csv(args.id, args.start, args.end)
        if args.parse:
            parse_logs(csvs, args.id)

        if args.parse and args.faults:
            # Passing the system ID to function find_faults in faults.py
            find_faults(str(args.id))

        exit(0)

    parser.print_help()


# Calls Fetch
def download_csv(id, start, end):
    return fetcher.get_csvs(id, start, end)


# Calls tygoLogParser to create Blade_A.csv, Inverters.csv etc.
def parse_logs(csvs, system_id):
    ignored_csv_file = 'panel'
    for csv in csvs:
        if ignored_csv_file in csv:
            continue
        print("\nParsing CSV file: %s\n" % csv)
        tygoLogParser.parse(csv, str(system_id))

    # # Calling performance analysis files to plot curves
    # perfAnalysis.convert_files(str(system_id))
    # perfAnalysis.rename_files(str(system_id))


# Lists online and offline systems
def list_systems():
    online, offline = fetcher.get_systems()

    print("Online Systems")
    _print_systems(online)

    print("Offline Systems")
    _print_systems(offline)


def list_dates(system_id):
    dates = fetcher.get_dates(system_id)
    print(json.dumps(dates, indent=2))


def _print_systems(systems_dict):
    print(json.dumps(systems_dict, indent=2))


if __name__ == "__main__":
    # Filling an ArgumentParser with information about program arguments is done by making calls to the add_argument() method.
    # Generally, these calls tell the ArgumentParser how to take the strings on the command line and turn them into objects.
    # This information is stored and used when parse_args() is called.
    parser.add_argument('-l', '--list', dest='list', action='store_true', default=False)
    parser.add_argument('-i', '--id', dest='id', type=int)
    parser.add_argument('-s', '--start', dest='start')
    parser.add_argument('-e', '--end', dest='end')
    parser.add_argument('-p', '--parse', dest='parse', action='store_true', default=False)
    parser.add_argument('-f', '--faults', dest='faults', action='store_true', default=False)

    # args looks like: Namespace(end=None, faults=False, id=None, list=True, parse=False, start=None)
    args = parser.parse_args()

    list_systems()
    system_id = raw_input('Enter System ID: ')
    args.id = system_id
    list_dates(system_id)

    print "Select from available dates above: "
    start_date = raw_input('Start Date yyyy-mm-dd: ')
    args.start = start_date

    end_date = raw_input('End date (not included) yyyy-mm-dd: ')
    args.end = end_date

    args.parse = True
    args.faults = False
    main(args)
