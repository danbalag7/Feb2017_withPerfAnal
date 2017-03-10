"""
Find the faults in a given file.

"""
import argparse
import pandas as pd
import glob
import os
from faultsUtil import StateChange, Blade, DetectFault, FilterTimes, Inverter
from collections import OrderedDict
import perfAnalysis


parser = argparse.ArgumentParser()
blade_columns = ['DATETIME', 'SGCtrl_State', 'Faults1',
                  'SGCtrl_StopReason', 'Real_Power',
                  'Va_RMS', 'Temp', 'WarningFlags', 'Faults2']

fault_times = list()

def process_blade(file):
    df = pd.read_csv(file)
    blade_functions = OrderedDict()
    blade_functions['detect_fault'] = DetectFault()
    blade_functions['state_change'] = StateChange()

    # List that tracks the history of the
    print(blade_functions['detect_fault'].history)
    df_out = pd.DataFrame(data=None, columns=df.columns)
    blade = Blade(df)

    for index, row in df.iterrows():
        df_out = blade.process_blade_line(row, blade_functions, df_out)
    return (df_out, blade)

def blade_drop_columns(blade, df_out):
    columns_needed = map(blade.find_col_name, blade_columns)
    column_names = df_out.columns.values
    for col in column_names:
        if col not in columns_needed:
            del df_out[col]

def last_columns():
    pass
def process_inverter(file, timestamps, blade_letter):
    df = pd.read_csv(file)
    inverter_functions = {
        'filter_times': FilterTimes(),
    }
    # columns_per_blade = [col for col in df.columns if "LMU_"+blade_letter in col]
    # first_col = 0
    # last_col = 0
    # if blade_letter == 'A':
    #     first_col = 0
    #     last_col = perfAnalysis.last_col_A - 1
    # elif blade_letter == 'B':
    #     first_col = perfAnalysis.last_col_A + 1
    #     last_col = perfAnalysis.last_col_B - 1
    # elif blade_letter == 'C':
    #     first_col = perfAnalysis.last_col_B + 1
    #     last_col = perfAnalysis.last_col_C - 1
    # df_out = pd.DataFrame(data=None, columns=df.columns[first_col:last_col])
    df_out = pd.DataFrame(data=None, columns=df.columns)

    inverter = Inverter(df, timestamps)
    for index, row in df.iterrows():
        df_out = inverter.process_inverter_line(row, inverter_functions, df_out)
    return (df_out, inverter)

def inverter_drop_columns(inverter, df_out):
    columns_needed = map(inverter.find_col_name, inverter_columns)
    print columns_needed
    column_names = df_out.columns.values
    for col in column_names:
        if col not in columns_needed:
            del df_out[col]

def find_faults(dir):
    # The Blades data looks like: Blade_A.csv, Blade_B.csv...
    # The Inverter data looks like: Inverters.csv

    blade_files = glob.glob(os.path.join(dir, "Blade*.csv"))
    inverter_files = glob.glob(os.path.join(dir, "Inverter*.csv"))

    for file in blade_files:
        print("Checking faults in: ", file)

        # TODO: fixme (add to Blade class)
        df_blade, blade = process_blade(file)
        blade_drop_columns(blade, df_blade)

        basename = os.path.splitext(os.path.basename(file))[0]
        dirname = os.path.dirname(file)
        processed_file_name = "processed-%s" % basename

        if '_A' in basename:
            blade_letter = 'A'
        elif '_B' in basename:
            blade_letter = 'B'
        elif '_C' in basename:
            blade_letter = 'C'
        # append relevant inverter data - the +- 2 records for each blade fault timestamp
        # assert: Statement that makes the program tests this condition and triggers an error if condition is false.
        assert(len(inverter_files) == 1)
        # inverter_files[0] = '27994\Inverters.csv'
        df_inverter, inverter = process_inverter(inverter_files[0], blade.fault_times, blade_letter)

        # merge here with df_blade
        combined = df_blade.join(df_inverter, lsuffix='_df_blade', rsuffix='_df_inverter', how='outer')
        combined.to_csv("%s.csv" % os.path.join(dirname, processed_file_name))

if __name__ == "__main__":
    parser.add_argument('-d', '--dir', dest='dir')
    args = parser.parse_args()
    find_faults(args.dir)