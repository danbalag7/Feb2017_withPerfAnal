"""
This file contains various state machines processing state.

"""

from __future__ import print_function
import dateparser
import sys


class Blade:
    column_names = list()
    fault_times = list()

    def __init__(self, df):
        # Stores headers of all columns of Blade file
        self.column_names = df.columns.values

    # TODO(get rid of the df_out parameter to avoid side effects) - What kind of side effects?
    # TODO(populate one set of results per state machine)
    def process_blade_line(self, line, blade_functions, df_out):
        for name, fn in blade_functions.iteritems():
            mutable_df_out = [df_out]
            pline = fn.process(self, line, mutable_df_out)
            df_out = mutable_df_out[0]
            if pline is not None:
                df_out = df_out.append(pline)
        return df_out

    def find_col_name(self, name):
        return next((s for s in self.column_names if name in s), None)

class Inverter:
    def __init__(self, df, timestamps):
        self.column_names = df.columns.values
        self.timestamps = timestamps

    # Here inverter_functions = blade.fault_times
    def process_inverter_line(self, line, inverter_functions, df_out):
        for name, fn in inverter_functions.iteritems():
            pline = fn.process(self, line)
            if pline is not None:
                df_out = df_out.append(pline)
        return df_out

    def find_col_name(self, name):
        return next((s for s in self.column_names if name in s), None)


# All Blade Functions.
class DetectFault:
    good_states = ['Running', 'Starting Ultraverter',
                   'Pre-grid Connect',
                   'Discovery', 'Finalization', 'Gnd Impedance',
                   'Ultrablade Pre-start Check', 'Close Relay']

    bad_states = ['Idle', 'Pending Ready', 'Fault',
                  'Stopping Ultraverter']

    def __init__(self):
        self.history = []
        self.entries = 0
        self.prev_bad_state = None
        self.HISTORY_LEN = 10

    def process(self, blade_instance, line, mut_df_out):
        state_field_name = blade_instance.find_col_name("SGCtrl_State")
        datetime_field_name = blade_instance.find_col_name("DATETIME")
        if state_field_name not in line:
            return None

        self.history.append(line)
        if len(self.history) > self.HISTORY_LEN:
            self.history = self.history[1:]

        if self.entries > 0:
            self.entries -= 1
            return line

        if line[state_field_name] in self.good_states:
            self.prev_bad_state = None
            return None

        if line[state_field_name] in self.bad_states \
                and (self.entries == 0) \
                and line[state_field_name] != self.prev_bad_state :

            # add the previous 10 entries from history
            for loc in self.history:
                mut_df_out[0] = mut_df_out[0].append(loc)

            self.entries = self.HISTORY_LEN
            self.prev_bad_state = line[state_field_name]
            blade_instance.fault_times.append(line[datetime_field_name])
            return line

        return None

## All Blade Functions
class StateChange:
    def __init__(self):
        self.previous_state = None
        self.previous_time = "Start"

    def process(self, blade_instance, line, df_out):
        state_field_name = blade_instance.find_col_name("SGCtrl_State")
        datetime_field_name = blade_instance.find_col_name("DATETIME")
        if state_field_name not in line:
            return None

        if self.previous_state != line[state_field_name]:
            date_transition = "%s -> %s" % (self.previous_time, line[datetime_field_name])
            state_transition = "%s -> %s" % (self.previous_state, line[state_field_name])

            current_state = line[state_field_name]
            current_time = line[datetime_field_name]

            self.change_line(line, "%s -- %s" % (date_transition, state_transition))
            self.previous_state = current_state
            self.previous_time = current_time
            return line
        return None

    def change_line(self, line, contents):
        line.iloc[0] = contents
        for i in range(1, len(line.index)):
            line.iloc[i] = ""


# Filter class for the inverter
class FilterTimes:
    time_range = 11*60
    date_time_format = "%Y/%m/%d %H:%M:%S.%f"

    # line / row_date: corresponds to every timestamp in Inverters.csv
    # inverter_instance.timestamps / date: contains all timestamps at which the Blade had a BAD state change.

    def process(self, inverter_instance, line):
        datetime_field_name = inverter_instance.find_col_name("DATETIME")
        for time in inverter_instance.timestamps:
            date = dateparser.parse(time, date_formats=[self.date_time_format])
            row_date =  dateparser.parse(line[datetime_field_name], date_formats=[self.date_time_format])
            # Prints the previous 2 and next 2 records of the inverter logs at blade_fault_time
            if (date - row_date).seconds < self.time_range or (row_date - date).seconds < self.time_range:
                return line
        return None

