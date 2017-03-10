import os, glob, csv, win32com.client
from xlsxwriter.workbook import Workbook
import matplotlib.pyplot as plt
import matplotlib.pyplot as text
import matplotlib.dates as mdate
import datetime as dt
import xlwings as xw
from matplotlib.backends.backend_pdf import PdfPages
import FileDialog
from os.path import basename
import time
from xlwings.constants import Direction
import numpy as np


def convert_files(outdir):
    # function to convert all .csv files into .xlsx files
    pathname = os.path.join(os.getcwd(), outdir)
    for csvfile in glob.glob(os.path.join(pathname, '*.csv')):
        workbook = Workbook(csvfile + '.xlsx')
        worksheet = workbook.add_worksheet()

        # rb refers to mode of opening a file - r = reading, b = appending to a binary file
        with open(csvfile, 'rb') as f:
            reader = csv.reader(f)
            for r, row in enumerate(reader):
                for c, col in enumerate(row):
                    worksheet.write(r, c, col)

        print ("fileWasconverted")
        workbook.close()

def rename_files(outdir):
    # get file names from a folder
    pathname = os.path.join(os.getcwd(), outdir)
    file_list = os.listdir(pathname)

    # for each file, rename filename
    for file_name in file_list:
        # only if file name contains .xlsx will it be renamed
        # all original .csv files remain unchanged
        if "csv.xlsx" in file_name:
            new_file_name = file_name.replace(".csv", "")
            file_name = os.path.join(pathname, file_name)
            new_file_name = os.path.join(pathname, new_file_name)
            os.rename(file_name, new_file_name)

    finding_paths(outdir)

def finding_paths(outdir):
    # This function for all excel files to be correctly named before accessing for macro run
    # get file names from a folder
    pathname = os.path.join(os.getcwd(), outdir)
    file_list = os.listdir(pathname)
    # file_list = os.listdir(os.curdir)
    # saved_path = os.getcwd()
    for file_name in file_list:
        if ".xlsx" in file_name:
            if file_name.startswith("Inverters"):
                # inv_path = saved_path + "\\" + file_name
                inv_path = os.path.join(pathname, file_name)
                run_macro(inv_path)
                file_name = os.path.splitext(os.path.basename(inv_path))[0] # To remove the file extension
                inv_charting(inv_path,file_name, outdir)
            if file_name.startswith("Blade"):
                # blade_path = saved_path + "\\" + file_name
                blade_path = os.path.join(pathname, file_name)
                run_macro(blade_path)
                file_name = os.path.splitext(os.path.basename(blade_path))[0] # To remove the file extension
                blade_charting(blade_path, file_name, outdir)

def run_macro(path_name):
    print "entering run_macro"
    # print path_name

    x1 = win32com.client.Dispatch("Excel.Application")
    x1.Visible = False
    xlsPath = os.path.expanduser(path_name)
    xlsMacro = os.path.expanduser(r'C:\Users\divyaa\AppData\Roaming\Microsoft\AddIns\forall.xlam')

    wb =x1.Workbooks.Open(Filename=xlsPath)
    x1.Workbooks.Open(Filename=xlsMacro)

    x1.Application.Run('forall.xlam!Module1.autoformat')
    print('ran macro')

    x1.DisplayAlerts = True
    wb.Close(True)
    x1.Application.Quit()

def blade_charting(path_name, file_name, outdir):
    # start_blade = time.time()
    #----------------------------------------------------------------------------------------------------------------------
    # Interacting with Excel using Python's xlwings
    #----------------------------------------------------------------------------------------------------------------------
    # Opening an existing workbook
    print "entering blade chart"
    # print path_name
    app = xw.App(visible=False)
    wb = app.books.open(path_name)
    ws = wb.sheets[0]
    folder_name = os.path.join(os.getcwd(), outdir)
    last_row = 0
    last_col = 0

    # Finding the last non-zero row and last non-zero column
    # Max possible rows in a single XL sheet
    last_row = ws.cells(1048576, 1).end(Direction.xlUp).row
    max_params_per_inv = 40
    max_inverters_per_blade = 15
    max_blades = 3
    last_col = ws.cells(1, (max_params_per_inv * max_inverters_per_blade * max_blades)).end(Direction.xlToLeft).column

    # List to capture the 'day'
    day = []
    day = [ws.cells(i, 1).options(dates=dt.date).value for i in range(2, last_row + 1)]
    end_blade = time.time()
    # print day

    # This list is not actually required since the required parameter is the list of first and last rows.
    # list_of_days = []
    list_of_firstlastrows = []
    # Appending cell(2,1) as the first row, after ignoring the headers in Row 1
    list_of_firstlastrows.append(2)

    # Range starts at 0 for the first element in the list
    # Range ends at (last_row-3) instead of (last_row-1) in order to ignore header & the last_row
    # (which is written manually)
    for i in range(0, last_row - 3):
        if day[i] != day[i + 1]:
            # list_of_days.append(day[i])
            # Becomes the last row of the first day
            list_of_firstlastrows.append(i + 2)
            # Becomes the first row of the next day
            list_of_firstlastrows.append(i + 3)

    # list_of_days.append(day[last_row - 2])
    # Appending the last_row of the Excel sheet to the last element in the list
    list_of_firstlastrows.append(last_row)

    # Finding the Blade Letter
    # Using the 4th header which contains LMU_A/B/C
    blade_letter = ws.cells(1, 4).value[4:5]

    # Opening a PDF document to save plots
    pp = PdfPages(folder_name + "\\" + file_name + '.pdf')

    # start_cols = time.time()
    list_reqd_cols = []
    list_reqd_cols = [col for col in range(1, last_col + 1) if "_SGCtrl_State_Int" in ws.cells(1, col).value or
                      "_Real_Power" in ws.cells(1, col).value]
    # end_cols = time.time()
    # print ("cols time: ", (end_cols - start_cols))

    # Assigning first and last rows to each day of data
    # start_mainfor = time.time()
    for day_loop in range (0, (len(list_of_firstlastrows)/2)):
        fig = plt.figure(figsize=(12,7))
        first_row = list_of_firstlastrows[2*day_loop]
        last_row = list_of_firstlastrows[2*day_loop + 1]

        # date string to datetime object
        dates_array_lim = []
        row_dates_lim = []
        dates_array = ws.range((first_row, 1), (last_row, 1)).options(dates=dt.datetime).value
        # ------------------------------------------------------------------------------------------
        # NEED TO WORK ON SETTING LIMITS TO DATE ROWS:
        # earliest_time = dt.time(5, 0, 0)
        # latest_time = dt.time(21, 0, 0)
        # start_date = time.time()
        # print first_row
        # print last_row
        # dates_array_lim = [dates_array[dts]
        #                    for dts in range(first_row-2, last_row - 1)
        #                    if dates_array[dts].time() > earliest_time and
        #                    dates_array[dts].time() < latest_time]
        # row_dates_lim = [r_dts for r_dts in range(first_row-2, last_row - 1)
        #                  if dates_array[r_dts].time() > earliest_time and
        #                  dates_array[r_dts].time() < latest_time]
        # print row_dates_lim
        # end_date = time.time()
        # print ("date limit: ", (end_date - start_date))

        # Create formatted string for given time/date/datetime object according to specified "dateonly" format.
        dateonly = '%b-%d-%Y'
        newdatetitle = dt.datetime.strftime(dates_array[0], dateonly)

        # Setting x-tick format string
        date_fmt = '%H:%M'

        for k in list_reqd_cols:
            if ws.cells(1,k).value == 'LMU_' + blade_letter+ '_SGCtrl_State_Int':
                # -----------------------------------------------------------------------------------------
                # Blade State Curve
                # -----------------------------------------------------------------------------------------
                # Setting all x-axis attributes
                # Tip: If only date to be displayed, (dt.dt.strptime(date, xxx)).
                # Used to convert dates into datetime.datetime format.
                ax1 = fig.add_subplot(211)
                ax1.grid(b=True)

                # # Setting x-tick format string
                # date_fmt = '%H:%M'
                # Use a DateFormatter to set the data to the correct format.
                date_formatter = mdate.DateFormatter(date_fmt)
                ax1.xaxis.set_major_formatter(date_formatter)
                # Setting x-axis range
                ax1.set_xlim([dt.datetime(dates_array[0].year, dates_array[0].month, dates_array[0].day, 5, 0, 0),
                              dt.datetime(dates_array[0].year, dates_array[0].month, dates_array[0].day, 21, 0, 0)])

                # Specifying x-tick labels.
                for tick in ax1.xaxis.get_major_ticks():
                    tick.label.set_fontsize(9)
                # --------------------------------------------------------------------------------------------------------
                # Setting all y-axis attributes
                # --------------------------------------------------------------------------------------------------------
                blade_states = ws.range((first_row,k),(last_row,k)).value
                # Setting y-axis range
                min_blade_states = 0
                max_blade_states = 17
                ax1.set_ylim(min_blade_states-1, max_blade_states+1)
                ax1.set_ylabel("Blade State", fontsize=11, fontstyle='italic', weight='bold')
                # Specifying y-tick labels.
                ax1.set_yticks((0,1,2,3,4,5,8,13,16,17))
                ax1.set_yticklabels(('Idle', 'PendingReady', 'CloseRel', 'Starting', 'Running',
                                     'Fault','Stopping', 'Gnd Impedance', 'Pre-start', 'Pre-grid',),
                                    fontsize = 8, fontstyle = 'italic')
                for tick in ax1.yaxis.get_major_ticks():
                    tick.label.set_fontsize(10)
                # --------------------------------------------------------------------------------------------------------
                # Plotting Blade State Curve
                # --------------------------------------------------------------------------------------------------------
                # Plotting as a step function
                ax1.step(dates_array, blade_states)
                ax1.set_title('LMU_' + blade_letter+ '_SGCtrl_State',
                              fontsize = 12, fontstyle = 'italic', weight ='bold')
            # ----------------------------------------------------------------------
            # String Power Curve
            # ----------------------------------------------------------------------
            elif ws.cells(1,k).value=='LMU_' + blade_letter+ '_Real_Power':
                # --------------------------------------------------------------------------------------------------------
                # Setting all x-axis attributes
                # --------------------------------------------------------------------------------------------------------

                ax2 = fig.add_subplot(212)
                ax2.grid(b = True)
                # Setting x-tick format string
                # date_fmt = '%H:%M'
                # Use a DateFormatter to set the data to the correct format.
                date_formatter = mdate.DateFormatter(date_fmt)
                ax2.xaxis.set_major_formatter(date_formatter)
                # Setting x-axis range
                ax2.set_xlim([dt.datetime(dates_array[0].year, dates_array[0].month, dates_array[0].day, 5, 0, 0),
                              dt.datetime(dates_array[0].year, dates_array[0].month, dates_array[0].day, 21, 0, 0)])

                # Specifying x-tick labels.
                for tick in ax2.xaxis.get_major_ticks():
                    tick.label.set_fontsize(9)

                ax2.set_ylabel("String Power (W)", fontsize = 11, fontstyle = 'italic', weight ='bold')
                # ----------------------------------------------------------------------------------------------------------------------
                # Setting all y-axis attributes
                # ----------------------------------------------------------------------------------------------------------------------
                # Setting y-axis range
                string_power = ws.range((first_row, k), (last_row, k)).value

                # Specifying y-tick labels.
                for tick in ax2.yaxis.get_major_ticks():
                    tick.label.set_fontsize(10)
                # ----------------------------------------------------------------------------------------------------------------------
                # Plotting String Power Curve
                # ----------------------------------------------------------------------------------------------------------------------
                ax2.plot(dates_array, string_power, color = 'red')
                ax2.set_title(ws.cells(1,k).value, fontsize = 12, fontstyle = 'italic', weight ='bold')

        plt.subplots_adjust(left=None, bottom=None, right=None, top=None, wspace=None, hspace=None)
        plt.suptitle("BLADE "+ blade_letter + " : " + newdatetitle, fontsize = 16, fontstyle='italic', weight='bold')
        pp.savefig()
        plt.close()
    end_mainfor = time.time()
    # print ("mainfor dayloop time: ", (end_mainfor - start_mainfor))

    pp.close()
    wb.close()
    # end_blade = time.time()
    # print ("Bladechart time: ", (end_blade-start_blade))
    # plt.show()

def inv_charting(path_name, file_name, outdir): #outdir

    # ----------------------------------------------------------------------------------------------------------------------
    # Interacting with Excel using Python's xlwings
    # ----------------------------------------------------------------------------------------------------------------------
    print "entering inv chart"
    # start_inv = time.time()
    folder_name = os.path.join(os.getcwd(), outdir)

    # Opening an existing workbook
    app = xw.App(visible=False)
    wb = app.books.open(path_name)
    ws = wb.sheets[0]

    # Calculation of Last rows and columns
    last_row = 0
    last_col = 0

    last_row = ws.cells(1048576, 1).end(Direction.xlUp).row
    # nasty bug - doesn't detect hidden columns to the left. Fix: Macro unhides VoutQ (last col)
    max_params_per_inv = 40
    max_inverters_per_blade = 15
    max_blades = 3
    last_col = ws.cells(1, (max_params_per_inv * max_inverters_per_blade * max_blades)).end(Direction.xlToLeft).column
    # print last_col
    # end_inv = time.time()
    # print ("finding lastrowcol: ", end_inv - start_inv)

    # # ----------------------------- Finding # inverters on each blade, and last col of each blade
    lastinv_bladeA = 0
    lastinv_bladeB = 0
    lastinv_bladeC = 0
    first_col_A = 0
    first_col_B = 0
    first_col_C = 0
    last_col_A= 0
    last_col_B = 0
    last_col_C = 0
    k = 0
    countA = 0
    countB = 0
    countC = 0

    # --------------------- Method 1
    start_inv = time.time()
    # for k in range(1,last_col+1):
    #     if "LMU_A" in ws.cells(1,k).value:
    #         countA= k
    #     elif "LMU_B" in ws.cells(1, k).value:
    #         countB = k
    #     elif "LMU_C" in ws.cells(1, k).value:
    #         countC = k
    # --------------------- Method 2
    # for k in range(1,last_col+1):
    #     if "LMU_A" in ws.cells(1,k).value:
    #         countA = k
    # end_inv = time.time()
    # for k in range(countA, last_col + 1):
    #     if "LMU_B" in ws.cells(1, k).value:
    #         countB = k
    #         print countB
    # for k in range(countB, last_col + 1):
    #     if "LMU_C" in ws.cells(1, k).value:
    #         countC = k
    #         print countC
    # # end_inv = time.time()
    # print ("finding lastrowcol: ", end_inv - start_inv)

    # --------------------- Method 3
    # start_list = time.time()
    list_last_cols_A = [k for k in range(1,last_col+1) if "LMU_A" in ws.cells(1,k).value]
    countA = list_last_cols_A[len(list_last_cols_A)-1]

    # Assigning first and last columns to each blade's inverters
    if countA <>0:
        lastinv_bladeA = ws.cells(1, countA).value[5:ws.cells(1, countA).value.rfind('_')]  # r is not reverse, it means right-most
        first_col_A = 1
        last_col_A = countA
        list_last_cols_B = [k for k in range(countA, last_col + 1) if "LMU_B" in ws.cells(1, k).value]
        if list_last_cols_B:
            countB = list_last_cols_B[len(list_last_cols_B) - 1]
    if countB <> 0:
        lastinv_bladeB = ws.cells(1, countB).value[5:ws.cells(1, countB).value.rfind('_')]
        first_col_B = last_col_A+1
        last_col_B = countB
        list_last_cols_C = [k for k in range(countB, last_col + 1) if "LMU_C" in ws.cells(1, k).value]
        if list_last_cols_C:
            countC = list_last_cols_C[len(list_last_cols_C) - 1]
    if countC <> 0:
        lastinv_bladeC = ws.cells(1, countC).value[5:ws.cells(1, countC).value.rfind('_')]
        first_col_C = last_col_B+1
        last_col_C = countC

    # print list_last_cols_A[len(list_last_cols_A)-1]
    # print list_last_cols_B[len(list_last_cols_B)-1]
    # print list_last_cols_C[len(list_last_cols_C)-1]
    # end_list = time.time()
    # print ("finding list, if: ", end_list - start_list)

    day = []
    for i in range(2, last_row + 1):
        days = ws.cells(i, 1).options(dates=dt.date).value
        day.append(days)

    list_of_days = []
    list_of_firstlastrows = []
    list_of_firstlastrows.append(2)
    for i in range(0, last_row - 3):
        if day[i] != day[i + 1]:
            # print "entered if"
            list_of_days.append(day[i])
            list_of_firstlastrows.append(i + 2)
            list_of_firstlastrows.append(i + 3)

    list_of_days.append(day[last_row - 2])
    list_of_firstlastrows.append(last_row)
    # end_days = time.time()
    # print ("days time: ", (end_days - start_days))

    # end_inv = time.time()
    # print ("Invchart time: ", (end_inv - start_inv))

    pp = PdfPages(folder_name + "\\" + file_name + '.pdf')
    start_for = time.time()
    for dayloop in range(0, len(list_of_days)):
        if lastinv_bladeA:
            lastinv_bladeA = int(lastinv_bladeA)
            # end_inv = time.time()
            # print ("till before curve plotting is called: ", end_inv-start_inv)
            curve_plotting(lastinv_bladeA, 'A', list_of_firstlastrows[2*dayloop], list_of_firstlastrows[2*dayloop + 1],
                           first_col_A, last_col_A, ws,pp)
        if lastinv_bladeB:
            lastinv_bladeB = int(lastinv_bladeB)
            curve_plotting(lastinv_bladeB, 'B', list_of_firstlastrows[2*dayloop], list_of_firstlastrows[2*dayloop + 1],
                           first_col_B, last_col_B,ws,pp)
        if lastinv_bladeC:
            lastinv_bladeC = int(lastinv_bladeC)
            curve_plotting(lastinv_bladeC, 'C', list_of_firstlastrows[2*dayloop], list_of_firstlastrows[2*dayloop + 1],
                       first_col_C, last_col_C,ws,pp)
    end_for = time.time()
    print "total for time ", (end_for-start_for)

    # plt.show()
    pp.close()
    wb.close() # Be mindful of lower case 'close'

def curve_plotting(lastinv_blade, bladeLetter, first_row, last_row, first_col, last_col, ws, pp):
    # start_curve = time.time()
    inputpower_list = []
    list_reqd_cols = []
    list_reqd_cols = [col for col in range(first_col, last_col + 1)
                      if "_State_Int" in ws.cells(1, col).value
                      or "_Power" in ws.cells(1, col).value
                      or "_IoutQAvg" in ws.cells(1, col).value
                      or "_VoutQAvg" in ws.cells(1, col).value]

    out_pow = []
    out_pow_range = []
    iout = 0
    vout = 0
    i = 0
    dts = 0
    # print last_col
    # print last_row

    # date string to datetime object
    dates_array_lim =[]
    row_dates_lim =[]
    dates_array = ws.range((first_row, first_col), (last_row, first_col)).options(dates=dt.datetime).value

    # Date title for all saved images
    dateonly = '%b-%d-%Y'
    # print dates_array[0]
    newdatetitle = dt.datetime.strftime(dates_array[0], dateonly)

    # Setting x-tick format string
    date_fmt = '%H:%M'
    # Use a DateFormatter to set the data to the correct format.
    date_formatter = mdate.DateFormatter(date_fmt)

    for n in range(1, lastinv_blade + 1):
        print bladeLetter+`n`
        # start_curve = time.time()
        fig1 = plt.figure(figsize=(12, 6))  # (width) x (height)
        plt.suptitle("LMU_" + bladeLetter + `n` + " " + newdatetitle,
                     fontsize=18, fontstyle='italic', weight='bold')

        # ----------------------------------------------------------------------
        # Inverter AC Power Array
        # ----------------------------------------------------------------------
        out_pow = [(float(ws.cells(i, iout).value) * float(ws.cells(i, vout).value) / 2)
                   for vout in list_reqd_cols if "LMU_" + bladeLetter + `n` + "_VoutQAvg" in ws.cells(1, vout).value
                   for iout in list_reqd_cols if "LMU_" + bladeLetter + `n` + "_IoutQAvg" in ws.cells(1, iout).value
                   for i in range(first_row, last_row + 1)]
        out_pow_range.append(out_pow)

        for k in list_reqd_cols:

            start_for2 = time.time()
            # ----------------------------------------------------------------------------------------------------------------------
            # Inverter State Curve
            # ----------------------------------------------------------------------------------------------------------------------
            if ("LMU_" + bladeLetter + `n` + "_State_Int") in ws.cells(1, k).value:
                # start_invstate = time.time()
                # print ws.cells(1, k).value
                # ----------------------------------------------------------------------------------------------------------------------
                # Setting all x-axis attributes
                # ----------------------------------------------------------------------------------------------------------------------
                # Splitting the chart into figure and axes for easy access as separate objects
                ax1 = fig1.add_subplot(121)
                ax1.grid(b=True)

                ax1.xaxis.set_major_formatter(date_formatter)
                # Setting x-axis range
                ax1.set_xlim([dt.datetime(dates_array[0].year, dates_array[0].month, dates_array[0].day, 5, 0, 0),
                              dt.datetime(dates_array[0].year, dates_array[0].month, dates_array[0].day, 21, 0, 0)])

                # Changing font size of x-axis ticks
                for tick in ax1.xaxis.get_major_ticks():
                    tick.label.set_fontsize(10)
                    tick.label.set_rotation(45)
                    tick.label.set_ha('right')

                # ------------------------------------------------------------------
                # Setting all y-axis attributes
                # ------------------------------------------------------------------
                inv_states = ws.range((first_row, k), (last_row, k)).value

                # Setting y-axis range
                min_inv_states = 0
                max_inv_states = 10
                ax1.set_ylim(min_inv_states - 1, max_inv_states + 1)

                # Specifying y-tick labels.
                ax1.set_yticks((0, 1, 2, 3, 4, 5, 6, 10))
                ax1.set_yticklabels(('Idle', '???', 'VPan_Wait', 'Ready', 'Boost_St', 'Running', 'Fault','Gnd_Imp'),
                                    size='small')

                # ------------------------------------------------------------------
                # Plotting Inverter State Curve
                # ------------------------------------------------------------------

                # Plotting as a step function
                ax1.step(dates_array, inv_states, color='red')
                ax1.set_title('Inverter State', fontsize=12, fontstyle='italic', weight='bold')
                ax1.set_ylabel("Machine States", fontsize=12, fontstyle='italic', weight='bold')
                # end_invstate = time.time()
                # print ("invstate " +`n` , (end_invstate-start_invstate))
            # ----------------------------------------------------------------------
            # Inverter DC Power Curve
            # ----------------------------------------------------------------------
            if ("LMU_" + bladeLetter + `n` + "_Power") in ws.cells(1, k).value:
                start_invpow = time.time()
                # ------------------------------------------------------------------
                # Setting all x-axis attributes
                # ------------------------------------------------------------------
                # Setting x-tick format string
                ax2 = fig1.add_subplot(122)
                ax2.grid(b=True)

                # # Use a DateFormatter to set the data to the correct format.
                # date_formatter = mdate.DateFormatter(date_fmt)
                ax2.xaxis.set_major_formatter(date_formatter)

                # Setting x-axis range
                ax2.set_xlim([dt.datetime(dates_array[0].year, dates_array[0].month, dates_array[0].day, 5, 0, 0),
                              dt.datetime(dates_array[0].year, dates_array[0].month, dates_array[0].day, 21, 0, 0)])

                # Changing font size of x-axis ticks
                for tick in ax2.xaxis.get_major_ticks():
                    tick.label.set_fontsize(10)
                    tick.label.set_rotation(45)
                    tick.label.set_ha('right')

                # ------------------------------------------------------------------
                # Setting all y-axis attributes
                # ------------------------------------------------------------------
                dc_power = ws.range((first_row, k), (last_row, k)).value
                # dcpower_list.append(dc_power)
                #
                # Setting y-axis range
                max_power_inv = 300
                ax2.set_ylim(0, max_power_inv)

                # Specifying y-tick labels.
                for tick in ax2.yaxis.get_major_ticks():
                    tick.label.set_fontsize(10)
                # ------------------------------------------------------------------
                # Plotting String Power Curve
                # ------------------------------------------------------------------
                ax2.plot(dates_array, dc_power, color='blue', label= 'Panel Power')
                ax2.plot(dates_array, out_pow_range[n - 1], color='green', label= 'Output Power')
                ax2.set_title('Output & Panel Power', fontsize=12, fontstyle='italic', weight='bold')
                ax2.set_ylabel("Power (W)", fontsize=12, fontstyle='italic', weight='bold')
                ax2.legend(loc='upper right', fontsize='small')
                end_invpow = time.time()
                # print ("invpow ", `n`, (end_invpow - start_invpow))

        fig1.autofmt_xdate()
        pp.savefig()
        plt.close()
        # end_curve = time.time()
        # print "curve time ", (end_curve - start_curve)
    # ------------------------------------------------------------------------------------
    # Plotting All-in-One AC Power Curve
    # ------------------------------------------------------------------------------------
    fig2 = plt.figure(figsize=(12, 6))
    plt.suptitle("Blade " + bladeLetter + " Inverter AC Power Curves : " + newdatetitle,
                 fontsize=18, fontstyle='italic', weight='bold')
    ax = fig2.add_subplot(111)

    ax.xaxis.set_major_formatter(date_formatter)

    # Setting x-axis range
    ax.set_xlim([dt.datetime(dates_array[0].year, dates_array[0].month, dates_array[0].day, 5, 0, 0),
                 dt.datetime(dates_array[0].year, dates_array[0].month, dates_array[0].day, 21, 0, 0)])

    # Changing font size of x-axis ticks
    for tick in ax.xaxis.get_major_ticks():
        tick.label.set_fontsize(10)
    fig2.autofmt_xdate()

    # Setting y-axis range
    max_power_panel = 300
    ax.set_ylim(0, max_power_panel)

    # Specifying y-tick labels.
    for tick in ax.yaxis.get_major_ticks():
        tick.label.set_fontsize(10)

    ax.set_ylabel("Power (AC W)", fontsize=12, fontstyle='italic', weight='bold')

    ax.grid(b=True)
    # print lastinv_blade
    for n in range(1, lastinv_blade + 1):
        ax.plot(dates_array, out_pow_range[n - 1], label="Inv " + bladeLetter + `n`)
    plt.legend(loc='upper right', fontsize='small')
    pp.savefig()
    plt.close()
    # end_curve = time.time()
    # print "curve time ", (end_curve-start_curve)


if __name__ == '__main__':
    pass
    # start_in = time.time()
    # eff_choice = raw_input("Should I plot the efficiency curves of inverters? (y/Y): ")
    # convert_files()
    # rename_files()
    # end_in = time.time()
    # print (end_in - start_in)
    # run_macro(r'C:\Users\divyaa\Desktop\PowerCurves\thomassolar\FullRun\Inverters_faller_edit.xlsx')
    # blade_charting(r'C:\Users\divyaa\Desktop\PowerCurves\thomassolar\FullRun\Blade_B_TS.xlsx', 'Blade_B_TS.xlsx')
    # start_in = time.time()
    # inv_charting(r'C:\Users\divyaa\Desktop\Fault_Perf_Analysis\r\Feb2017_withPerfAnal\28831\Inverters.xlsx', 'Inverters.xlsx', r'28831')
    # end_in = time.time()
    # print (end_in - start_in)