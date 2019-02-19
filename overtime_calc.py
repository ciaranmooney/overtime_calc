#! /usr/bin/env python
# A script to take timesheets from my employer and calculate a summary of the 
# hours worked.
# Ciarán Mooney
# 2019

import datetime
import os
from xlrd import open_workbook

class weekHours(object):
    ''' Holder for hours worked.
    '''

    def __init__(self, days, weekNo):
        self.days = days
        self.weekNumber = weekNo

        self.std = 0
        self.time = 0
        self.timehalf = 0
        self.double = 0
        
        for day in self.days:
            self.std = self.std + day.hours[0]
            self.time = self.time + day.hours[1]
            self.timehalf = self.timehalf + day.hours[2]
            self.double = self.double + day.hours[3]

    def hoursSummary(self):
        ''' Outputs the standard, time, time and half, and double time totals for
            the whole week.
        '''
        
        return (self.std, self.time, self.timehalf, self.double)

class hoursWorked(object):
    ''' Class to hold a date and hours tuple with hours work.
    '''

    def __init__(self, date, hours):
        '''
        '''
        self.date = date
        self.hours = hours

def excel_date_conv(excel_date):
    ''' Date from Excel is an integer for the number of days since
        1st January 1900.
    '''
    epoch = datetime.date(1900,1,1)
    return epoch + datetime.timedelta(days=excel_date - 2) # -2 due to bug in Excel

def sum_column(cells):
    ''' Sums the list of columns
    '''
    total = 0 
    for hour in cells:
        if hour.value != '':
            total = total + float(hour.value)
    return total

def parse_hours(cell_ref, workbook, weekend=False):
    ''' cell_ref is tuple of (row, startcol, endcol)
    '''
    if weekend == False:
        std = workbook.col_slice(cell_ref[0],cell_ref[1],cell_ref[2])
        std_total = sum_column(std)
        
        time = workbook.col_slice(cell_ref[0]+1,cell_ref[1],cell_ref[2])
        time_total = sum_column(time)
            
        timehalf = workbook.col_slice(cell_ref[0]+2,cell_ref[1],cell_ref[2])
        timehalf_total = sum_column(timehalf)
        
        timedouble = workbook.col_slice(cell_ref[0]+3,cell_ref[1],cell_ref[2])
        timedouble_total = sum_column(timedouble)
            
        return (std_total, time_total, timehalf_total, timedouble_total)

    else:
        timehalf = workbook.col_slice(cell_ref[0],cell_ref[1],cell_ref[2])
        timehalf_total = sum_column(timehalf)
        
        timedouble = workbook.col_slice(cell_ref[0]+1,cell_ref[1],cell_ref[2])
        timedouble_total = sum_column(timedouble)
        return (0.0, 0.0, timehalf_total, timedouble_total)
    
def excel_date_parse(excel_file):
    ''' Parses an excel timesheet file to produce the list of days and hours
        worked.

        Returns a week object.
    '''
    
    with open_workbook(excel_file, 'rb') as wb:
        timeSheet = wb.sheet_by_index(0)

    timeSheetDate = excel_date_conv(timeSheet.cell(2,26).value)
    weekNumber = timeSheetDate.isocalendar()[1]
    
    tues_date = timeSheetDate + datetime.timedelta(days=1)
    wed_date = timeSheetDate + datetime.timedelta(days=2)
    thurs_date = timeSheetDate + datetime.timedelta(days=3)
    fri_date = timeSheetDate + datetime.timedelta(days=4)
    sat_date = timeSheetDate + datetime.timedelta(days=5)
    sun_date = timeSheetDate + datetime.timedelta(days=6)
    mon_date = timeSheetDate + datetime.timedelta(days=7)
    
    tuesday_hours = parse_hours((3,6,31), timeSheet)
    wednesday_hours = parse_hours((7,6,31), timeSheet)
    thursday_hours = parse_hours((11,6,31), timeSheet)
    friday_hours = parse_hours((15,6,31), timeSheet)
    saturday_hours = parse_hours((19,6,31), timeSheet)
    sunday_hours = parse_hours((23,6,31), timeSheet, True)
    monday_hours = parse_hours((25,6,31), timeSheet, True)
    
    tuesday =  hoursWorked(tues_date, tuesday_hours)
    wednesday =  hoursWorked(wed_date, wednesday_hours)
    thursday =  hoursWorked(thurs_date, thursday_hours)
    friday =  hoursWorked(fri_date, friday_hours)
    saturday =  hoursWorked(sat_date, saturday_hours)
    sunday =  hoursWorked(sun_date, sunday_hours)
    monday =  hoursWorked(mon_date, monday_hours)

    week = weekHours((tuesday, wednesday, thursday, friday, saturday, sunday, monday), weekNumber)

    return week

def sumWeeks(weeks, start, stop):
    ''' Takes a list of week objects, and a start week and stop week. Sums the
        hours to give a summary.
    '''

    std_total = 0
    time_total = 0
    timehalf_total = 0
    double_total = 0
    
    for week in weeks:
        if start <= week.weekNumber <= stop:
            std_total = std_total + week.std
            time_total = time_total + week.time
            timehalf_total = timehalf_total + week.timehalf
            double_total = double_total + week.double

    return (std_total, time_total, timehalf_total, double_total)

if __name__ == '__main__':
    month = 9 # 1, january, 2, february ...
    month_names = ['January', 'February', 'March', 'April','May','June','July','August','September','October','November','December']
    timesheet_directory = # add path for directory.

    weeks = []
    for timesheet in os.listdir(timesheet_directory):
        if timesheet.split('.')[1] == 'xls':
            path = timesheet_directory + '\\' + timesheet
            weeks.append(excel_date_parse(path))

    startWeekNum = int(input('Enter start week number:\n'))
    endWeekNum = int(input('Enter end week number:\n'))
        
    print(sumWeeks(weeks, startWeekNum, endWeekNum))
