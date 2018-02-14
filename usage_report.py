"""usage_report.py
As of now, this is complete and self-contained! (Other than installed modules)
Close Excel before running.
"""
import pyodbc
import pandas as pd
import os.path
import datetime as dt
from datetime import timedelta as td
import win32com.client as win32

def setup_FTP():
    server = '***REMOVED***'
    database = ***REMOVED***
    username =***REMOVED***
    password = '***REMOVED***'
    cnxn = pyodbc.connect('DRIVER=***REMOVED***;SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
    return cnxn

def setup_dest_file(name = "ERROR no name provided.xlsx", loc = os.path.join(os.getcwd(),"Usage Reports")):
    path = os.path.join(loc, name)
    if not os.path.exists(loc):
        os.makedirs(loc)
    writer = pd.ExcelWriter(path, 'xlsxwriter', datetime_format = 'mm/dd/yyyy hh:mm:ss', date_format='mm/dd/yyyy')
    return writer


#returns the most recent Friday BEFIRE today, returns last friday if run on this friday
def last_friday(d):
    if d.weekday() == 4:
        return d-td(days = 7)
    return d - td(days=d.weekday() + 3)

#Returns today (if it is Friday) or the next Friday (if it is not)
def this_friday(d):
    if d.weekday() == 4:
        return d
    if d.weekday() < 4:
        return d + td(days= 4 - d.weekday())
    return d + td(days= 11 - d.weekday())

def friday_last_year(a):
    b = a.replace(year=a.year - 1)
    return this_friday(b)

def setup_dates():
    today = dt.date.today()
    #Current Week:
    cw_start = last_friday(today)
    cw_end = this_friday(today)
    #Last Week:
    lw_end = cw_start
    lw_start = last_friday(lw_end)
    #Last Year
    ly_start = friday_last_year(cw_start)
    ly_end = friday_last_year(cw_end)

    dates = {"This Week":(cw_start, cw_end), "Last Week":(lw_start, lw_end), "Last Year":(ly_start, ly_end)}
    return dates

def fill_dates(df, week):
    idx = pd.date_range(week[0],week[1])
    df.index = pd.DatetimeIndex(df.index)
    df = df.reindex(idx, fill_value = 0)

    return df

def sql_week(s, week):
    return """Declare @weekstart datetime, @weekend datetime
set @weekstart='"""+str(week[0])+"""'
set @weekend='"""+str(week[1])+"'" + s


def main():
    #getting connection object for FTP queries
    connection = setup_FTP()

    #getting week start and end dates for report based on today's date
    dates = setup_dates()

    #Setting up Excel file and using the date of the end of this week to name it
    name = str(dates["This Week"][1]) + " Weekly Usage Report.xlsx"
    writer = setup_dest_file(name)

    #grabbing reference to the workbook and adding some formats we'll use later when writing
    workbook  = writer.book
    f_week = workbook.add_format({
        'bold': True,
        'font_color':'white',
        'bg_color':'#4f81bd'
        })
    f_purple = workbook.add_format({
        'bold': True,
        'font_color':'white',
        'bg_color':'#8064a2'
        })
    f_header = workbook.add_format({
        'bold': True,
        'border': 0
        })
    f_percent = workbook.add_format({'num_format': '0.00%'})
    f_header_percent = workbook.add_format({
        'bold': True,
        'border':0,
        'num_format': '0%'
        })
    f_date = workbook.add_format({
        'border':0
        })

    #setting up some variables for keeping track of which week we're talking about and placing the data properly on each sheet
    weekname = ["This Week","Last Week","Last Year"]
    weekcount, C1, C2, C3, C4, C5, C6 = 0, 0, 0, 0, 0, 0, 0
    R1A = 1
    R1B = 17
    R2 = 1
    R3 = 1
    R4A = 2
    R4B = 14
    R5 = 2
    R6A = 2
    R6B = 14

    #Some pre-written chunks of excel formula
    weekly_change = ["Weekly Change", "=(B10-G10)/G10", "=(C10-H10)/H10", "=(D10-I10)/I10"]
    yearly_change = ["Yearly Change", "=(B10-L10)/L10", "=(C10-M10)/M10", "=(D10-N10)/N10"]



    #Some bits of SQL code, saved to variables to keep our code cleaner
    sql_1A ="""---- # of Test Results by Date
select CONVERT(varchar(10), tr.UpdatedDate,101) as Date, COUNT(tr.testresultid) as Total, sum(case when tr.qtionlinetestsessionid is not null then 1 else 0 end) as OnlineTests, sum(case when tr.BubbleSheetID is not null then 1 else 0 end) as BubbleSheets from TestResult tr
join VirtualTest vt on vt.VirtualTestID=tr.VirtualTestID
join Student s on s.StudentID=tr.StudentID
join District d on d.DistrictID=s.DistrictID
join State st on st.StateID=d.StateID
where tr.UpdatedDate>@weekstart and tr.UpdatedDate<@weekend and d.Name not like '%demo%' and (tr.BubbleSheetID is not null or tr.QTIOnlineTestSessionID is not null)
group by CONVERT(varchar(10), tr.UpdatedDate, 101)
order by CONVERT(varchar(10), tr.UpdatedDate, 101)"""
    sql_1B ="""---- # of Test Results by Date sans A Beka, BEC, A List and CEE
select CONVERT(varchar(10), tr.UpdatedDate,101) as Date, COUNT(tr.testresultid) as Total, sum(case when tr.qtionlinetestsessionid is not null then 1 else 0 end) as OnlineTests, sum(case when tr.BubbleSheetID is not null then 1 else 0 end) as BubbleSheets from TestResult tr
join VirtualTest vt on vt.VirtualTestID=tr.VirtualTestID
join Student s on s.StudentID=tr.StudentID
join District d on d.DistrictID=s.DistrictID
join State st on st.StateID=d.StateID
where tr.UpdatedDate>@weekstart and tr.UpdatedDate<@weekend and d.Name not like '%demo%' and (tr.BubbleSheetID is not null or tr.QTIOnlineTestSessionID is not null) and d.DistrictID not in (2680, 2479) and d.DistrictGroupID not in (112,114) and d.name not like '%frog street%'
group by CONVERT(varchar(10), tr.UpdatedDate, 101)
order by CONVERT(varchar(10), tr.UpdatedDate, 101)"""
    sql_2 ="""---- # of Test Results by Client
select st.Name as State, d.Name as District, count(1) TotalResults,sum(case when tr.qtionlinetestsessionid is not null then 1 else 0 end) as OnlineTests, sum(case when tr.BubbleSheetID is not null then 1 else 0 end) as BubbleSheets from TestResult tr
join VirtualTest vt on vt.VirtualTestID=tr.VirtualTestID
join Student s on s.StudentID=tr.StudentID
join District d on d.DistrictID=s.DistrictID
join State st on st.StateID=d.StateID
where tr.UpdatedDate>@weekstart and tr.UpdatedDate<@weekend and d.Name not like '%demo%' and (tr.BubbleSheetID is not null or tr.QTIOnlineTestSessionID is not null)
group by st.Name, d.Name
order by count(1) desc"""
    sql_3 ="""---- # of LinkIt Benchmarks by Client
select st.Name as State, d.Name as District, count(1) TotalResults,sum(case when tr.qtionlinetestsessionid is not null then 1 else 0 end) as OnlineTests, sum(case when tr.BubbleSheetID is not null then 1 else 0 end) as BubbleSheets from TestResult tr
join VirtualTest vt on vt.VirtualTestID=tr.VirtualTestID
join Student s on s.StudentID=tr.StudentID
join District d on d.DistrictID=s.DistrictID
join State st on st.StateID=d.StateID
where tr.UpdatedDate>@weekstart and tr.UpdatedDate<@weekend and d.Name not like '%demo%' and (tr.BubbleSheetID is not null or tr.QTIOnlineTestSessionID is not null) and vt.Name like '%linkit%form%'
group by st.Name, d.Name
order by count(1) desc"""
    sql_4A ="""---- # of Online Test Sessions by Start Time ---
select CONVERT(varchar(10), qots.startdate,101) as [Date Started], count(1) as [Total # of Online Tests], sum(case when qots.statusid=1 then 1 else 0 end) as [# of Created], sum(case when qots.statusid=2 then 1 else 0 end) as [# of Started], sum(case when qots.statusid=3 then 1 else 0 end) as [# of Paused], sum(case when qots.statusid=5 then 1 else 0 end) as [# of Pending Review], sum(case when qots.statusid=4 then 1 else 0 end) as [# of Completed] from QTIOnlineTestSession qots With (nolock)
join student s With (nolock) on s.studentid=qots.studentid
join district d With (nolock) on d.DistrictID=s.districtid
where d.name not like '%demo%' and qots.StartDate>@weekstart and qots.StartDate<@weekend
group by CONVERT(varchar(10), qots.startdate,101)
order by  CONVERT(varchar(10), qots.startdate,101)"""
    sql_4B ="""---- # of Online Test Sessions by Last Log In Time ---
select CONVERT(varchar(10), qots.LastLoginDate,101) as [Date Last Log In], count(1) as [Total # of Online Tests], sum(case when qots.statusid=1 then 1 else 0 end) as [# of Created], sum(case when qots.statusid=2 then 1 else 0 end) as [# of Started], sum(case when qots.statusid=3 then 1 else 0 end) as [# of Paused], sum(case when qots.statusid=5 then 1 else 0 end) as [# of Pending Review], sum(case when qots.statusid=4 then 1 else 0 end) as [# of Completed] from QTIOnlineTestSession qots With (nolock)
join student s With (nolock) on s.studentid=qots.studentid
join district d With (nolock) on d.DistrictID=s.districtid
where d.name not like '%demo%' and qots.LastLoginDate>@weekstart and qots.LastLoginDate<@weekend
group by CONVERT(varchar(10), qots.LastLoginDate,101)
order by  CONVERT(varchar(10), qots.LastLoginDate,101)"""
    sql_5 ="""---- # of Online Test Sessions by Hour by Last Log In Time ---
select CONVERT(varchar(13), dateadd(hour, -4,qots.LastLoginDate),120) as [Hour], count(1) as [Number of Sessions], SUM(case when d.DistrictID=2479 then 1 else 0 end) as [A Beka], sum(case when d.districtgroupid=112 then 1 else 0 end) as BEC, sum(case when d.name like '%frog street%' then 1 else 0 end) as [Frogstreet] from QTIOnlineTestSession qots With (nolock)
join student s With (nolock) on s.studentid=qots.studentid
join district d With (nolock) on d.DistrictID=s.districtid
where d.name not like '%demo%' and qots.LastLoginDate>@weekstart and qots.LastLoginDate<@weekend
group by CONVERT(varchar(13), dateadd(hour, -4,qots.LastLoginDate),120)
order by  count(1) desc"""
    sql_6A ="""---- # of Results Entry by Date
select CONVERT(varchar(10), tr.UpdatedDate,101) as Date, COUNT(tr.testresultid) as Total from TestResult tr
join VirtualTest vt on vt.VirtualTestID=tr.VirtualTestID
join Student s on s.StudentID=tr.StudentID
join District d on d.DistrictID=s.DistrictID
join State st on st.StateID=d.StateID
where tr.UpdatedDate>@weekstart and tr.UpdatedDate<@weekend and d.Name not like '%demo%' and vt.virtualtestsourceid=3 and vt.virtualtesttype in (1,5)
group by CONVERT(varchar(10), tr.UpdatedDate, 101)
order by CONVERT(varchar(10), tr.UpdatedDate, 101)"""
    sql_6B ="""---- # of Results Entry by District
select st.Name as State, d.Name as District, COUNT(tr.testresultid) as Total from TestResult tr
join VirtualTest vt on vt.VirtualTestID=tr.VirtualTestID
join Student s on s.StudentID=tr.StudentID
join District d on d.DistrictID=s.DistrictID
join State st on st.StateID=d.StateID
where tr.UpdatedDate>@weekstart and tr.UpdatedDate<@weekend and d.Name not like '%demo%' and vt.virtualtestsourceid=3 and vt.virtualtesttype in (1,5)
group by st.Name, d.Name
order by COUNT(tr.testresultid) desc"""


    for week in dates.values():
        #Part 1A
        sql = sql_1A
        R = R1A
        C = C1
        N = "# of Results by Date"


        #Get data from database
        df = pd.read_sql(sql_week(sql, week), connection)


        #Write the data to the file
        df.to_excel(writer, sheet_name = N, index = False, header=False, startrow = R+1, startcol = C)
        worksheet = writer.sheets[N]

        #Then formatted headers (to_excel uses an ugly format you can't overwrite)
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(R, C+col_num, value, f_header)


        #Add "This Week", "Last Week", or "Last Year" above table as appropriate
        worksheet.write_string(R-1, C, weekname[weekcount], f_week)

        #Add 'Total' row at bottom of table
        worksheet.write_string(R+8, C, 'Total', f_header)
        worksheet.write(R+8, C+1, df['Total'].sum(), f_header)
        worksheet.write(R+8, C+2, df['OnlineTests'].sum(), f_header)
        worksheet.write(R+8, C+3, df['BubbleSheets'].sum(), f_header)

        #Add weekly/yearly change
        if weekcount is 0:
            for i in range(4):
                worksheet.write(R+10, C+i, weekly_change[i], f_header_percent)
                worksheet.write(R+11, C+i, yearly_change[i], f_header_percent)


        dfa = df
        #Part 1B
        sql = sql_1B
        R = R1B
        #Get data from database
        df = pd.read_sql(sql_week(sql, week), connection)
        #Write the data to the file, then formatted headers
        df.to_excel(writer, sheet_name = N, index = False, header=False, startrow = R+1, startcol = C)
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(R, C+col_num, value, f_header)


        #Add "This Week", "Last Week", or "Last Year" above table as appropriate
        worksheet.write_string(R-1, C, weekname[weekcount], f_week)

        #Add 'Total' row at bottom of table
        worksheet.write_string(R+8, C, 'Total', f_header)
        worksheet.write(R+8, C+1, df['Total'].sum(), f_header)
        worksheet.write(R+8, C+2, df['OnlineTests'].sum(), f_header)
        worksheet.write(R+8, C+3, df['BubbleSheets'].sum(), f_header)

        #Add weekly/yearly change
        if weekcount is 0:
            for i in range(4):
                worksheet.write(R+10, C+i, weekly_change[i].replace("10","26"), f_header_percent)
                worksheet.write(R+11, C+i, yearly_change[i].replace("10","26"), f_header_percent)

        if weekcount is 2:
            #We want this written but not counted for column widths, so we do it at the end
            worksheet.write(R1B-2, 0, "Without BEC, A Beka, A List, CEE, Frog Street", f_purple)
            worksheet.write(R1B-2, 1, "", f_purple)
            worksheet.write(R1B-2, 2, "", f_purple)
            worksheet.write(R1B-2, 3, "", f_purple)
        C1 = C1 + 5

        #Part 2
        sql = sql_2
        R = R2
        C = C2
        N = "# of Results by Client"
        #worksheet = writer.sheets[N]

        #Get data from database
        df = pd.read_sql(sql_week(sql, week), connection)
        #Add '%' column by summing the TotalResults column and dividing each entry by that total
        tr_sum = df.TotalResults.sum(axis=0)
        df['%'] = (df['TotalResults']/tr_sum)
        #Write the data to the file, then formatted headers
        df.to_excel(writer, sheet_name = N, index = False, header=False, startrow = R+1, startcol = C)
        worksheet = writer.sheets[N]
        #Apply percent format to '%' column
        worksheet.set_column(C+5, C+5, None, f_percent)
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(R, C+col_num, value, f_header)

        worksheet.write_string(R-1, C, weekname[weekcount], f_week)
        C2 = C2 + 7

        #Part 3
        sql = sql_3
        R = R3
        C = C3
        N = "# of LinkIt Benchmarks"
        #worksheet = writer.sheets[N]

        #Get data from database
        df = pd.read_sql(sql_week(sql, week), connection)
        #Add 'Total' row at top
        df.loc[-1] = ['Total', '', df['TotalResults'].sum(), df['OnlineTests'].sum(), df['BubbleSheets'].sum()]  # adding a row
        df.index = df.index + 1  # shifting index
        df = df.sort_index()  # sorting by index

        #Write the data to the file, then formatted headers
        df.to_excel(writer, sheet_name = N, index = False, header=False, startrow = R+1, startcol = C)
        worksheet = writer.sheets[N]
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(R, C+col_num, value, f_header)

        worksheet.write_string(R-1, C, weekname[weekcount], f_week)
        worksheet.set_row(R+1, None, f_header)

        C3 = C3 + 6
        #Part 4A
        sql = sql_4A
        R = R4A
        C = C4
        N = "# of Online by Date"

        #Get data from database
        df = pd.read_sql(sql_week(sql, week), connection)
        #Write the data to the file, then formatted headers
        df.to_excel(writer, sheet_name = N, index = False, header=False, startrow = R+1, startcol = C)
        worksheet = writer.sheets[N]
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(R, C+col_num, value, f_header)

        worksheet.write_string(R-2, C, weekname[weekcount], f_week)
        worksheet.write_string(R-1, C, "By Start Date", f_header)

        #Part 4B
        sql = sql_4B
        R = R4B
        #Get data from database
        df = pd.read_sql(sql_week(sql, week), connection)
        #Write the data to the file, then formatted headers
        df.to_excel(writer, sheet_name = N, index = False, header=False, startrow = R+1, startcol = C)
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(R, C+col_num, value, f_header)

        worksheet.write_string(R-2, C, weekname[weekcount], f_week)
        worksheet.write_string(R-1, C, "By Last Login Date", f_header)

        C4 = C4 + 8

        #Part 5
        sql = sql_5
        R = R5
        C = C5
        N = "# of Online by Hour"

        #Get data from database
        df = pd.read_sql(sql_week(sql, week), connection)
        #Add some columns at the end
        df['Others'] = df['Number of Sessions'] - (df['A Beka'] + df['BEC'] + df['Frogstreet'])
        df['% of A Beka'] = df['A Beka'] / df['Number of Sessions']
        df['% of BEC'] = df['BEC'] / df['Number of Sessions']
        df['% of Frog Street'] = df['Frogstreet'] / df['Number of Sessions']
        df['% of Others'] = df['Others'] / df['Number of Sessions']





        #Write the data to the file, then formatted headers
        df.to_excel(writer, sheet_name = N, index = False, header=False, startrow = R+1, startcol = C)
        worksheet = writer.sheets[N]
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(R, C+col_num, value, f_header)

        worksheet.write_string(R-2, C, weekname[weekcount], f_week)
        worksheet.write_string(R-1, C, "By Last Login Date", f_header)
        #Adjust column widths
        #for i, width in enumerate(get_col_widths(df)):
        #    worksheet.set_column(C+i, width)
        C5 = C5 + 11

        #Part 6A
        sql = sql_6A
        R = R6A
        C = C6
        N = "# of Data Locker"

        #Get data from database
        df = pd.read_sql(sql_week(sql, week), connection)

        #Fill in missing days (almost always just christmas)
        #df = fill_dates(df, week)


        #Write the data to the file, then formatted headers
        df.to_excel(writer, sheet_name = N, index = False, header=False, startrow = R+1, startcol = C)
        worksheet = writer.sheets[N]
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(R, C+col_num, value, f_header)

        worksheet.write_string(R-1, C, weekname[weekcount], f_week)

        #Add Total to bottom
        worksheet.write_string(R+8, C, "Total", f_header)
        worksheet.write(R+8, C+1, df["Total"].sum(), f_header)


        if (weekcount == 0):
            worksheet.write_string(R-2, C, "By Date", f_week)

        #Part 6B
        sql = sql_6B
        R = R6B
        #Get data from database
        df = pd.read_sql(sql_week(sql, week),connection)
        #Write the data to the file, then formatted headers
        df.to_excel(writer, sheet_name = N, index = False, header=False, startrow = R+1, startcol = C)
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(R, C+col_num, value, f_header)

        worksheet.write_string(R-1, C, weekname[weekcount], f_week)
        if (weekcount == 0):
            worksheet.write_string(R-2, C, "By Client", f_week)
        #Adjust column widths
        #for i, width in enumerate(get_col_widths(df)):
        #    worksheet.set_column(C+i, width)
        C6 = C6 + 4

        weekcount = weekcount + 1
    '''Handle some formatting the easier way (actually opening an invisible
       instance of Excel and making it do our dirty work). Mostly for column
       widths, but it happened to be a neat way to format some percents.'''
    path = os.path.join(os.getcwd(),"Usage Reports", name)
    writer.save()
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(path)
    for ws in wb.Worksheets:
        ws.Columns.AutoFit()
    ws = wb.Worksheets("# of Results by Date")
    ws.Columns(1).ColumnWidth = 15
    ws = wb.Worksheets("# of Online by Hour")
    ws.Range('G:J,R:U,AC:AF').NumberFormat = '0%'
    ws = wb.Worksheets("# of Data Locker")
    ws.Range("A4:J10")
    if os.path.exists(path):
        os.remove(path)
    wb.SaveAs(path)
    excel.Application.Quit()

def create_report():
    try:
        main()
    except Exception as ex:
        print(ex, type(ex))
        raise ex
        return False
    else:
        return True
