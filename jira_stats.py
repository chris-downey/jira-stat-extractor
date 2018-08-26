import sys
import configparser
import datetime as dt
import time
import requests
from requests.auth import HTTPBasicAuth
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
from xlsxwriter.utility import xl_col_to_name




##############################################################
#  
#  Write out as excel spreadsheet
#
##############################################################


def writeToExcel(json_obj, boardName):
   
  ts = time.time()
  print("Current System Time", ts)
  dtCurrentDate = dt.datetime.fromtimestamp(ts)
  
  st = dtCurrentDate.strftime('%Y-%m-%d-%H%M')
  output_file = boardName + "_" + st + ".xlsx"
  
  # Create a workbook and add a worksheet.
  workbook = xlsxwriter.Workbook(output_file)
  worksheet = workbook.add_worksheet('jira export')    
  
  # Add a bold format to use to highlight cells.
  bold_center = workbook.add_format({'bold': True, 'align': 'center'})
  bold_left = workbook.add_format({'bold': True, 'align': 'left'})
  center      = workbook.add_format({'align': 'center'})
  time_format = workbook.add_format({'align': 'center','num_format':'0.00'})
  #time_format_rounded = workbook.add_format({'align': 'center','num_format':'#0'})
  date_time_format = workbook.add_format({'align': 'center','num_format':'yyyy-mm-dd hh:mm'})
  date_format = workbook.add_format({'align': 'center','num_format':'yyyy-mm-dd'})
  # d"d" h"h" mm"m"
  # converting from milisecs to days
  # /(1000*60*60*24)
  
  
  ###################
  ## Write headers ##  
  ###################
  
  column_names = []
  column_names.append("Key")
  column_names.append("Summary")
  
  for col_index in range(len(json_obj['columns'])): 
     column_names.append(json_obj['columns'][col_index]['name'])
  
  column_names.append("Completion Date")
  column_names.append("Completion Month")
  column_names.append("Cycle Time")
  column_names.append("Lead Time")

  print("Headers: ", column_names)
  # List of issue cycle times in milliseconds

  cycleStartPosition = column_names.index(cycleStartState)
  cycleEndPosition = column_names.index(cycleEndState)
  
  leadStartPosition = column_names.index(leadStartState)
  leadEndPosition = column_names.index(leadEndState)  
  
  #print("Cycle Start Position", cycleStartPosition)

  export_date = json_obj['currentTime'] 

  ###################
  ## Write to File ##
  ###################
 
  # Write headers
  row = 0
  col = 0
  
  for name in column_names:
      worksheet.write(row, col, name, bold_center)
      col += 1
 
  # freeze the top pane to make it easier to review the output.
  worksheet.freeze_panes(1,0)
  
  # reset column/row coordinates for first line of data.
  row = 1
  
  # counter to track the position of the actual data that was extracted from jira.  
  # its purpose is to ensure the program can handle any column combinations.
  
  jira_column_start = 0
  jira_column_end = 0

  #  Note, this will be the coordinate e.g. A1, would be 0,0
  last_row_written_to = 0
  
  uniqueMonths = set()

  # reference points to be used later on for summaries and graphs

  cycleTimeColumn = 0  
  leadTimeColumn = 0
  dateCompletedColumn = 0
  
  if len(json_obj['issues']) > 0:
      
      #sort the data first as xlsxwriter doesn't have sort
      #functionality      
 
      ################################################################################     
      # The following works for sorting by key.
      #json_obj['issues'] = sorted(json_obj['issues'], key=lambda issue: issue['key'])    
      ################################################################################
      json_obj['issues'] = sorted(json_obj['issues'], key=lambda issue: int(issue['totalTime'][-1]),reverse=True)    
      
      
      for issue in json_obj['issues']:
        
        
        col = 0
        worksheet.write(row, col, issue['key'])   
        col += 1
        worksheet.write(row, col, issue['summary'])   
        col += 1
        
        # write out the actual timing data
        jira_column_start = col
     
        # calculate date the state was entered
        
        
        for time_spent in issue['totalTime']:        
            time_spent_converted = int(time_spent)/(1000*60*60*24)
            worksheet.write(row, col, time_spent_converted)
            col += 1
        
        jira_column_end = col-1
  
        
        # calculate the cycle time - get the column range as per what was configured in the ini file
   
        cycle_ref_start = xl_rowcol_to_cell(row, cycleStartPosition)  # e.g. C1        
        cycle_ref_end = xl_rowcol_to_cell(row, cycleEndPosition-1 )  # e.g. C5     
        
        # calculate the lead time
        lead_ref_start = xl_rowcol_to_cell(row, leadStartPosition)  # e.g. C1        
        lead_ref_end = xl_rowcol_to_cell(row, leadEndPosition-1 )  # e.g. C5     
        
        
        # For the forumula, you want to sum up the columns up to the end state
        # e.g. To Do...In Progresss....Live
        # you don't want to count the time spent in the live state, just the time up to that point
        
        cycle_time_formula = '=SUM(' + cycle_ref_start + ':' + cycle_ref_end + ')'         
        lead_time_formula = '=SUM(' + lead_ref_start + ':' + lead_ref_end + ')'
    
        #print("cycle time formula : ", cycle_time_formula)    
        #print("lead time formula : ", lead_time_formula)    
    
        time_in_done = issue['totalTime'][-1];
        #print("time in done : ", time_in_done)
        
        date_of_completion = 0
        
        # only write out the compelted month/date if its in done state
        # there is a bug in jira, that if you put it into done, more than once 
        # (e.g. to change the resolution state)
        # then jira marks it as a zero.
        # Without this check, today's date would be written out.
        if time_in_done != 0:
            date_of_completion = (export_date - time_in_done)/1000
          
            #formatted_completion_date = dt.datetime.fromtimestamp(date_of_completion).strftime('%Y-%m-%d %H:%M')
           
            completion_date = dt.datetime.fromtimestamp(date_of_completion)    
            month_completed = dt.datetime.fromtimestamp(date_of_completion).strftime('%Y-%m')
            #print("month completed ", month_completed)
            # use a set to extract the unique values to be used later on for the summary
            uniqueMonths.add(month_completed)        
                  
            worksheet.write_datetime(row, col, completion_date, date_time_format)
            worksheet.set_column(row,col,20)
            col += 1
            worksheet.write(row, col, month_completed)
            
            
        # increment, so the columns don't get messed up    
        else:
            col +=1
      
        dateCompletedColumn = xl_col_to_name(col-1)        
      
        col += 1

        worksheet.write(row, col, cycle_time_formula)
        
        cycleTimeColumn = xl_col_to_name(col)
        
        col += 1

        worksheet.write(row, col, lead_time_formula)

        leadTimeColumn = xl_col_to_name(col)


        # save off to help summary calculations (add 1 to convert coordinates to row number e.g. A1 = 0,0)
        last_row_written_to = row +1
        # make sure it gets formatted correctly
        lead_time_column = col
          
        row += 1
  

  # There is a limitation to doing an auto-fit for columns, so setting 
  # them a wee bit bigger to make the ouptut easier on the eye
  worksheet.set_column('A:A', 20, center)
  worksheet.set_column('B:B', 60)
  worksheet.set_column('C:M', 20, center)
      
  # set format for the core extracted times in each jira state- making sure to include the done column
        
  worksheet.set_column(jira_column_start, jira_column_end, 20, time_format)
  worksheet.set_column(lead_time_column-1,lead_time_column, 20, time_format)
  ########################################################################### 
  # extract some summary details
  ###########################################################################
  
   
  summaryColumns = []
  summaryColumns.append("Start Date")
  summaryColumns.append("End Date")
  summaryColumns.append("Tickets Completed")
  summaryColumns.append("Average Cycle Time")

  summaryColumns.append("Average Lead Time")
  
  summaryColumns.append("Median Cycle Time")
  summaryColumns.append("Median Lead Time")
  
  print("Unique months : ", uniqueMonths)
  summaryData = []
  
  dateCompletedRange = dateCompletedColumn
  dateCompletedRange += str(2)
  dateCompletedRange += ":"
  dateCompletedRange += dateCompletedColumn
  dateCompletedRange += str(last_row_written_to)  

  cycleTimeRange = cycleTimeColumn
  cycleTimeRange += str(2)
  cycleTimeRange += ":"
  cycleTimeRange += cycleTimeColumn
  cycleTimeRange += str(last_row_written_to)  

  #print("Cycle Time Range : ", cycleTimeRange)

  leadTimeRange = leadTimeColumn
  leadTimeRange += str(2)
  leadTimeRange += ":"
  leadTimeRange += leadTimeColumn
  leadTimeRange += str(last_row_written_to)  

  #print("lead Time Range : ", leadTimeRange)
        
  # writing out the headers
  summaryDataColumn = 0
  
  row += 1

  worksheet.write(row, summaryDataColumn, "Monthly Summary", bold_left)   
  
  row += 2  


  for title in summaryColumns:
      worksheet.write(row, summaryDataColumn, title, bold_center)   
      summaryDataColumn += 1
      
  row += 1    

  sortedMonths = sorted(uniqueMonths, key=lambda d: dt.datetime.strptime(d,'%Y-%m'))
  
  print("Sorted Unique Months: ", sortedMonths)  
  
  # NOTE - need to add one to the row counter as row = 0 is the same as row #1 in the spreadsheet.
  summaryDataRowTracker = row +1
  
  summaryData = []
  monthlyDates = []
  monthlyDates = deriveMonthlyDates(sortedMonths)
  
  summaryData = generateSummaryData(monthlyDates, cycleTimeRange, dateCompletedRange, summaryDataRowTracker, leadTimeRange)  
  
  print("summaryData size : ", len(summaryData))
  #row += len(summaryData)
      
  firstSummaryRow = row + 1    

  row = writeOutSummaryData(summaryData, worksheet, row, date_format, time_format)
  
  lastSummaryRow = row 

  row += 1
  
  
 
  ##############################################################
  # make some pretty graphs  
  # Note - seems a bit fidly to have to know all the locations, but its being done 
  # this way to make it dynamic vs. hardcoded as much as possible
  #
  # Hardcoding assumptions
  # A : Month
  # B : Tickets Completed
  # C : Average Cycle Time
  # D : Average Lead Time
  ##############################################################

  chart1 = workbook.add_chart({'type': 'line'})  
  
  monthSummaryRange = "$A$" +  str(firstSummaryRow) + ":" + "$A$" + str(lastSummaryRow)
  totalTicketRange = "$C$" +  str(firstSummaryRow) + ":" + "$C$" + str(lastSummaryRow)
  
  averageCycleSummaryRange = "$D$" +  str(firstSummaryRow) + ":" + "$D$" + str(lastSummaryRow)
  averageLeadSummaryRange = "$E$" +  str(firstSummaryRow) + ":" + "$E$" + str(lastSummaryRow)

  
  chart1.add_series({
    'name':       'Average Cycle Time',
    'categories': "='jira export'!" + monthSummaryRange,
    'values':     "='jira export'!" + averageCycleSummaryRange,
  })


  # Add a chart title and some axis labels.
  chart1.set_title ({'name': 'Average Cycle Time'})
  chart1.set_x_axis({'name': 'Months'})
  chart1.set_x_axis({'num_format' : 'mmm-yyyy'})

  chart1.set_y_axis({'name': 'Average Cycle Time (days)'})
 
  # Set an Excel chart style. From design tab in excel
  chart1.set_style(3)

  # Insert the chart into the worksheet (with an offset).

  averageCycleChartPos = "A" + str(lastSummaryRow)
  worksheet.insert_chart(averageCycleChartPos , chart1, {'x_offset': 10, 'y_offset': 30})
  
  
  # Lead Time Graph

  chart2 = workbook.add_chart({'type': 'line'})  
  
  chart2.add_series({
    'name':       'Average Lead Time',
    'categories': "='jira export'!" + monthSummaryRange,
    'values':     "='jira export'!" + averageLeadSummaryRange,
  })


  # Add a chart title and some axis labels.
  chart2.set_title ({'name': 'Average Lead Time'})
  chart2.set_x_axis({'name': 'Months'})
  chart2.set_x_axis({'num_format' : 'mmm-yyyy'})
  chart2.set_y_axis({'name': 'Average Lead Time (days)'})
 

  chart2.set_style(3)

  # Insert the chart into the worksheet (with an offset).

  averageCycleChartPos = "C" + str(lastSummaryRow)
  worksheet.insert_chart(averageCycleChartPos , chart2, {'x_offset': 10, 'y_offset': 30})

  # Ticket Count

  chart3 = workbook.add_chart({'type': 'line'})  
  
  chart3.add_series({
    'name':       'Total Completed Tickets',
    'categories': "='jira export'!" + monthSummaryRange,
    'values':     "='jira export'!" + totalTicketRange,
  })


  # Add a chart title and some axis labels.
  chart3.set_title ({'name': 'Total Ticket'})
  chart3.set_x_axis({'name': 'Months'})
  chart3.set_x_axis({'num_format' : 'mmm-yyyy'})
  chart3.set_y_axis({'name': 'Total Done Tickets'})
 
  # Set an Excel chart style. Colors with white outline and shadow.
  chart3.set_style(3)

  # Insert the chart into the worksheet (with an offset).

  averageCycleChartPos = "G" + str(lastSummaryRow)
  worksheet.insert_chart(averageCycleChartPos , chart3, {'x_offset': 10, 'y_offset': 30})

  
##############################################################
 #
 # Sprint Summary
 #
 ##############################################################

  print("####################################################")
  print("Sprint Summary Section")
  #print("sprintStartDate : " , sprintStartDate)
  #print("sprintStartLength : " , sprintLength)
  #sprintSummary

  sprintDates = []
  sprintDates = deriveSprintDates(sprintStartDate, dtCurrentDate)
  
  #leave some space after the graphs 
  row += 16
  
  summaryDataColumn = 0
  worksheet.write(row, summaryDataColumn, "Sprint Summary", bold_left)   
  row += 2
  
  for title in summaryColumns:
      worksheet.write(row, summaryDataColumn, title, bold_center)   
      summaryDataColumn += 1  
  row += 1  
  
  sprintSummaryData = []
  
  summaryRowTracker = row
  summaryRowTracker += 1
  sprintSummaryData = generateSummaryData(sprintDates, cycleTimeRange, dateCompletedRange, summaryRowTracker, leadTimeRange)  
  
  firstSummaryRow = row + 1    

  row = writeOutSummaryData(sprintSummaryData, worksheet, row, date_format, time_format)
  
  lastSummaryRow = row 
  
  workbook.close()    
  print("Write out to file complete!" , row , " rows written out")

  #calculateSprintSummary(sprintStartDate, sprintLength, totalTicketRange)

#################################################################
# writing it out to the excel
#################################################################
def writeOutSummaryData(summaryData, worksheet, row, date_format, time_format) :      

  #writing out the data
  summaryDataColumn = 0
  
  for columnData in summaryData:
      summaryDataColumn = 0  

      for element in columnData:
          # bit hacky, but write out first column (dates) without the formula method
          # scope to make this much nicer
          if summaryDataColumn==0 or summaryDataColumn==1:
              worksheet.write_datetime(row, summaryDataColumn, element, date_format)   
          else:
              worksheet.write_formula(row, summaryDataColumn, element, time_format)   
              
          summaryDataColumn += 1    
      # now move on the the next column
      row += 1
  lastMonthlySummaryRow = row    
  
  row += 1
  
  return row

def generateSummaryData(dateRange, cycleTimeRange, dateCompletedRange, summaryDataRowTracker, leadTimeRange)  :

    summaryData = []
    counter = 0
    for counter, month in enumerate(dateRange):
          totalTicketFormula = ""
          averageCycleTimeFormula = ""
          
          zeroCycleTimeChecker = cycleTimeRange + "," + "\"<>0\""
          
          totalTicketFormula = "=COUNTIFS(" + dateCompletedRange + "," +'">="' + "&A" + str(summaryDataRowTracker) + "," + dateCompletedRange + "," + '"<="' + "&B" + str(summaryDataRowTracker)  + ")"
         
          
          if excludeZeroCycleTimes == "TRUE" :
              averageCycleTimeFormula = "=AVERAGEIFS(" + cycleTimeRange + "," + dateCompletedRange + "," + '">="' + "&A" + str(summaryDataRowTracker) + "," + dateCompletedRange + "," + '"<="' + "&B" + str(summaryDataRowTracker) + "," + zeroCycleTimeChecker + ")"
          else : 
              averageCycleTimeFormula = "=AVERAGEIFS(" + cycleTimeRange + "," + dateCompletedRange + "," + '">="' + "&A" + str(summaryDataRowTracker) + "," + dateCompletedRange + "," + '"<="' + "&B" + str(summaryDataRowTracker) + ")"
              
          averageLeadTimeFormula = "=AVERAGEIFS(" + leadTimeRange + "," + dateCompletedRange + "," + '">="' + "&A" + str(summaryDataRowTracker) + "," + dateCompletedRange + "," + '"<="' + "&B" + str(summaryDataRowTracker) + ")"
         
          if excludeZeroCycleTimes == "TRUE" :
              medianCycleTimeForumla = "{=MEDIAN(IF(" + dateCompletedRange + ">=A" + str(summaryDataRowTracker)  + "," + "IF(" + dateCompletedRange + "<=B" + str(summaryDataRowTracker) + "," + "IF(" + cycleTimeRange +">0" +"," + cycleTimeRange + ")" + ")" + ")" + ")" + "}"
              #print("Median Cycle Time Formula :", medianCycleTimeForumla) 
          else : 
              medianCycleTimeForumla = "{=MEDIAN(IF(" + dateCompletedRange + ">=A" + str(summaryDataRowTracker)  + "," + "IF(" + dateCompletedRange + "<=B" + str(summaryDataRowTracker) + "," + cycleTimeRange + ")" + ")" + ")" + "}"
          
          
          medianLeadTimeForumla = "{=MEDIAN(IF(" + dateCompletedRange + ">=A" + str(summaryDataRowTracker)  + "," + "IF(" + dateCompletedRange + "<=B" + str(summaryDataRowTracker) + "," + leadTimeRange + ")" + ")" + ")" + "}"
         
          #"""
          print("Total Ticket Formula :", totalTicketFormula)    
          print("Average Cycle Time Formula :", averageCycleTimeFormula)   
          print("Average Lead Time Formula :", averageLeadTimeFormula)    
          print("Median Cycle Time Formula :", medianCycleTimeForumla)    
          print("Median Lead Time Formula :", medianLeadTimeForumla)    
          #"""
          # add a list for each column
          summaryData.append([])
          
          # Note - write them out as date objects and not strings in excel
          summaryData[counter].append(dateRange[counter][0]) 
          summaryData[counter].append(dateRange[counter][1]) 
          summaryData[counter].append(totalTicketFormula) 
          summaryData[counter].append(averageCycleTimeFormula) 
          summaryData[counter].append(averageLeadTimeFormula)       
          summaryData[counter].append(medianCycleTimeForumla)       
          summaryData[counter].append(medianLeadTimeForumla)            
              
          # increment for it's relative row position.
          summaryDataRowTracker += 1
          
    return summaryData

def deriveMonthlyDates(dateRange) :
    monthlyDates = []
    counter = 0
    for counter,month in enumerate(dateRange):  
        monthlyDates.append([])
        startMonth = dt.datetime.strptime(month, '%Y-%m')
          
        startDate = mkFirstOfMonth(startMonth)
        formattedStartDate = startDate.strftime('%d/%m/%Y')
        monthlyDates[counter].append(startDate)
        print("derived start date : ", formattedStartDate )
          
        # derive end date
        endDate = mkLastOfMonth(startMonth)
        formattedEndDate =   endDate.strftime('%d/%m/%Y')
        monthlyDates[counter].append(endDate)
        print("derived end date : ", formattedEndDate )
        counter += 1
    return monthlyDates
    
    
def deriveSprintDates(sprintStartDate, maxEndDate): 
    
    sprintDates = []    
    counter = 0     
    keepGoing = True
      
    dtSprintStartDate = dt.datetime.strptime(sprintStartDate, '%d/%m/%Y %H:%M:%S')
      
      # Derive the end of sprint date + next sprint date.  
      # Keep going until the next sprint date is in the future.     
    while keepGoing:
    
        sprintDates.append([])
        #print("sprint start date : ", dtSprintStartDate)      
        
        sprintDates[counter].append(dtSprintStartDate)
          
        # derive end date
        dtSprintEndDate = dtSprintStartDate + dt.timedelta(days=int(sprintLength),seconds=-1)
        #print("derived end date : ", endDate.strftime('%d/%m/%Y %H:%M:%S'))
        
        sprintDates[counter].append(dtSprintEndDate)
        
        dtSprintStartDate = dtSprintStartDate + dt.timedelta(days=int(sprintLength))
          
        counter += 1
        
          #print("sprint start date : ", dtSprintStartDate.date(), " current date " , dtCurrentDate.date())
          
        keepGoing = (dtSprintStartDate.date() < maxEndDate.date() )

    return sprintDates    
    

    
###############################################################
#
# Date Time Utilities
#   
###############################################################
    
def mkDateTime(dateString,strFormat="%Y-%m-%d"):
    # Expects "YYYY-MM-DD" string
    # returns a datetime object
    eSeconds = time.mktime(time.strptime(dateString,strFormat))
    return dt.datetime.fromtimestamp(eSeconds)

def formatDate(dtDateTime,strFormat="%Y-%m-%d"):
    # format a datetime object as YYYY-MM-DD string and return
    return dtDateTime.strftime(strFormat)

def mkFirstOfMonth2(dtDateTime):
    #what is the first day of the current month
    ddays = int(dtDateTime.strftime("%d"))-1 #days to subtract to get to the 1st
    delta = dt.timedelta(days= ddays)  #create a delta datetime object
    return dtDateTime - delta                #Subtract delta and return

def mkFirstOfMonth(dtDateTime):
    #what is the first day of the current month
    #format the year and month + 01 for the current datetime, then form it back
    #into a datetime object
    return mkDateTime(formatDate(dtDateTime,"%Y-%m-01"))

def mkLastOfMonth(dtDateTime):
    dYear = dtDateTime.strftime("%Y")
    dMonth = str(int(dtDateTime.strftime("%m"))%12+1)
    dDay = "1"
    if dMonth == '1':
       dYear = str(int(dYear)+1)
    nextMonth = mkDateTime("%s-%s-%s"%(dYear,dMonth,dDay))
    delta = dt.timedelta(seconds=1)
    return nextMonth - delta
    
##############################################################
#  
#  Connect to jira using the rest api
#
#  Note - if a swimlane is not used, then the API wont return
#         any issue....grrrr.... 
##############################################################


def configureURL(baseURL, rapidBoardID, jiraSection, swimlaneIDs):

# example target format : 
# https://your-company-jira/rest/greenhopper/1.0/rapid/charts/controlchart?rapidViewId=1159&swimlaneId=3050    

 
    # Append RapidView ID to parameters
    URL_params = '?rapidViewId=' + rapidBoardID
    for swimlane in swimlaneIDs:
        URL_params += '&swimlaneId=' + swimlane
 
    URL_all = baseURL + jiraSection + URL_params
    print("URL to connect : " , URL_all)
    return URL_all


def getSwimlanes(swimlaneConfig):

       ids = []

       for swimlane in swimlaneConfig:
           ids.append(str(swimlane['id']))

       return ids


#############################################################
#
# This is where the action happens
#
#############################################################
trimmed_data = []

configParser = configparser.ConfigParser(allow_no_value=True)
configParser.read('jira_stats.ini')

baseURL =                configParser.get('JIRAParams','baseURL')
boardID =                configParser.get('JIRAParams','boardID')
swimlaneID =             configParser.get('JIRAParams','swimLaneID')

cycleStartState =        configParser.get('JIRAParams', 'cycleStartState')
cycleEndState =          configParser.get('JIRAParams', 'cycleEndState')
leadStartState =         configParser.get('JIRAParams', 'leadStartState')
leadEndState =           configParser.get('JIRAParams', 'leadEndState')

excludeZeroCycleTimes =  configParser.get('JIRAParams', 'excludeZeroCycleTimes')
sprintStartDate =        configParser.get('JIRAParams', 'sprintStartDate')
sprintLength =           configParser.get('JIRAParams', 'sprintLength')

user =                   configParser.get('LoginParams','user')
pwd =                    configParser.get('LoginParams','pwd')


#print("Config Params Read - Board ID : ", boardID, 
#      " swim Lane : " , swimlaneID,
#      " cycleStartState : " , cycleStartState,
#      " cycleEndState : " , cycleEndState)

# if a specific swimlane hasn't been provided, then grab the board config from 
# jira and extract all swimlanes.

swimLanes = []

boardName = ""
    
  
boardURL =configureURL(baseURL, boardID,"rapidviewconfig/editmodel.json",[])
r = requests.get(boardURL,auth=HTTPBasicAuth(user,pwd))
     
print("Return code: ", r.status_code)
    # Check if request was a success
if r.status_code == requests.codes.ok:
    data = r.json()
    #print("Board Config Data Received : ", data)
    swimlaneConfig = data['swimlanesConfig']['swimlanes']
    boardName = data['name']        
    filterID = data['filterConfig']['id']    
    filterQuery = data['filterConfig']['query']
    
    
    print("swim lane config : ",swimlaneConfig)
    print("boardName : ",boardName)
    print("filterID : " , filterID)
    print("filterQuery : " , filterQuery)
    
    if swimlaneID == "":
        print("Specific swimlane not provided - use all available swimlanes")
        swimLanes = getSwimlanes(swimlaneConfig)    
        print("swim lanes : ",swimLanes)
        
    else :
        print("Specific swimlane provided")
        swimLanes.append(swimlaneID)

# This is the main call to get all the data back

URL_to_connect = configureURL(boardID,"rapid/charts/controlchart",swimLanes)
print("connection url : ", URL_to_connect)
print("Connecting.....")   

r = requests.get(URL_to_connect,auth=HTTPBasicAuth(user,pwd))
 
if r.status_code == requests.codes.ok:
    data = r.json()   
    
    #print("==========================================")
    #print("Raw data received: ", data)    
    #print("Trimming Data.......")
    del data['workRateData']
    trimmed_data.append(data)
    #print("Trimmed Data : " , trimmed_data)

else:
    sys.exit()


# Parse JSON
for obj in trimmed_data:
    #print("Looping object from trimmed data")
    writeToExcel(obj, boardName)


 





