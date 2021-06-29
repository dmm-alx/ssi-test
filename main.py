#import sys

from requests.sessions import session
#from pyasn1.type.univ import Null
#sys.path.append("./lib")
from datetime import datetime,timedelta
import calendar
import time
from googleapiclient.discovery import build
from google.oauth2 import service_account
import SSheet

SCOPES = [
            'https://www.googleapis.com/auth/admin.reports.audit.readonly',
            'https://www.googleapis.com/auth/drive.activity.readonly',
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
#Defs
SERVICE_ACCOUNT_FILE = "creds.json"
SPREADSHEET_ID = "19DJieQwDNmwoLwAFctefhdCu9Tel-brgrGQJVTXijJ4"
SHEET_NAME = "ShareDriveList"
USER_EMAIL = "is-tk@kcgrp.jp"
FOLDER_ID = "0AMAekuCS_A73Uk9PVA"
DIFF_JST_FROM_UTC = 9


def main(request):

    credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    delegated_credentials = credentials.with_subject(USER_EMAIL)

    # Initialize services
    service_reports = build('admin', 'reports_v1', credentials=delegated_credentials)
    service_sheet = build('sheets', 'v4', credentials=delegated_credentials)
    service_drive = build("drive", "v3", credentials=delegated_credentials)

    sharedDList = SSheet.getSSValues(service_sheet,SPREADSHEET_ID,SHEET_NAME)
    
    #sharedDriveID = '0AHyIMBjUcuZzUk9PVA'
    pageToken = None
    pageSize = 100
    maxResults=1000
    applicationName="drive"
    now = datetime.now()
    
    eventNames =[
                    'delete',
                    'view',
                    'edit',
                    'move',
                    'trash',
                    'download',
                    'rename',
                    'upload',
                    'create',
                    'change_user_access',
                    'change_document_visibility',
                    'change_document_access_scope',
                    'shared_drive_membership_change'
                ]
    if sharedDList:
        
        i = 1
        #remove first line
        sharedDList.pop(0)
        period = getPeriod()
        # get exe start time
        exestarttime = time.time()

        for sharedDriveID in sharedDList:
            
            # Call the Admin SDK Reports API
            print('Getting the events data -----------------------------')
            
            nowutc = datetime.utcnow() + timedelta(hours=DIFF_JST_FROM_UTC)
            nowutc = nowutc.isoformat(" ")
            i += 1
            #print(sharedDriveID)
            # Get the Sheet ID or Create new Sheet and get the ID
            if len(sharedDriveID) >= 3:
                sheet_id = sharedDriveID[2]

                # Check if run completed for this day continue
                lastRun = sharedDriveID[3].split(" ")
                (year,month,day) = [int(s) for s in lastRun[0].split("-")]
                
                if datetime(year,month,day).date() == now.date() and int(sharedDriveID[4])==1:
                    print("already run")
                    continue
                
                # 2 when process took long and was interrupted
                if sharedDriveID[4]==2:
                    lastRow = SSheet.getLastRow(service_sheet,sheet_id,"シート1")
                    lastDate = SSheet.getSSFixValues(service_sheet,sheet_id,"シート1","!D"+str(lastRow))
                    period = getLastLineDate(lastDate)
                    
                #body = {'values': [[nowutc]]}
                #SSheet.writeToSS(service_sheet,SPREADSHEET_ID,SHEET_NAME,body,'!D'+str(i))

            else:
                fname = str(now.year)+str(now.month)+"-"+sharedDriveID[1]
                r = SSheet.creatSpreadSheet(service_drive,FOLDER_ID,fname)
                sheet_id = r['id']
                #body = {'values': [[sheet_id,nowutc]]}
                body = {'values': [[sheet_id]]}
                SSheet.writeToSS(service_sheet,SPREADSHEET_ID,SHEET_NAME,body,'!C'+str(i))
                
            # Iterate through the event
            for eventName in eventNames:
                args = {
                        'userKey': "all",
                        'maxResults': pageSize,  # Accepted range is [1, 1000]
                        'eventName':eventName,
                        'startTime': period['start'],
                        'endTime': period['end'],
                        'filters': "shared_drive_id=="+sharedDriveID[0],
                        'maxResults':maxResults
                    }
                
                while True:
                    stat=1
                    # check elapsed time and break
                    # get diff time
                    diff = time.time() - exestarttime
                    print(diff)
                    #return
                    if diff > 1500:
                        stat = 2
                        timeout = {'values': [[nowutc,stat]]}
                        SSheet.writeToSS(service_sheet,SPREADSHEET_ID,SHEET_NAME,timeout,'!D'+str(i))
                        break

                    try:
                        results = service_reports.activities().list(
                            applicationName=applicationName, 
                            pageToken=pageToken, 
                            **args).execute()
                        
                        pageToken = results.get('nextPageToken') 
                        activities = results.get('items', [])
                        
                    except Exception as e:
                        print(e)
                        return e

                    else:
                        # Get last row number of the Sheet
                        last_row_id = SSheet.getLastRow(service_sheet,sheet_id,"シート1")

                        # Set the Header row of the Sheet 
                        if last_row_id==0:
                            headers = getHeader()
                            body = {'values': [headers]}
                            SSheet.writeToSS(service_sheet,sheet_id,"シート1",body,'!A1')

                        if not activities:
                            # Set status as not complete while looping in events
                            exetime = {'values': [[nowutc,"2"]]}
                            SSheet.writeToSS(service_sheet,SPREADSHEET_ID,SHEET_NAME,exetime,'!D'+str(i))
                            print('Logs無し')
                        else:
                            res = getActivityLine(activities)
                            body = {'values': res}
                            # Add rows to sheet
                            SSheet.appendToSS(service_sheet,sheet_id,"シート1",body,'!A'+str(last_row_id+1))
                            # Set status as not complete while looping in events
                            exetime = {'values': [[nowutc,"2"]]}
                            SSheet.writeToSS(service_sheet,SPREADSHEET_ID,SHEET_NAME,exetime,'!D'+str(i))
                            print('Logs有り')
                    if pageToken is None:
                        break
                    
            if now.day == 1:
                # reinitialise sheet for new month
                resetsheet = {'values': [["","","",""]]}
                SSheet.writeToSS(service_sheet,SPREADSHEET_ID,SHEET_NAME,resetsheet,'!C'+str(i))
            else:
                exetime = {'values': [[nowutc,stat]]}
                SSheet.writeToSS(service_sheet,SPREADSHEET_ID,SHEET_NAME,exetime,'!D'+str(i))

    return "complete"

# Format the the activity line to write
def getActivityLine(activities):
    res = []
    for activity in activities:
        params = getParams(activity['events'][0]['parameters'])
        Visibility = ""
        Old_visibility = ""
        Visitor = ""

        if 'old_visibility' in params and params['old_visibility']=='shared_internally': Old_visibility = '内部で共有中'
        else: Old_visibility = ""

        if 'visibility' in params and params['visibility']=='shared_internally': Visibility = '内部で共有中'
        else: Visibility = params['visibility']
        
        if 'visitor' in params: Visitor = params['visitor'] 
        else: Visitor = "No"

        if 'email' in activity['actor']: email = activity['actor']['email']
        else: email = activity['actor']['key']

        if 'ipAddress' in activity: ip = activity['ipAddress']
        else: ip = 'empty'

        res.append([
            params['doc_title'],
            email+'がアイテムを'+activity['events'][0]['name']+'しました',
            email,
            activity['id']['time'],
            activity['events'][0]['name'],
            params['doc_id'],
            params['doc_type'],
            params['owner'],
            Old_visibility,
            Visibility,
            ip,
            params['billable'],
            Visitor
        ])
    return res

# Check and format values
def getParams(parameters):
    params = {}
    for param in parameters:
        #print(param)
        if 'value' in param:
            params[param['name']] = param['value']
        elif 'boolValue' in param:
            if param['boolValue'] == True:
                params[param['name']] = 'Yes'
            else:
                params[param['name']] = 'No' 
        elif 'multiValue' in param:
            params[param['name']] = ",".join(param['multiValue'])
        else:
            params[param['name']] = param['value']
    return params

# Get last line date
def getLastLineDate(lastDate):
    
    DateTimeZone = lastDate.split("T")
    (year,month,day) = [int(s) for s in DateTimeZone[0].split("-")]
    dotsplit = DateTimeZone[1].split(".")
    (hour,min,sec) = [int(s) for s in dotsplit[0].split(":")]
    
    st = datetime(year,month,day,hour,min,sec)
    ed = datetime(year,month,day, 23, 59, 59)
    startTime = st.isoformat('T') + "Z"
    endTime = ed.isoformat('T') + "Z"
    return {'start':startTime, 'end':endTime}
    
# get the the period
def getPeriod():
    now = datetime.now()
    
    if now.day==1 and now.month==1:
        st = datetime(now.year-1, 12, 31, 0, 0, 0)
        ed = datetime(now.year-1, 12, 31, 23, 59, 59)
    elif now.day == 1 :
        lastDayofLastMonth = calendar.monthlen(now.year, now.month-1)
        st = datetime(now.year, now.month-1, lastDayofLastMonth, 0, 0, 0)
        ed = datetime(now.year, now.month-1, lastDayofLastMonth, 23, 59, 59)
    else:
        st = datetime(now.year, now.month, now.day-1, 0, 0, 0)
        ed = datetime(now.year, now.month, now.day-1, 23, 59, 59)
    
    startTime = st.isoformat('T') + "Z"
    endTime = ed.isoformat('T') + "Z"
    return {'start':startTime, 'end':endTime}

# Get Headers of the data sheet
def getHeader():
    headers = [
          'アイテム名', 
          'イベントの説明', 
          'ユーザー', 
          '日付', 
          'イベント名', 
          'アイテム ID', 
          'アイテムタイプ', 
          'オーナー', 
          '以前の公開設定',
          '公開設定', 
          'IP アドレス', 
          '請求可能',
          'ビジター'
        ]
    return headers

if __name__ == '__main__':
    main()
