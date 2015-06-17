import win32com.client, datetime as dt
from dateutil.relativedelta import relativedelta
from datetime import datetime
import os
from apiclient.http import BatchHttpRequest
from apiclient.discovery import build
from httplib2 import Http
import oauth2client
from oauth2client import client
from oauth2client import tools

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None


###Connects to outlook and creates an object of all items in folder 9
Outlook = win32com.client.Dispatch("Outlook.Application")
ns = Outlook.GetNamespace("MAPI")
appointments = ns.GetDefaultFolder(9).Items     #Folder 9 is the calendar folder, although I did not see this documented anywhere
appointments.Sort('[Start]')                # Without these two lines, recurring appointments will give wierd results
appointments.IncludeRecurrences = True      # Found the solution here - https://msdn.microsoft.com/en-us/library/office/ff866969(v=office.14).aspx
                                            # Had to translate example #2 (from link above) from VBA to python.


SCOPES = 'https://www.googleapis.com/auth/calendar'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'gcal-exch' #'JS_cal_import'
ACCOUNT_CAL = 'primary' # for Google Apps calendard, use email account ex example@the-example.com

# Get the AppointmentItem objects
# http://msdn.microsoft.com/en-us/library/office/aa210899(v=office.11).aspx


# takes appointments and Restrict to items in the specified days
# Found most of the below code at: http://stackoverflow.com/questions/21477599/read-outlook-events-via-python
def restrictedItems(appointmentst):
    begin = dt.date.today()
    end = begin + dt.timedelta(days = 7); #Change days to equal how many days in the future (today inclusive) to retrieve.
    restriction = "[Start] >= '" + begin.strftime("%m/%d/%Y") + ' 12:00 AM'"' AND [End] <= '" +end.strftime("%m/%d/%Y") + ' 12:00 AM'"'"
    restrictedItems1 = appointmentst.Restrict(restriction)
    return restrictedItems1

###takes the exchange time and converts it to rfc 3339 format for googal calendar
## This is a total hack.  I don't have a firm grasp on datetime, so use my code with caution
def time_conv(exch_time):
    exch_time = str(exch_time).replace(' ','/')
    exch_time = exch_time.replace(':','/')
    inter_time = exch_time.split('/')
    gcal_time = "20{0}-{1}-{2}T{3}:{4}:{5}-04:00".format(inter_time[2], inter_time[0],
                                                        inter_time[1], inter_time[3],
                                                        inter_time[4], inter_time[5])
    return gcal_time
 



### This code grabs (or creates) credentials in the form of a json file from your computer.  Then validates oauth2
### It's important that you set up permissions in the google api dashboard and create a project or the code won't work
### I found this code here: https://developers.google.com/google-apps/calendar/quickstart/python
def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   APPLICATION_NAME)

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatability with Python 2.6
            credentials = tools.run(flow, store)
        print 'Storing credentials to ' + credential_path
    return credentials


def main():
    ### The code below does the following:
    ### 1.) Get's the credentials for the google calendar
    ### 2.) Loads the Exchange events
    ### 3.) Queries the google calendar for the Exchange EventID (if it's already there, skips and moves on)
    ### 4.) If the Event Summary has the word "canceled" in it, moves on (in my organization, it's common to cancel a meeting this way.  frustrating)
    ### 5.) Define the event for google
    ### 6.) Load the event into Google Calendar
    ### 7.) If there's an error, switch to Update (This happened to me when I deleted an event on Google Calendar)
    ###
    ### Some of the links I referenced for the code in this area came from the following:
    ### https://developers.google.com/google-apps/calendar/v3/reference/events/insert#examples
    ### https://developers.google.com/google-apps/calendar/v3/reference/events/update
    ### https://msdn.microsoft.com/en-us/library/office/ff866969(v=office.14).aspx
    ### https://msdn.microsoft.com/en-us/library/office/ff869597(v=office.14).aspx
    ### http://stackoverflow.com/questions/466345/converting-string-into-datetime
    ### http://www.gossamer-threads.com/lists/python/python/513992

    credentials = get_credentials()
    service = build('calendar', 'v3', http=credentials.authorize(Http()))

    for appointmentItem in restrictedItems(appointments):
        noprint = False
        try:
            gcal_event = service.events().get(calendarId=ACCOUNT_CAL,summary=appointmentItem.Subject, eventId=str(appointmentItem.EntryID)).execute() #This is the google query to bounce against so no repeats
        except:
            #The above statement returns an error if the calendar event does not exist so we should create it.
            if 'canceled' in appointmentItem.Subject.lower():# is True:  #If subject line has Canceled in it, move on.
                noprint = True
         
            else: # 
                print dt.date.today()
                
                print("{0} Start: {1}, End: {2}, Organizer: {3}, Recurring: {4}, EntryID: {5}".format(
                  appointmentItem.Subject, appointmentItem.Start, 
                  appointmentItem.End, appointmentItem.Organizer, appointmentItem.IsRecurring, appointmentItem.EntryID))
                GC_summary = appointmentItem.Subject
                GC_start_dateTime = time_conv(appointmentItem.Start)
                GC_end_dateTime = time_conv(appointmentItem.End)
                GC_organizer = appointmentItem.Organizer
                GC_event_id = str(appointmentItem.EntryID).lower()

                event = {
                  'summary': GC_summary,    #'test',
                  'start': {
                    'dateTime': GC_start_dateTime,   #'2015-06-12T09:00:00-07:00',  
                    'timeZone': 'America/New_York',
                  },
                  'end': {
                    'dateTime': GC_end_dateTime,
                    'timeZone': 'America/New_York',
                  },
                  'id': GC_event_id,
                }
                try:
                    event = service.events().insert(calendarId=ACCOUNT_CAL, body=event).execute()
                    print 'Event created: %s' % (event.get('htmlLink'))
                    print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
                except: #if there is a collision error, switch to update
                    # First retrieve the event from the API.
                    event = service.events().get(calendarId=ACCOUNT_CAL, eventId=GC_event_id).execute()
                    event = {
                      'summary': GC_summary,    #'test',
                      'start': {
                        'dateTime': GC_start_dateTime,   #'2015-06-12T09:00:00-07:00',  
                        'timeZone': 'America/New_York',
                      },
                      'end': {
                        'dateTime': GC_end_dateTime,
                        'timeZone': 'America/New_York',
                      },
                      'id': GC_event_id,
                    }
                    updated_event = service.events().update(calendarId=ACCOUNT_CAL, eventId=GC_event_id, body=event).execute()
                    


if __name__ == '__main__':
    main()

'''
[DONE] Map Exchange properties to Google Propterties
    appointmentItem.Subject     -     GC_summary
    appointmentItem.Start       -     GC_start_dateTime
    appointmentItem.End         -     GC_end_dateTime
    appointmentItem.Organizer   -     
    appointmentItem.EntryID     -     
Placing Events in Google
    [DONE]Load the events from Exchange
    [DONE]Check Google Calendar for same EntryID
    [DONE kind of]Check Google Calendar for same dateTime and Subject
        If EntryID exists and dateTime is same, skip
        if EntryID exists and dateTime is different, delete Google Cal version insert Exchange version
        #if EntryID does not exist and dateTime and summary match, delete Google Calverioninsert Exchange version
        Else if Entry ID, Summary and dateTime are new, insert exchange item
Todo's
    [DONE]Fix Exchange restricted.  It's currently pulling from older dates.
    [DONE]re-flow the program to be properly laid out
    [DONE] delete unused code (moved to Test Code section)
    [DONE]comment each function to make clear
    [DONE]Change from jamesfaith account to spencer6524 acount
    [DONE]***translate Exchange date format to Google date format****



Test Code

    #Getting someone else's calendar
    recipient = ns.createRecipient("example@example.com")
    resolved = recipient.Resolve()
    sharedCalendar = ns.GetSharedDefaultFolder(recipient, 9)

    # Restrict to items in the next 30 days (using Python 3.3 - might be slightly different for 2.7)
    begin = dt.date.today()
    end = begin + dt.timedelta(days = 30);
    restriction = "[Start] >= '" + begin.strftime("%m/%d/%Y") + "' AND [End] <= '" +end.strftime("%m/%d/%Y") + "'"
    restrictedItems = sharedCalendar.Restrict(restriction)

    # Iterate through restricted AppointmentItems and print them
    for appointmentItem in restrictedItems:
        print("{0} Start: {1}, End: {2}, Organizer: {3}".format(
              appointmentItem.Subject, appointmentItem.Start, 
              appointmentItem.End, appointmentItem.Organizer))

#USed un Main() to query the google calendar

    now = datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
    print 'Getting the upcoming 10 events'
    eventsResult = service.events().list(
        calendarId=ACCOUNT_CAL, timeMin=now, maxResults=10, singleEvents=True,
        orderBy='startTime').execute()
    events = eventsResult.get('items', [])

    if not events:
        print 'No upcoming events found.'
    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        print start, event['summary']


#batch.add(service.events().insert(calendarId=ACCOUNT_CAL, body=event))
'''
'''
elif appointmentItem.IsRecurring is True: #  Fix date differences and check for end date
    if recur_check(appointmentItem) is False:  #check to see if still current returns flag true or false: True is still good False is expired
        noprint = True
    else: #do the thing.  Calculate the date and time of the recuring event  Make the below updater a function and call it here.
        recur_calc(appointmentItem)
def recur_check(apptItm):
    #This needs to check apptItm to see if it is expired.  Return False if it is, True if it isn't
    test_recur = apptItm.GetRecurrencePattern()
    test_recur_end = dt.datetime.strptime(str(test_recur.PatternEndDate), '%m/%d/%y %H:%M:%S')  #07/20/15 00:00:00
    if test_recur.NoEndDate == True:
        return True
    elif test_recur_end < dt.datetime.today():
        print 'Success',test_recur_end,dt.date.today()
        return False
    else:
        return True


def recur_calc(apptItm):
    test_recur = apptItm.GetRecurrencePattern()
    print 'Type:  ',test_recur.RecurrenceType
    print 'Occurrences:  ',test_recur.Occurrences
    print 'Pattern Start: ',test_recur.PatternStartDate
    print 'Pattern End:  ',test_recur.PatternEndDate
    print 'Duration: ',test_recur.Duration
    print 'Interval:  ',test_recur.Interval
    print 'No End Date:  ',test_recur.NoEndDate   
'''
'''
###TEst Point
for its in appointments:
    try:
        
        if its.IsRecurring is True:
            print '***********'
            print its.Subject
            print its.Start, its.End, its.IsRecurring
            test_recur = its.GetRecurrencePattern()
            print 'Type:  ',test_recur.RecurrenceType
            print 'Occurrences:  ',test_recur.Occurrences
            print 'Pattern Start: ',test_recur.PatternStartDate
            print 'Pattern End:  ',test_recur.PatternEndDate
            print 'Duration: ',test_recur.Duration
            print 'Interval:  ',test_recur.Interval
            print 'No End Date:  ',test_recur.NoEndDate
            print '***********'
    except:
            print "[*] [*]  Can't load this item [*] [*]"

###End TEst Point.  The challenge here is recurring items must be calculated to show on a calendar.  PITA
            ###Look here for the codes.  https://msdn.microsoft.com/en-us/library/office/ff863458.aspx


# This handles the callback from the batch request.  Raised errors should get handled here
def insert_event(request_id, response, exception):
  if exception is not None:
    # Do something with the exception
     pass
  else:
    # Do something with the response
    pass

    ###Part of the batch processing for google calendar
    
    #batch = BatchHttpRequest(callback=insert_event)
    # Iterate through restricted AppointmentItems and print them
'''
