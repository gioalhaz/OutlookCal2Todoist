import configparser
import requests
import locale
import win32com.client as win32
import time
import datetime
from datetime import timezone
import todoist


def get_outlook_calendar_entries(days = 1):
    """
    Returns calender entries for 'days' period
    """
    Outlook = win32.Dispatch('outlook.application')    
    
    ns = Outlook.GetNamespace("MAPI")
    appointments = ns.GetDefaultFolder(9).Items
    appointments.Sort("[Start]")
    appointments.IncludeRecurrences = "True"
    
    date_from = datetime.datetime.today()
    begin = date_from.date().strftime("%x")
    
    date_to = datetime.timedelta(days=(days+1)) + date_from
    end = date_to.date().strftime("%x")

    filter = "[Start] >= '" + begin + "' AND [END] <= '" + end + "'"

    print(filter)
    
    appointments = appointments.Restrict(filter)
    events = []
    
    for a in appointments:
        #print("from appointment " + str(a.Start))
        event_date = a.Start.replace(tzinfo=timezone(datetime.timedelta(seconds=time.localtime().tm_gmtoff)))
        events.append([event_date, a.Subject, a.Duration, a.Location])
        
    return events


# == Main =====

INI_FILE_NAME = "OutlookCal2Todoist.ini"

print("Start")

config = configparser.ConfigParser()
config.read(INI_FILE_NAME)

api_token = config["todoist"]["api_token"]
api_base_url = config["todoist"]["api_base_url"]
project_id = int(config["todoist"]["project_id"])
label_id = int(config["todoist"]["label_id"])
days_count = int(config["config"]["days"])

locale.setlocale(locale.LC_ALL, locale.getdefaultlocale()[0])


td_api = todoist.Todoist()
td_api.connect(api_base_url, api_token)

events = get_outlook_calendar_entries(days_count)

td_api.delete_tasks(project_id)

if len(events) != 0:

    time_now = datetime.datetime.now().replace(tzinfo=timezone(datetime.timedelta(seconds=time.localtime().tm_gmtoff)))
    for event in events:
        content = event[1] if len(event[3]) == 0 else f"{event[1]} ({event[3]})"

        #print(f"{event[0]}  {type(event[0])})")
        if event[0] > time_now:
            #print("yes")
            date_string = event[0].isoformat("T")
            td_api.add_new_task(project_id, content, date_string, label_id)
else:
    print(f"There is no events in {days_count} days period")
