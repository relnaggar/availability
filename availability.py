import datetime
import json
import pyperclip

from outlook_calendar import OutlookCalendar, yesnoinput


CACHE_PATH = "cache"
EVENTS_PATH = f"{CACHE_PATH}/events.json"

USER_PRINCIPAL_NAME = "ramsey.el-naggar@outlook.com"
TARGET_CALENDAR_NAMES = ["Calendar", "Tutoring"]

WORKDAYS = [1,2,3,4,5,6]
START_WORK_TIME = datetime.time(hour=15, minute=0)
END_WORK_TIME = datetime.time(hour=21, minute=0)


def create_cache_directory():
  import os
  if not os.path.exists(CACHE_PATH):
    os.makedirs(CACHE_PATH)

def get_cached_events():
  try:
    with open(EVENTS_PATH, "r") as f:
      events = json.load(f)
    return events
  except FileNotFoundError:
    pass
  return None

def cache_events(events):
  with open(EVENTS_PATH, "w") as f:
    json.dump(events, f, indent=2)

def get_events(start_datetime, end_datetime):
  outlook_calendar = OutlookCalendar(USER_PRINCIPAL_NAME, TARGET_CALENDAR_NAMES)
  events = outlook_calendar.get_events(start_datetime, end_datetime)
  cache_events(events)  
  return events

def get_ordinal_suffix(day):
  if 4 <= day <= 20 or 24 <= day <= 30:
    return "th"
  else:
    return ["st", "nd", "rd"][day % 10 - 1]

def is_overlap(start_datetime_1, end_datetime_1, start_datetime_2, end_datetime_2):
  latest_start = max(start_datetime_1, start_datetime_2)
  earliest_end = min(end_datetime_1, end_datetime_2)
  return latest_start < earliest_end

def main():
  create_cache_directory()

  now = datetime.datetime.now()
  end_of_next_week = (now + datetime.timedelta(days=14)).replace(hour=23, minute=59, second=59)

  start_datetime = now
  end_datetime = end_of_next_week
  events = get_cached_events()
  if events is None or yesnoinput("Use cached events? ", default="n") == "n":
    events = get_events(start_datetime, end_datetime)

  for event in events:
    event["startDateTime"] = datetime.datetime.fromisoformat(event["startDateTime"])
    event["endDateTime"] = datetime.datetime.fromisoformat(event["endDateTime"])

  valid = False
  while not valid:
    meeting_type = input("Meeting type (1: 55-minute lesson, 2: 15-minute meeting): ")
    if meeting_type == "1" or meeting_type == "2":
      valid = True
    else:
      print("Invalid meeting type")
  
  if meeting_type == "1":
    meeting_duration = datetime.timedelta(hours=1)
    actual_meeting_duration = datetime.timedelta(minutes=55)
  else:
    meeting_duration = datetime.timedelta(minutes=30)
    actual_meeting_duration = datetime.timedelta(minutes=15)

  event_number = 0
  today_date = datetime.date.today()
  end_date = end_datetime.date()
  current_date = today_date
  availabile_hours = []
  while current_date <= end_date:
    if current_date.isoweekday() in WORKDAYS:
      if current_date == today_date:
        start_of_next_hour_datetime = (start_datetime.replace(minute=0, second=0, microsecond=0) + datetime.timedelta(hours=1))
        start_hour_datetime = start_of_next_hour_datetime
      else:
        start_hour_datetime = datetime.datetime.combine(current_date, START_WORK_TIME)
      current_hour_datetime = start_hour_datetime
      # print("hour",current_hour_datetime)
      while current_hour_datetime.time() < END_WORK_TIME and current_hour_datetime < end_datetime:
        end_current_hour_datetime = current_hour_datetime+meeting_duration
        if START_WORK_TIME <= current_hour_datetime.time() and end_current_hour_datetime.time() <= END_WORK_TIME:
          while current_hour_datetime >= events[event_number]["endDateTime"] and event_number < len(events)-1:
            event_number += 1
            # print("event",events[event_number])
          if not is_overlap(current_hour_datetime, end_current_hour_datetime, events[event_number]["startDateTime"], events[event_number]["endDateTime"]):
            availabile_hours.append(current_hour_datetime)
        current_hour_datetime += meeting_duration
        # print("hour",current_hour_datetime)
    current_date += datetime.timedelta(days=1)

  availability_by_date = {}
  for hour_datetime in availabile_hours:
    date = hour_datetime.date()
    if date in availability_by_date:
      availability_by_date[date].append(hour_datetime)
    else:
      availability_by_date[date] = [hour_datetime]

  availability = f"Current availability for a {int(actual_meeting_duration.total_seconds() // 60)}-minute session (UK local time):"
  for date in sorted(availability_by_date.keys()):
    availability += "\n"
    availability += date.strftime(f"%A {date.day}{get_ordinal_suffix(date.day)} %B: ")
    for hour_datetime in availability_by_date[date][:-1]:
      availability += hour_datetime.strftime("%H:%M") + "-" + (hour_datetime+actual_meeting_duration).strftime("%H:%M") + ", "
    hour_datetime = availability_by_date[date][-1]
    availability += hour_datetime.strftime("%H:%M") + "-" + (hour_datetime+actual_meeting_duration).strftime("%H:%M")

  print("")
  print(availability)
  print("")

  pyperclip.copy(availability)
  print("Copied to clipboard")


if __name__ == "__main__":
  main()

