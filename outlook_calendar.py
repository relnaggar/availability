import json
import time
import requests
import urllib.parse

def yesnoinput(prompt, default=None):
    valid = False
    user_input = ""
    while valid == False:
      user_input = input("(" + ("Y" if default == "y" else "y") + "/" + ("N" if default == "n" else "n") + ") " + prompt)
      if user_input == "" and default is not None:
        user_input = default
      if user_input in ["y", "n"]:
        valid = True
      else:
        print("invalid -- you must type 'y' or 'n'")
    return user_input

class OutlookCalendar:
  CACHE_PATH = "cache"
  ACCESS_TOKEN_PATH = f"{CACHE_PATH}/access_token.json"
  CALENDARS_PATH = f"{CACHE_PATH}/calendars.json"

  TENANT_ID = "consumers"
  CLIENT_ID = open('secrets/CLIENT_ID').read().strip()
  CLIENT_SECRET = open('secrets/CLIENT_SECRET').read().strip()
  HOST = 'localhost'
  PORT = 65432
  REDIRECT_URI = f"http://{HOST}:{PORT}"
  SCOPE = "calendars.read"

  def __init__(self, user_principal_name, target_calendar_names):
    self.user_principal_name = user_principal_name
    self.target_calendar_names = target_calendar_names

    self.access_token = self.get_cached_access_token()
    if self.access_token is None or yesnoinput("Use cached access_token? ", default="y") == "n":
      authorization_code = self.get_authorization_code()
      self.access_token = self.get_access_token(authorization_code)

    self.calendars = self.get_cached_calendars()
    if len(self.calendars) == 0 or yesnoinput("Use cached calendars? ", default="y") == "n":
      self.calendars = self.get_calendars()

  def get_cached_access_token(self):
    try:
      with open(self.ACCESS_TOKEN_PATH, "r") as f:
        data = json.load(f)
      now = time.time()
      if data["expiry_time"] - now > 10:
        return data["access_token"]
    except FileNotFoundError:
      pass
    return None

  def get_authorization_code(self):
    import random
    import pyperclip
    import socket

    state = random.randint(100,1000)
    authorize_url = f"https://login.microsoftonline.com/{self.TENANT_ID}/oauth2/v2.0/authorize?client_id={self.CLIENT_ID}&redirect_uri={urllib.parse.quote(self.REDIRECT_URI)}&scope={self.SCOPE}&response_type=code&response_mode=query&state={state}"

    print("User authorization")
    print("authorize_url:",authorize_url)
    pyperclip.copy(authorize_url)
    print("Copied to clipboard")
    print("")

    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
      s.bind((self.HOST, self.PORT))
      s.listen()
      print(f"Server listening on {self.HOST}:{self.PORT}")
      conn, addr = s.accept()
      with conn:
        print('Connected by', addr)
        data = conn.recv(4096).decode('utf-8')
        # Parse the GET request line to extract the path and query string
        request_line = data.split('\r\n')[0]  # e.g. "GET /?code=...&state=926 HTTP/1.1"
        path = request_line.split(' ')[1]      # e.g. "/?code=...&state=926"
        query_string = path.split('?', 1)[1]
        params = urllib.parse.parse_qs(query_string)
        assert params.get('state', [None])[0] == str(state)
        authorization_code = params['code'][0]
        print("Received authorization code")
    print("")
    return authorization_code

  def cache_access_token(self, access_token, expiry_time):
    cache = {
      "access_token": access_token,
      "expiry_time": expiry_time
    }
    with open(self.ACCESS_TOKEN_PATH, "w") as f:
      json.dump(cache, f, indent=2)
    print("Access token cached")

  def get_access_token(self, authorization_code):
    token_url = f"https://login.microsoftonline.com/{self.TENANT_ID}/oauth2/v2.0/token"
    headers = {
      "Content-Type": "application/x-www-form-urlencoded"
    }
    data = {
      "client_id": self.CLIENT_ID,
      "scope": self.SCOPE,
      "grant_type": "authorization_code",
      "code": authorization_code,
      "redirect_uri": self.REDIRECT_URI,
      "client_secret": self.CLIENT_SECRET
    }
    print("Requesting access token")
    start_time = time.time()
    response = requests.post(token_url, headers=headers, data=data, timeout=10)
    print(response.status_code)
    response_json = response.json()
    if response.status_code != 200:
      print("Error requesting access token:", response_json)
      raise Exception("Failed to get access token")
    access_token = response_json["access_token"]
    expiry_time = start_time + response_json["expires_in"]
    print("Received access token")
    # print(access_token)

    self.cache_access_token(access_token, expiry_time)
    return access_token

  def get_cached_calendars(self):
    try:
      with open(self.CALENDARS_PATH, "r") as f:
        calendars = json.load(f)
      return calendars
    except FileNotFoundError:
      pass
    return []

  def cache_calendars(self, calendars):
    with open(self.CALENDARS_PATH, "w") as f:
      json.dump(calendars, f, indent=2)
    print("Calendars cached")

  def get_calendars(self):
    headers = {
        "Authorization": f"Bearer {self.access_token}",
        "Content-Type": "application/json"
    }

    query_url = f"https://graph.microsoft.com/v1.0/users/{urllib.parse.quote(self.user_principal_name)}/calendars"
    print("Requesting list of calendars")  
    print("query_url:",query_url)
    response = requests.get(query_url, headers=headers, timeout=10)
    print(response.status_code)
    print("")

    data = response.json()
    # print(json.dumps(data, indent=2))

    calendars = []
    for calendar in data["value"]:
      if calendar["name"] in self.target_calendar_names:
        print("appending calendar named",calendar["name"])
        calendars.append({
          "id": calendar["id"],
          "name": calendar["name"]
        })
    print("")

    self.cache_calendars(calendars)
    return calendars

  def get_calendar_view(self, calendar, startDateTime, endDateTime):
    headers = {
        "Authorization": f"Bearer {self.access_token}",
        "Prefer": 'outlook.timezone="GMT Standard Time"'
    }

    params = {
        "startDateTime": startDateTime.isoformat(),
        "endDateTime": endDateTime.isoformat(),
        "$orderby": "start/dateTime asc",
    }

    query_url = f"https://graph.microsoft.com/v1.0/users/{urllib.parse.quote(self.user_principal_name)}/calendars/{calendar['id']}/calendarView?{urllib.parse.urlencode(params)}"
    print(f"Requesting calendar information for {calendar['name']}")
    print("params",params)

    querying = True
    datas = []
    while querying:
      print("query_url:",query_url)
      response = requests.get(query_url, headers=headers, timeout=10)
      print(response.status_code)
      print("")

      data = response.json()
      datas.append(data)
      try:
        query_url = data["@odata.nextLink"]
      except KeyError:
        querying = False

    events = []
    for data in datas:
      for event in data["value"]:
        events.append({
          "subject": event["subject"],
          "startDateTime": event["start"]["dateTime"],
          "endDateTime": event["end"]["dateTime"]
        })

    return events

  def get_events(self, startDateTime, endDateTime):
    events = []
    for calendar in self.calendars:
      events += self.get_calendar_view(calendar, startDateTime, endDateTime)
    return sorted(events, key=lambda event: event['startDateTime']) 

