from os import path, getenv
from sys import stdout, exit
from dotenv import load_dotenv

from re import search
from datetime import datetime, timedelta
from time import perf_counter_ns

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError



# Code mostly appropriated from Google API example
# https://developers.google.com/sheets/api/quickstart/python
def get_authorized_session():
    session = None
    if path.exists("token.json"):
        session = Credentials.from_authorized_user_file("token.json")
  
    if not session or not session.valid:
        if session and session.expired and session.refresh_token:
            session.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json",
                scopes=['https://www.googleapis.com/auth/spreadsheets'])
            session = flow.run_local_server(port=0)
  
    with open("token.json", "w") as token:
        token.write(session.to_json())
    return session


def get_time_values(user, cell_range):
  try:
    service = build("sheets", "v4", credentials=user)

    sheet = service.spreadsheets()
    response = (
        sheet.values()
        .get(spreadsheetId=SHEET_ID, range=cell_range)
        .execute()
    )
    return response.get("values", [])
  except HttpError as err:
    stdout.write("\n"+err)
    
def write_time_values(user, cell_range, cell_values):
    try:
        service = build("sheets", "v4", credentials=user)
        body = {"values": cell_values}
        response = (
            service.spreadsheets()
            .values()
            .update(
                spreadsheetId=SHEET_ID,
                range=cell_range,
                valueInputOption="USER_ENTERED",
                body=body,
            )
            .execute()
        )
        return response
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error


def generate_timedeltas(timestamps):
    # First get the time range, then add it to the sum.
    for i in range(0, len(timestamps), 2):
        span = timestamps[i+1]-timestamps[i]
        yield span if span.days >= 0 else span + timedelta(days=1)


def sum_timedeltas(deltas):
    time_sum = timedelta()
    for i in deltas:
        time_sum += i
    return time_sum


def timedelta_to_HM(td):
    hours = int(td.total_seconds() // 3600)
    minutes = int((td.total_seconds() - hours*3600) // 60)
    return ":".join(map(str, (hours, minutes)))



if __name__ == "__main__":
    load_dotenv()
    stdout.write("\nLoading .env...")
    SHEET_ID = getenv("GOOGLE_SHEET")
    if SHEET_ID:
        stdout.write("OK!")
    else:
        SHEET_ID = search(
            "(?<=docs\\.google\\.com\\/spreadsheets\\/d\\/)\\w+",
            input("\nSHEET_ID not found in .env.\nEnter spreadsheet URL: ")
        )
        while not SHEET_ID:
            SHEET_ID = input("\nInvalid URL. Enter spreadsheet URL: ")
            SHEET_ID = search(
                "(?<=docs\\.google\\.com\\/spreadsheets\\/d\\/)\\w+",
                SHEET_ID
            )
        SHEET_ID = SHEET_ID.group()
        
    
    stdout.write("\nAuthorizing...")
    auth_user = get_authorized_session()
    if auth_user:
        stdout.write("OK!")
    
    cell_range = None
    while not cell_range:
        cell_range = input("\nEnter cell range (default A:G): ")
        cell_range = cell_range if search(
            "(?:^[A-Z]+\\d+:[A-Z]+\\d+)|(?:^[A-Z]+:[A-Z]+)|(?:^\\d+:\\d+(?![A-Z]+))",
            cell_range) else "A:G"
    
    # Start timer
    begin_timer = perf_counter_ns()
    stdout.write(
        f"\nFetching Google Sheets API data for ranges: {cell_range}..."
    )
    sheet = get_time_values(auth_user, "A:G")
    complete_timer = perf_counter_ns()
    stdout.write(f"OK! Elapsed time: {complete_timer - begin_timer} ns")
    
    stdout.write("\n\nBeginning data operations.")
    headers = sheet[0]
    data = sheet[1:]
    
    indices = []
    users = []
    
    begin_timer = perf_counter_ns()
    # Yes, this would be faster if I requested 
    # column-major order from the Google Sheets API
    stdout.write("\nFiltering data...")
    
    for col, name in filter(
        lambda x: x[1] not in {"Tasks", "Clock in/out"},
        enumerate(headers)
    ):
        indices.append(col)
        users.append(name)
    
    complete_timer = perf_counter_ns()
    stdout.write(f"OK! Elapsed time: {complete_timer - begin_timer} ns")
    
    begin_timer = perf_counter_ns()
    # Assign timestamps to each user by index of column on row
    stdout.write("\nConverting timestamp strings...")
    time_stamps = {
        users[i]:
            list(
                map(lambda x: datetime.strptime(x, "%H:%M"),
                    filter(bool,
                    (row[col] for row in data if len(row) > col)
                    )
                )
            )
            for i, col in enumerate(indices)
    }   # Aplogogies to readability
        # Dict{ User's name: Non-empty entries in the set of extant entries }
    complete_timer = perf_counter_ns()
    stdout.write(f"OK! Elapsed time: {complete_timer - begin_timer} ns")

    # Create generators which yield the difference between 
    # every second timestamp and the previous timestamp.
    # If the hours overflow, a day is added.
    # e.g. 23:00 to 1:00 is not 1-23=-22, but 25-23=2
    begin_timer = perf_counter_ns()
    stdout.write("\nComputing timestamp deltas...")
    time_deltas = {
        name:generate_timedeltas(time_stamps[name]) for name in users
    }
    complete_timer = perf_counter_ns()
    stdout.write(f"OK! Elapsed time: {complete_timer - begin_timer} ns")
    
    # Sum the timedelta objects
    stdout.write("\nSumming time deltas...")
    begin_timer = perf_counter_ns()
    time_totals = {name:sum_timedeltas(time_deltas[name]) for name in users}
    complete_timer = perf_counter_ns()
    stdout.write(f"OK! Elapsed time: {complete_timer - begin_timer} ns")
    
    stdout.write("\nDisplaying results by name:\n\n")
    stdout.write("\n"+"\t|".join(time_totals.keys()))
    stdout.write("\n"+"\t|".join(
        timedelta_to_HM(time_totals[name]) for name in time_totals
    ) + "\n\n\n\n")
    
    
    # Write values to Google Sheets
    write_cells = input("Write to cells? [y/N]: ")
    if write_cells in {'y', 'Y'}:
        write_cells = input(
            f"Provide cell range with dimensions >= 2-wide and {len(users)}-high: "
        )
        write_cells = write_cells if search(
            "(?:^[A-Z]+\\d+:[A-Z]+\\d+)|(?:^[A-Z]+:[A-Z]+)|(?:^\\d+:\\d+(?![A-Z]+))",
            write_cells) else None
        while not write_cells:
            write_cells = input(
                "Invalid cell range." +
                f"\nProvide cell range with dimensions >= 2-wide and {len(users)}-high: "
            )
            write_cells = write_cells if search(
                "(?:^[A-Z]+\\d+:[A-Z]+\\d+)|(?:^[A-Z]+:[A-Z]+)|(?:^\\d+:\\d+(?![A-Z]+))",
                write_cells) else None
        
        write_time_values(
            auth_user,
            write_cells,
            [
                [name, timedelta_to_HM(time_totals[name])+":00.000"]
                for name in time_totals
            ]
        )

    stdout.write("\nExiting.\n")
    exit(0)