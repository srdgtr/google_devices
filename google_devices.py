import json
import os.path
import pickle
from datetime import datetime

import numpy as np
import pandas as pd
import pandas.io.formats.excel
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import xlsxwriter

date_today: str = datetime.now().strftime("%c").replace(":", "-")

# If modifying these scopes, delete the file token.pickle.

SCOPES: list[str] = [
    "https://www.googleapis.com/auth/admin.directory.device.chromeos",
]

"""
Get all info on chromebooks.
"""
CREDS = None
token_file = "token.pickle"
# The file token.pickle stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists(token_file):
    with open(token_file, "rb") as token:
        CREDS = pickle.load(token)
# If there are no (valid) credentials available, let the user log in.
if not CREDS or not CREDS.valid:
    if CREDS and CREDS.expired and CREDS.refresh_token:
        CREDS.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
        CREDS = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open(token_file, "wb") as token:
        pickle.dump(CREDS, token)

service = build("admin", "directory_v1", credentials=CREDS)

device_list = []

NEXT_PLACE_TOKEN: str | None = "one"
PAGE_TOKEN: str | None = None

while NEXT_PLACE_TOKEN:
    get_chromebooks_list = service.chromeosdevices().list(
        customerId="my_customer",
        orderBy="serialNumber",
        projection="FULL",
        pageToken=PAGE_TOKEN,
        maxResults=200,
        sortOrder=None,
        query=None,
        fields="nextPageToken,chromeosdevices(deviceId, serialNumber, status, lastSync,\
        supportEndDate,annotatedUser, annotatedLocation, annotatedAssetId,notes,model,\
        meid,orderNumber,willAutoRenew,osVersion,platformVersion,firmwareVersion,macAddress,\
        bootMode,lastEnrollmentTime, orgUnitPath, recentUsers, ethernetMacAddress, activeTimeRanges,\
        tpmVersionInfo, cpuStatusReports, systemRamTotal, systemRamFreeReports, diskVolumeReports,\
        manufactureDate, autoUpdateExpiration,lastKnownNetwork)",
    )
    chromebooks_list = get_chromebooks_list.execute()
    NEXT_PLACE_TOKEN = None

    if chromebooks_list:
        chromebooks_list_dict = json.loads(str(chromebooks_list["chromeosdevices"]).replace("'", '"').replace("\\", ""))
        for aRow in chromebooks_list_dict:
            if aRow["status"] == "ACTIVE":
                device_list.append(aRow)
    if "nextPageToken" in chromebooks_list:
        PAGE_TOKEN = chromebooks_list["nextPageToken"]
        NEXT_PLACE_TOKEN = chromebooks_list["nextPageToken"]
    else:
        break


def total_usage(time_range: str | list[dict[str, str]]) -> int:
    if not isinstance(time_range, list):
        return 0
    total: list = []
    for time in time_range:
        total.append(time["activeTime"])
    return int(sum(total) / 6000)

def support_date(unix_time:str):
    if not isinstance(unix_time, str):
        return ""
    return datetime.fromtimestamp(int(unix_time)/1000)

devices = pd.DataFrame(device_list).assign(
    usage_minuten=lambda x: x["activeTimeRanges"].apply(total_usage),
    support_end_date=lambda x: x["autoUpdateExpiration"].apply(support_date),
    lastKnownNetwork=lambda x: x["lastKnownNetwork"].apply(lambda x: x if not isinstance(x, list) else x[0] if len(x) else ""),
    ipaddress=lambda x: x["lastKnownNetwork"][pd.notna(x["lastKnownNetwork"])].apply(lambda x: x["ipAddress"] if x is not np.nan else x),
    wanIpAddress=lambda x: x["lastKnownNetwork"][pd.notna(x["lastKnownNetwork"])].apply(
        lambda x: x["wanIpAddress"] if x is not np.nan else x
    ),
    recentUser=lambda x: x["recentUsers"].apply(lambda x: x if not isinstance(x, list) else x[0] if len(x) else ""),
    lastuser=lambda x: x["recentUser"][pd.notna(x["recentUser"])].apply(lambda x: x.get("email", "") if x is not np.nan else x),
)

devices["lastKnownNetwork"].fillna(value="onbekend", inplace=True)
os_versions = devices["osVersion"].value_counts().reset_index()
os_versions.columns = ["os_versions", "aantal"]
end_date = devices["support_end_date"].value_counts().reset_index()
end_date.columns = ["support_end_date", "aantal"]
chromebook_models = devices["model"].value_counts().reset_index()
chromebook_models.columns = ["chromebook_models", "aantal"]
chromebook_location = devices["annotatedLocation"].value_counts().reset_index()
chromebook_location.columns = ["chromebook_locatie", "aantal"]

nodig_voor_controlen = devices[
    [
        "serialNumber",
        "lastuser",
        "annotatedAssetId",
        "annotatedLocation",
        "notes",
        "osVersion",
        "model",
        "lastKnownNetwork",
    ]
]

num_rows: int = len(devices)
writer = pd.ExcelWriter(f"active_chrome_devices_{date_today}.xlsx", engine="xlsxwriter")
pandas.io.formats.excel.ExcelFormatter.header_style = None
devices.to_excel(writer, sheet_name="chromebooks", index=False, float_format="%.2f")
nodig_voor_controlen.to_excel(writer, sheet_name="controlelijst", index=False, float_format="%.2f")
os_versions.to_excel(writer, sheet_name="os_versions", index=False)
end_date.to_excel(writer, sheet_name="end_date_support", index=False)
chromebook_models.to_excel(writer, sheet_name="chromebook_models", index=False)
writer.sheets["os_versions"].hide()
writer.sheets["chromebook_models"].hide()
workbook = writer.book
rotate_items = workbook.add_format({"rotation": "30"})
ean_format = workbook.add_format({"num_format": "000000000000000"})
noip = workbook.add_format({"bg_color": "#57a639"})
zoekgeraakt = workbook.add_format({"bg_color": "#a52019"})
worksheet = writer.sheets["chromebooks"]
worksheet.freeze_panes(1, 0)
worksheet.set_row(0, 40, rotate_items)
worksheet.set_column("C:C", 20)
worksheet.set_column("F:F", 20)
worksheet.set_column("G:G", 20)
worksheet.conditional_format(
    f"$A$2:$B${num_rows}",
    {"type": "formula", "criteria": '=INDIRECT("X"&ROW())="onbekend"', "format": noip},
)
worksheet.conditional_format(
    f"$C$2:$C${num_rows}",
    {"type": "formula", "criteria": '=INDIRECT("BY"&ROW())<>""', "format": zoekgeraakt},
)

# toevoegen van chart met os aantallen
worksheet = workbook.add_chartsheet("os_version")
chart = workbook.add_chart({"type": "column"})
chart.set_title({"name": "aantallen van elke os versie actieve chromebooks"})
chart.set_style(3)
chart.set_plotarea({"gradient": {"colors": ["#33ccff", "#80ffff", "#339966"]}})
chart.set_chartarea({"border": {"none": True}, "fill": {"color": "#bfbfbf"}})
chart.add_series(
    {
        "name": "aantallen",
        "values": "=os_versions!$B$2:$B$20",
        "categories": "=os_versions!$A$2:$A$20",
        "gap": 25,
        "name_font": {"size": 14, "bold": True},
        "data_labels": {
            "value": True,
            "position": "inside_end",
            "font": {"name": "Calibri", "color": "white", "rotation": 345},
        },
    }
)
chart.set_x_axis(
    {
        "name": "os versions",
        "name_font": {"size": 14, "bold": True},
    }
)
chart.set_y_axis(
    {
        "major_unit": 20,
        "name": "aantal",
        "major_gridlines": {"visible": False},
    }
)
chart.set_legend({"none": True})
worksheet.set_chart(chart)

# toevoegen van chart met os aantallen
worksheet = workbook.add_chartsheet("aantallen chromebook")
chart = workbook.add_chart({"type": "column"})
chart.set_title({"name": "aantallen van elke model actieve chromebooks"})
chart.set_style(3)
chart.set_plotarea({"gradient": {"colors": ["#33ccff", "#80ffff", "#339966"]}})
chart.set_chartarea({"border": {"none": True}, "fill": {"color": "#bfbfbf"}})
chart.add_series(
    {
        "name": "aantallen",
        "values": "=chromebook_models!$B$2:$B$10",
        "categories": "=chromebook_models!$A$2:$A$10",
        "gap": 25,
        "name_font": {"size": 14, "bold": True},
        "data_labels": {
            "value": True,
            "position": "inside_end",
            "font": {"name": "Calibri", "color": "white", "rotation": 345},
        },
    }
)
chart.set_x_axis(
    {
        "name": "chromebook model",
        "name_font": {"size": 14, "bold": True},
    }
)
chart.set_y_axis(
    {
        "major_unit": 20,
        "name": "aantal",
        "major_gridlines": {"visible": False},
    }
)
chart.set_legend({"none": True})
worksheet.set_chart(chart)

writer.close()
