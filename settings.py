from pathlib import Path

import pygsheets

gs_mgr = pygsheets.authorize(
    service_account_file=Path(__file__).parent.joinpath("src", "api_key.json")
).open_by_url("https://docs.google.com/spreadsheets/d/XXXXXXXXXXXXXXXXXXXXXXXXXX/")
