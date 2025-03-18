import sys
import os

currentUserAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36'
payroll_path = os.path.dirname(os.path.dirname(__file__))
app_secret = '76043436eda19d85dc348a826df08dbc'
# human persona permanent page token
page_access_token = 'EAANzgCY9MAgBO0aOZARvI7FZAUBkAFZBLEwGz7pDcknYMVIVZAh98RIJKPPzE1N69g3ONKvB0faqEZB6Emipghtvo0CUunSsZBnOq3WuOUHdvtC5GphHvEJ9mDmcBK1itnZBlbjZCXMd3ZCl3YifKM5l70nGaisY5X4ad4fGUdLNhwJk3I7LeeFJOP1KM94oZAaZAQ6jCC6uQwZD'
system_user_token = 'EAANzgCY9MAgBOwImxslqEouzdoD54z21Qn2kKnzlSnvuuN4NzPmYICAi6HCtypEnVuNmr2ESj7ByNRtwDmDbliLd6zuj5CAVNGdC8NVPuIz8F068V5vsbZAYBcQJcWISSeK28iAI5qjaf6EwynAUH4NTZBzTzZBc9UtsOJ9tz8ZCKhibtopAqWkRaSIh8Qjv'
api_key = 'AIzaSyD01wgDEDh6BaTjmZor9pOsB_9ctOAOj4s'
    
page_id = '273928372944146'  # Human Persona
spreadsheet_id = '1wcH_Aa2Igpxz4SvOYFNvp_HGAlHuzbr1bOipVJ-lgn0'
fb_api_version = '19.0'
windows_user = os.environ.get('USERNAME')
sender_email = 'humanpersonaphils@gmail.com'
app_pw = 'uukm wjlc wwtm mogk'

def get_resource_path():
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath('.')
    return base_path

resource_path = get_resource_path()