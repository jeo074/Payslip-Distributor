import sys
import os

payroll_path =os.path.abspath('.') #  Use with Auto-py-to-exe
payroll_path = os.path.dirname(os.path.dirname(__file__))  #  run in IDE

fb_api_version = '19.0'
windows_user = os.environ.get('USERNAME')
company_name = "THE COMPANY HR Management Services"

def get_resource_path():
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath('.')
    return base_path

resource_path = get_resource_path() # Use with Auto-py-to-exe
resource_path = payroll_path # run in IDE