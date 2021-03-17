# Import from packages
import os
import argparse
import logging
import win32com.client

# Import from modules
from MyCls import initialise_app, finalise_app, handle_exception

# Initialise project
CURR_DIR, CURR_FILE = os.path.split(__file__)
PROJ_NAME = CURR_FILE.split('.')[0]

# Get command line arguments
my_arg_parser = argparse.ArgumentParser(description=f"{PROJ_NAME}")
#my_arg_parser.add_argument("arg1", help="Text1")
my_arg_parser.add_argument("--log", help="DEBUG to enter debug mode")
args = my_arg_parser.parse_args()

# Initialise app
initialise_app(PROJ_NAME, args.log)
logger = logging.getLogger("my_logger")

# # Get environment variables
# env_var1 = os.getenv("env_var1")
# env_var2 = os.getenv("env_var2")
# if env_var1 == None or env_var2 == None:
#     handle_exception("Missing environment variables!")

##################################################
# Variables
##################################################


##################################################
# Functions
##################################################


##################################################
# Main
##################################################
app = win32com.client.Dispatch("Outlook.Application")
my_namespace = app.GetNamespace("MAPI")
folder = my_namespace.GetDefaultFolder(9).Folders("[Import]") # 9 for Calendar folder

appt_itm = folder.Items.Add(1) # 1 for AppointmentItem object
appt_itm.Subject = "Test"
appt_itm.Start = "3/17/2021 12:00:00 AM"
appt_itm.End = "3/18/2021 12:00:00 AM"
appt_itm.AllDayEvent = True
appt_itm.ReminderSet = False
appt_itm.Save()

finalise_app()