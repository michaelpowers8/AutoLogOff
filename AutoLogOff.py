import os
import json
import smtplib
import platform
import traceback
import subprocess
from time import sleep
from xml_logging import XML_Logger
from email.mime.text import MIMEText
from datetime import datetime,timedelta
from email.mime.multipart import MIMEMultipart
from tkinter import Tk, messagebox, simpledialog

def get_configuration() -> dict[str,str]|None:
    config_path:str = "Config.json"
    try:
        if(os.path.exists(config_path)):
            with open(config_path,"r",encoding="utf-8") as file:
                configuration = json.load(file)
            return configuration
        else:
            return None
    except:
        return None

def _verify_configuration(configuration:dict[str,str], logger:XML_Logger) -> bool:
    required_keys:dict[str,str] =   {
                                        "Logger_Base_Directory":str,
                                        "Logger_Filename":str,
                                        "Logger_Archive_Folder":str,
                                        "SMTP_SSL_Host":str,
                                        "SMTP_SSL_Port":int,
                                        "Sender_Email":str,
                                        "Sender_Email_Password":str,
                                        "To_Email":str,
                                        "CC_Email":str,
                                        "Warn_User_Of_Logoff":bool,
                                        "Logoff_Warning_Time_Left":int,
                                        "DEBUG":bool
                                    }
    missing_keys:list[str] = []
    for key,item in required_keys.items():
        if(
            (key not in configuration.keys())or
            (not(isinstance(configuration[key],item)))or
            (configuration[key] is None) or
            (configuration[key] == "")
          ):
            missing_keys.append(key)
        if(
            (key in configuration.keys())and
            (key == "Logoff_Warning_Time_Left")and 
            (configuration[key] < 1)
          ):
            missing_keys.append(key)
    if missing_keys:
        logger.log_to_xml(message=f"Missing required keys in configuration: {missing_keys}. Terminating program.",status="CRITICAL",basepath=logger.base_dir)
        return False
    return True

def get_logger(configuration:dict[str,str]) -> XML_Logger:
    logger:XML_Logger = XML_Logger(
                                    log_file=configuration["Logger_Filename"], 
                                    archive_folder=configuration["Logger_Archive_Folder"],
                                    log_retention_days=30,
                                    base_dir=configuration["Logger_Base_Directory"]
                                  )
    return logger

def display_discretion_message() -> None:
    # Display information to the user
    messagebox.showinfo("DISCRETION", """This is the automatic logoff program. It will log the computer off after a pre-designated number of minutes.
If the program is closed at any point while the computer is logged on, the IT director and the administrator will be notified immediately and you will be asked to leave the computer.""")

def get_number_of_user_minutes() -> int:
    # Get user input
    logoff_minutes:str = simpledialog.askstring("TIME", "How many total minutes will be permitted until the computer is logged off?\nMinimum 1 minute. Maximum 120 minutes. ")
    if(logoff_minutes is None):
        return -1
    while(
            (not(logoff_minutes.isnumeric()))
            or 
            (
                (int(logoff_minutes)<=0)or
                (int(logoff_minutes)>120)
            )
        ):
        logoff_minutes:str = simpledialog.askstring("TIME", "How many total minutes will be permitted until the computer is logged off?\nMinimum 1 minute. Maximum 120 minutes.")
        if(logoff_minutes is None):
            return -1
    minutes:int = int(logoff_minutes)
    return minutes

def get_start_and_end_times(minutes:int) -> str:
    start_time:datetime = datetime.now()
    start_hour:int = start_time.hour
    am_pm_start = "A.M." if start_hour < 12 else "P.M."
    if(start_hour>12):
        start_hour -= 12
        start_hour:str = str(start_hour)
    else:
        start_hour:str = str(start_hour)
        
    start_minute:int = start_time.minute
    if(start_minute<10):
        start_minute:str = f"0{str(start_minute)}"
    else:
        start_minute:str = str(start_minute)
        
    end_time = start_time+timedelta(minutes=minutes)
    end_hour = end_time.hour
    am_pm_end = "A.M." if end_hour < 12 else "P.M."
    if(end_hour>12):
        end_hour -= 12
        end_hour:str = str(end_hour)
    else:
        end_hour:str = str(end_hour)
    end_minute = end_time.minute
    if(end_minute<10):
        end_minute:str = f"0{str(end_minute)}"
    else:
        end_minute:str = str(end_minute)
    return end_time,start_hour,start_minute,end_hour,end_minute,f"The login time is {start_hour}:{start_minute} {am_pm_start}\nYou will be logged off at {end_hour}:{end_minute} {am_pm_end}"

def email_receipt(logger:XML_Logger, start_hour:int, start_minute:int, end_hour:int, end_minute:int, configuration:dict[str,str], logging_in:bool) -> None:
    try:
        if(logging_in):
            subject = f"Computer {platform.node()} Logged In"
            body = f"Computer {platform.node()} was logged in at {start_hour}:{start_minute} and should be logged off by {end_hour}:{end_minute}. If it is still on, ask the person to leave as they have violated the computer policy."
        else:
            subject = f"Computer {platform.node()} Log Off"
            body = f"Computer {platform.node()} successfully logged off."
        message = MIMEMultipart('alternative')
        rcpt = [configuration["CC_Email"],configuration["CC_Email"]]
        message['Subject'] = subject
        message['From'] = configuration["Sender_Email"]
        message['To'] = configuration["CC_Email"]
        message['Cc'] = configuration["CC_Email"]
        html_part = MIMEText(body)
        message.attach(html_part)
        with smtplib.SMTP_SSL(configuration["SMTP_SSL_Host"], configuration["SMTP_SSL_Port"]) as server:
            server.login(configuration["Sender_Email"], configuration["Sender_Email_Password"])
            server.sendmail(configuration["Sender_Email"], rcpt, message.as_string())
        logger.log_to_xml(f"Login email successfully sent to {rcpt}.","SUCCESS",basepath=logger.base_dir)
    except Exception as e:
        logger.log_to_xml(f"Email failed to send. Official error: {traceback.format_exc()}","ERROR",basepath=logger.base_dir)

def run_sleep_loop(logger:XML_Logger,end_time:datetime,minutes:int,warning_minutes:int) -> None:
    while datetime.now() < end_time:
        minutes_left_float:float = ((end_time-datetime.now()).total_seconds())/60
        if(minutes_left_float < 1):
            sleep((minutes_left_float*60)+0.01)  # Sleep for the proper number of seconds in the last minute plus one second.
        else:
            sleep(max(60 - (datetime.now().second % 60), 0))  # Align to whole minutes
        minutes_left:int = round(((end_time-datetime.now()).total_seconds())/60)
        logger.log_to_xml(f"{minutes_left:,.0f}/{minutes} minutes remaining","INFO",basepath=logger.base_dir)
        if(minutes_left == warning_minutes):
            messagebox.showwarning("LOGOFF WARNING",f"Logging off in {minutes_left} minutes. Save your progress!")

def logoff_computer():
    return None # Used for testing everything else without logging off and having to wait to log back in
    if platform.system() == "Windows":
        subprocess.run(["shutdown", "-l"], shell=True)
    else:
        subprocess.run(["pkill", "-SIGTERM", "-u", os.getenv("USER")])

def main() -> None:
    configuration:dict[str,str] = get_configuration()
    logger:XML_Logger = get_logger(configuration=configuration)
    if(not(_verify_configuration(configuration=configuration,logger=logger))):
        return
    root:Tk = Tk()
    root.withdraw() # Hide the main window
    if configuration.get("DEBUG", True):
        messagebox.showinfo("DEBUG", f"Logger: {logger}")
    display_discretion_message()
    minutes:int = get_number_of_user_minutes()
    if(minutes == -1):
        return
    end_time,start_hour,start_minute,end_hour,end_minute,start_end_time_message = get_start_and_end_times(minutes)
    messagebox.showinfo("LOGOFF TIME",start_end_time_message)
    logger.log_to_xml(start_end_time_message,basepath=logger.base_dir,status="INFO")
    email_receipt(logger=logger, start_hour=start_hour, start_minute=start_minute, end_hour=end_hour, end_minute=end_minute, logging_in=True, configuration=configuration)
    run_sleep_loop(logger, end_time, minutes, configuration["Logoff_Warning_Time_Left"])
    email_receipt(logger=logger, start_hour=start_hour, start_minute=start_minute, end_hour=end_hour, end_minute=end_minute, logging_in=False, configuration=configuration)
    logoff_computer()

if __name__ == "__main__":
    main()