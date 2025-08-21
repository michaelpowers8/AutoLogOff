import os
import json
import smtplib
import platform
import traceback
import subprocess
from time import sleep
from xml_logging import XML_Logger
from email.mime.text import MIMEText
from cryptography.fernet import Fernet
from datetime import datetime,timedelta
from email.mime.multipart import MIMEMultipart
from tkinter import Tk, messagebox, simpledialog, Toplevel, Label 

def block_alt_f4(event):
    return 'break'

def get_configuration() -> dict[str,str|bool|int]|None:
    """
    Get the configuration from corresponding Config.json.

    Returns a configuration made as a dictionary.
    """
    try:
        # Load encryption key
        with open("secret.key", "rb") as key_file:
            key = key_file.read()

        fernet = Fernet(key)

        # Decrypt the config file
        with open("Config.encrypted", "rb") as encrypted_file:
            decrypted_data:bytes = fernet.decrypt(encrypted_file.read())
        
        configuration = json.loads(decrypted_data.decode(errors="ignore"))
        return configuration
    except Exception as e:
        traceback.print_exc()
        return None
    
def get_configuration_without_encryption() -> dict[str,str|bool|int]|None:
    """
    Get the configuration from corresponding Config.json.

    Returns a configuration made as a dictionary.
    """
    try:
        # Load encryption key
        with open("Config.json", "rb") as key_file:
            return json.loads(key_file.read())
    except Exception as e:
        traceback.print_exc()
        return None

def _verify_configuration(configuration:dict[str,str|bool|int], logger:XML_Logger) -> bool:
    if configuration is None:
        return False
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

def get_logger(configuration:dict[str,str|bool|int]) -> XML_Logger:
    """Get the logger that will be used to log information and errors to developer for future debugging."""
    logger:XML_Logger = XML_Logger(
                                    log_file=configuration["Logger_Filename"], 
                                    archive_folder=configuration["Logger_Archive_Folder"],
                                    log_retention_days=30,
                                    base_dir=configuration["Logger_Base_Directory"]
                                  )
    return logger

def display_discretion_message() -> None:
    """
    Display a warning message to the user that this is an automated program to shut the computer off after a set period of time.
    Terminating the program early will result in computer privileges revoked.
    """
    # Display information to the user
    messagebox.showinfo("DISCRETION", """This is the automatic logoff program. It will log the computer off after a pre-designated number of minutes.
If the program is closed at any point while the computer is logged on, the IT director and the administrator will be notified immediately and you will be asked to leave the computer.""")

def get_number_of_user_minutes() -> int:
    """
    Displays a message for management asking how many total minutes the user will be permitted to use the computer.
    Number of minutes must be an integer between 1-120 inclusive. If bad input is entered, a new window with the 
    same message will pop up.

    Returns:
        >>> int representing number of total minutes the user will be permitted to use the computer before it is forcefully logged off.
    """
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

def get_start_and_end_times(minutes:int) -> tuple[datetime,int,int,int,int,str]:
    """
    Get the exact date and time the user's computer time will start and end down to the microsecond. Will use these times to inform the user of their
    logoff time, and maintain proper computations and accurate timing.

    Arguments
    ---------
        int minutes representing the total number of minutes the user will be allowed to be logged into the computer.
    
    Returns:
    --------
        datetime end_time for the exact time the logoff process will begin
        int start_hour for the hour number that will be displayed to the user
        int start_minute for the minute number that will be displayed to the user
        int end_hour for the hour number that will be displayed to the user
        int end_minute for the minute number that will be displayed to the user
        str representing the full message that the user will see informing them of the exact time their login time is and when their log off time will be
    """
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

def email_receipt(logger:XML_Logger, start_hour:int, start_minute:int, end_hour:int, end_minute:int, configuration:dict[str,str|bool|int], logging_in:bool) -> None:
    """
    Email To and CC any management or developers about when a user logs on or off the computer. This is meant as an external record in case the logs on the 
    computer are corrupted in any way. 

    Arguments
    --------
        logger : XML_Logger to log any error that occurs while trying to send the email out, or log when an email is successfully sent
        start_hour : int to inform the hour the computer was logged in
        start_minute : int to inform the minute the computer was logged in
        end_hour : int to inform the hour the computer was logged off
        end_minute : int to inform the minute the computer was logged off

    Returns:
    --------
        None
    """
    try:
        if(logging_in):
            subject = f"Computer {platform.node()} Logged In"
            body = f"Computer {platform.node()} was logged in at {start_hour}:{start_minute} and should be logged off by {end_hour}:{end_minute}. If it is still on, ask the person to leave as they have violated the computer policy."
        else:
            subject = f"Computer {platform.node()} Log Off"
            body = f"Computer {platform.node()} successfully logged off."
        message = MIMEMultipart('alternative')
        rcpt = [configuration["To_Email"],configuration["CC_Email"]]
        message['Subject'] = subject
        message['From'] = configuration["Sender_Email"]
        message['To'] = configuration["To_Email"]
        message['Cc'] = configuration["CC_Email"]
        html_part = MIMEText(body)
        message.attach(html_part)
        with smtplib.SMTP_SSL(configuration["SMTP_SSL_Host"], configuration["SMTP_SSL_Port"]) as server:
            server.login(configuration["Sender_Email"], configuration["Sender_Email_Password"])
            server.sendmail(configuration["Sender_Email"], rcpt, message.as_string())
        logger.log_to_xml(message=f"Login email successfully sent to {rcpt}.",status="SUCCESS",basepath=logger.base_dir)
    except Exception as e:
        logger.log_to_xml(message=f"Email failed to send. Official error: {traceback.format_exc()}",status="ERROR",basepath=logger.base_dir)

def non_blocking_warning(root: Tk, title: str, message: str, duration_ms: int = 10000) -> None:
    """
    Show a non-modal, top-most notification window that auto-destroys after `duration_ms`.
    This does not block the main thread (provided run_sleep_loop calls root.update()
    periodically so Tk event callbacks run).
    """
    try:
        popup = Toplevel(root)
        popup.title(title)
        popup.overrideredirect(True)            # remove window decorations (no close button)
        popup.attributes("-topmost", True)      # stay on top
        popup.resizable(False, False)

        # fixed size + position bottom-right-ish
        width, height = 340, 90
        screen_w = root.winfo_screenwidth()
        screen_h = root.winfo_screenheight()
        x = screen_w - width - 12
        y = screen_h - height - 80
        popup.geometry(f"{width}x{height}+{x}+{y}")

        lbl = Label(popup, text=message, justify="left", anchor="w",
                    padx=10, pady=10, wraplength=width-20)
        lbl.pack(expand=True, fill="both")

        # schedule auto-destroy
        popup.after(duration_ms, popup.destroy)
    except Exception:
        # swallow UI errors but log them if logger is available
        # (don't raise — notifications are best-effort)
        pass

def run_sleep_loop(logger: XML_Logger, end_time: datetime, minutes: int, configuration: dict[str, str | bool | int], root: Tk) -> None:
    """
    Silently keep the program running in the background checking every 60 seconds how many minutes are left and logging the information. 
    Logging every minute helps us keep track if a user willingly logs off early, the program will terminate, but the logs will inform
    management that no rules were violated since the logs will inform of the early log off. Also, inform the user when their time is almost
    out.

    Arguments
    ---------
        logger : XML_Logger to log every minute that passes to keep tabs in case user logs off early
        end_time : datetime to terminate the loop when the current date surpasses the end time
        minutes : int representing the total number of minutes the user has access to the computer
        configuration : dictionary of how the program is meant to run. Configurations used in this are to decide if the user is warned about low time, and how many minutes the user has left before being logged off
    """
    while datetime.now() < end_time:
        seconds_left = int((end_time - datetime.now()).total_seconds())
        minutes_left = round(seconds_left / 60)

        logger.log_to_xml(message=f"{minutes_left:,.0f}/{minutes} minutes remaining",
                          status="INFO", basepath=logger.base_dir)

        # conditional extra warning when configured
        if (configuration.get("Warn_User_Of_Logoff") and
                (minutes_left == configuration.get("Logoff_Warning_Time_Left"))):
            non_blocking_warning(root, "LOGOFF WARNING",
                                 f"Logging off in {minutes_left} minute(s). Save your progress!",
                                 duration_ms=10000)

        # Wait up to the next minute but keep Tk events flowing.
        # Sleep in 1-second increments and call root.update() so that popup.after works.
        # If there's less than 60 seconds left, loop only for that many seconds.
        wait_seconds = min(60, max(1, seconds_left))
        for _ in range(wait_seconds):
            sleep(1)
            try:
                root.update()
            except Exception:
                # If root has been destroyed, just continue; we still need to reach end_time.
                pass

def logoff_computer(debugging:bool):
    """
    Dynamic system to log off the computer whether it is windows, Linux or Mac.
    """
    if debugging:
        return None # Used for testing everything else without logging off and having to wait to log back in
    if platform.system() == "Windows":
        subprocess.run(["shutdown", "-l"], shell=True)
    else:
        subprocess.run(["pkill", "-SIGTERM", "-u", os.getenv("USER")])

def main() -> None:
    configuration:dict[str,str|bool|int] = get_configuration_without_encryption()
    logger:XML_Logger = get_logger(configuration=configuration)
    if(not(_verify_configuration(configuration=configuration,logger=logger))):
        return
    root:Tk = Tk()
    root.withdraw() # Hide the main window
    root.bind('<Alt-F4>', block_alt_f4)
    root.protocol("WM_DELETE_WINDOW", lambda: None)  # Disable the close button
    if configuration.get("DEBUG", True):
        messagebox.showinfo("DEBUG", f"Logger: {logger}")
    display_discretion_message()
    minutes:int = get_number_of_user_minutes()
    if(minutes == -1):
        return
    end_time,start_hour,start_minute,end_hour,end_minute,start_end_time_message = get_start_and_end_times(minutes)
    non_blocking_warning(root, "LOGOFF TIME", start_end_time_message, duration_ms=10000)
    logger.log_to_xml(start_end_time_message,basepath=logger.base_dir,status="INFO")
    email_receipt(logger=logger, start_hour=start_hour, start_minute=start_minute, end_hour=end_hour, end_minute=end_minute, logging_in=True, configuration=configuration)
    run_sleep_loop(logger, end_time, minutes, configuration, root)
    email_receipt(logger=logger, start_hour=start_hour, start_minute=start_minute, end_hour=end_hour, end_minute=end_minute, logging_in=False, configuration=configuration)
    logoff_computer(configuration["DEBUG"])

if __name__ == "__main__":
    main()