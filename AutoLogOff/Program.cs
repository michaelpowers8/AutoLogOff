using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Diagnostics;
using System.Windows.Forms;
using System.Threading.Tasks;
using NLog;
using Newtonsoft.Json;

class Program
{
    private static readonly Logger logger = LogManager.GetCurrentClassLogger();

    static void Main()
    {
        try
        {
            var configuration = GetConfiguration();
            if (configuration == null || !VerifyConfiguration(configuration))
            {
                logger.Error("Invalid configuration. Program terminating.");
                return;
            }

            if (configuration.DEBUG)
            {
                MessageBox.Show($"Logger initialized successfully", "DEBUG");
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            DisplayDiscretionMessage();

            var minutes = GetNumberOfUserMinutes();
            if (minutes == -1) return;

            var (endTime, startHour, startMinute, endHour, endMinute, message) = GetStartAndEndTimes(minutes);

            MessageBox.Show(message, "LOGOFF TIME");
            logger.Info(message);

            EmailReceipt(startHour, startMinute, endHour, endMinute, configuration, true).Wait();

            RunSleepLoop(endTime, minutes).Wait();

            EmailReceipt(startHour, startMinute, endHour, endMinute, configuration, false).Wait();

            LogoffComputer();
        }
        catch (Exception ex)
        {
            logger.Error(ex, "Unhandled exception in main program");
        }
    }

    private static Configuration GetConfiguration()
    {
        const string configPath = "Config.json";
        try
        {
            if (File.Exists(configPath))
            {
                var json = File.ReadAllText(configPath);
                return JsonConvert.DeserializeObject<Configuration>(json);
            }
            return null;
        }
        catch (Exception ex)
        {
            logger.Error(ex, "Failed to load configuration");
            return null;
        }
    }

    private static bool VerifyConfiguration(Configuration config)
    {
        if (string.IsNullOrEmpty(config.Logger_Base_Directory)) return false;
        if (string.IsNullOrEmpty(config.Logger_Filename)) return false;
        if (string.IsNullOrEmpty(config.Logger_Archive_Folder)) return false;
        if (string.IsNullOrEmpty(config.SMTP_SSL_Host)) return false;
        if (config.SMTP_SSL_Port <= 0) return false;
        if (string.IsNullOrEmpty(config.Sender_Email)) return false;
        if (string.IsNullOrEmpty(config.Sender_Email_Password)) return false;
        if (string.IsNullOrEmpty(config.To_Email)) return false;
        if (string.IsNullOrEmpty(config.CC_Email)) return false;

        return true;
    }

    private static void DisplayDiscretionMessage()
    {
        MessageBox.Show(
            "This is the automatic logoff program. It will log the computer off after a pre-designated number of minutes.\n" +
            "If the program is closed at any point while the computer is logged on, the IT director and the administrator " +
            "will be notified immediately and you will be asked to leave the computer.",
            "DISCRETION");
    }

    private static int GetNumberOfUserMinutes()
    {
        using (var form = new Form())
        {
            int minutes = -1;
            bool validInput = false;

            while (!validInput)
            {
                var result = Microsoft.VisualBasic.Interaction.InputBox(
                    "How many total minutes will be permitted until the computer is logged off?\n" +
                    "Minimum 1 minute. Maximum 120 minutes.",
                    "TIME", "60");

                if (string.IsNullOrEmpty(result)) return -1;

                if (int.TryParse(result, out minutes) && minutes >= 1 && minutes <= 120)
                {
                    validInput = true;
                }
            }

            return minutes;
        }
    }

    private static (DateTime endTime, string startHour, string startMinute, string endHour, string endMinute, string message)
        GetStartAndEndTimes(int minutes)
    {
        var startTime = DateTime.Now;

        var startHour = startTime.Hour > 12 ? (startTime.Hour - 12).ToString() : startTime.Hour.ToString();
        var amPmStart = startTime.Hour >= 12 ? "P.M." : "A.M.";
        var startMinute = startTime.Minute < 10 ? $"0{startTime.Minute}" : startTime.Minute.ToString();

        var endTime = startTime.AddMinutes(minutes);
        var endHour = endTime.Hour > 12 ? (endTime.Hour - 12).ToString() : endTime.Hour.ToString();
        var amPmEnd = endTime.Hour >= 12 ? "P.M." : "A.M.";
        var endMinute = endTime.Minute < 10 ? $"0{endTime.Minute}" : endTime.Minute.ToString();

        var message = $"The login time is {startHour}:{startMinute} {amPmStart}\n" +
                     $"You will be logged off at {endHour}:{endMinute} {amPmEnd}";

        return (endTime, startHour, startMinute, endHour, endMinute, message);
    }

    private static async Task EmailReceipt(string startHour, string startMinute, string endHour, string endMinute,
        Configuration config, bool loggingIn)
    {
        try
        {
            var subject = loggingIn
                ? $"Computer {Environment.MachineName} Logged In"
                : $"Computer {Environment.MachineName} Log Off";

            var body = loggingIn
                ? $"Computer {Environment.MachineName} was logged in at {startHour}:{startMinute} and should be logged off by {endHour}:{endMinute}. " +
                  "If it is still on, ask the person to leave as they have violated the computer policy."
                : $"Computer {Environment.MachineName} successfully logged off.";

            var message = new MailMessage
            {
                From = new MailAddress(config.Sender_Email),
                Subject = subject,
                Body = body,
                IsBodyHtml = false
            };

            message.To.Add(config.To_Email);
            message.CC.Add(config.CC_Email);

            using (var client = new SmtpClient(config.SMTP_SSL_Host, config.SMTP_SSL_Port))
            {
                client.EnableSsl = true;
                client.Credentials = new NetworkCredential(config.Sender_Email, config.Sender_Email_Password);

                await client.SendMailAsync(message);
                logger.Info($"Email successfully sent to {config.To_Email} and {config.CC_Email}");
            }
        }
        catch (Exception ex)
        {
            logger.Error(ex, "Failed to send email");
        }
    }

    private static async Task RunSleepLoop(DateTime endTime, int totalMinutes)
    {
        while (DateTime.Now < endTime)
        {
            var timeLeft = endTime - DateTime.Now;
            var minutesLeft = (int)timeLeft.TotalMinutes;

            logger.Info($"{minutesLeft}/{totalMinutes} minutes remaining");

            if (minutesLeft == 1 || minutesLeft == 5)
            {
                MessageBox.Show($"Logging off in {minutesLeft} minutes. Save your progress!", "LOGOFF WARNING",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            // Wait until the next whole minute
            var delay = (int)(TimeSpan.FromMinutes(1) - TimeSpan.FromSeconds(DateTime.Now.Second)).TotalMilliseconds;
            await Task.Delay(Math.Max(delay, 0));
        }
    }

    private static void LogoffComputer()
    {
        try
        {
            if (Environment.OSVersion.Platform == PlatformID.Win32NT)
            {
                //Process.Start("shutdown", "/l");
            }
            else
            {
                //Process.Start("pkill", $"-SIGTERM -u {Environment.UserName}");
            }
        }
        catch (Exception ex)
        {
            logger.Error(ex, "Failed to log off computer");
        }
    }
}

public class Configuration
{
    public string Logger_Base_Directory { get; set; }
    public string Logger_Filename { get; set; }
    public string Logger_Archive_Folder { get; set; }
    public string SMTP_SSL_Host { get; set; }
    public int SMTP_SSL_Port { get; set; }
    public string Sender_Email { get; set; }
    public string Sender_Email_Password { get; set; }
    public string To_Email { get; set; }
    public string CC_Email { get; set; }
    public bool DEBUG { get; set; }
}