using System;
using System.Collections.Generic;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace CopyO2O
{
    class Program
    {
        public class ListOfEvents : List<Outlook.AppointmentItem> { };
        static bool logOutput = false;

        public static void OutputLog(string value, bool linebreak = false)
        {
            if (logOutput)
            {
                if (linebreak)
                    Console.WriteLine(value);
                else
                    Console.Write(value);
            }
        }

        static void Main(string[] args)
        {
            DateTime from = DateTime.Now.AddMonths(-1);
            DateTime to = DateTime.Now.AddMonths(1);
            string calenderName_source = "";
            string calenderName_destination = "";
            bool outlookAlreadyRunning = false;

            try
            {
                //iterate through all parameters
                foreach (string arg in args)
                {
                    string parameter = arg.Trim().Split(':')[0].ToUpper();
                    string parValue = arg.Trim().Substring(parameter.Length, arg.Trim().Length - parameter.Length).TrimStart(':');

                    switch (parameter)
                    {
                        case "/SRC": calenderName_source = parValue; break;
                        case "/DEST": calenderName_destination = parValue; break;
                        case "/FROM":
                            if (!DateTime.TryParse(parValue, out from))
                                from = DateTime.Today.AddDays(int.Parse(parValue));
                            break;
                        case "/TO":
                            if (!DateTime.TryParse(parValue, out to))
                                to = DateTime.Today.AddDays(int.Parse(parValue));
                            to = to.AddHours(23).AddMinutes(59).AddSeconds(59).AddMilliseconds(999);
                            break;
                        case "/LOG": logOutput = true; break;
                    }
                }

                //check if mandatory parameters were set stop execution
                if (calenderName_source.Equals("") || !calenderName_source.Contains("\\") || calenderName_destination.Equals("") || !calenderName_destination.Contains("\\"))
                    throw new Exception("Calender paths not valid.");

                //check if from-date is lower than to-date
                if (from >= to)
                    throw new Exception("FROM-date (" + from.ToShortDateString() + ") must be lower than TO-date (" + to.ToShortDateString() + ").");
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: " + e.Message);
                Console.WriteLine("Parameters:\n"
                    + "/src:<string>           : Name and path of the source calendar\n"
                    + "/dest:<string>          : Name and path of the destination calendar\n"
                    + "/from:<date>            : First date to sync (DD.MM.YYYY) or relative to today (in days; eg. -10)\n"
                    + "/to:<date>              : Last date to sync (DD.MM.YYYY) or relative to today (in days; eg. 8)\n"
                    + "/log                    : Verbose logging");
                System.Environment.Exit(-1);
            }

            string overview = "Start sync of "
                + calenderName_source + " >> " + calenderName_destination
                + " from " + from.ToShortDateString() + " to " + to.ToShortDateString();
            Console.WriteLine(overview);

            OutputLog("Open Outlook...", false);
            Outlook.Application outlookApp;
            try
            {
                outlookApp = (Outlook.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Outlook.Application");
                outlookAlreadyRunning = true;
            }
            catch { outlookApp = new Outlook.Application(); }

            Outlook.MAPIFolder src_calendar = GetCalendar(outlookApp, calenderName_source);
            Outlook.MAPIFolder dest_calendar = GetCalendar(outlookApp, calenderName_destination);
            OutputLog(" Done.", true);

            OutputLog("Get all source events...", false);
            ListOfEvents srcEvents = GetCalendarItems(src_calendar, from, to);
            OutputLog(" Done. " + srcEvents.Count.ToString() + " found.", true);

            ListOfEvents destEvents = null;
            try
            {
                DeleteCalendarItems(dest_calendar, from, to);
                CreateCalendarItems(dest_calendar, srcEvents);

                OutputLog("Get all destination events...", false);
                destEvents = GetCalendarItems(dest_calendar, from, to);
                OutputLog(" Done. " + destEvents.Count.ToString() + " found.", true);
            }
            finally
            {
                outlookApp.Session.SendAndReceive(false);

                //if outlook was not already running
                if (!outlookAlreadyRunning)
                {
                    OutputLog("Close Outlook...");
                    outlookApp.Quit();

                    outlookApp = null;
                    src_calendar = null;
                    dest_calendar = null;
                    srcEvents = null;
                    destEvents = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    OutputLog(" Done.", true);
                }
            }

            OutputLog("End.", true);
            Console.ReadLine();
            }

        public static ListOfEvents GetCalendarItems(Outlook.MAPIFolder calendar, DateTime from, DateTime to)
        {
            Outlook.Items tempEvents;
            tempEvents = calendar.Items;
            tempEvents.IncludeRecurrences = true;
            tempEvents.Sort("[Start]");

            string filter = "[Start] >= '" + from.ToString("g") + "'"
                + " AND " + "[Start] <= '" + to.ToString("g") + "'";
            Outlook.Items eventsFiltered = tempEvents.Restrict(filter);

            ListOfEvents result = new ListOfEvents();
            foreach (Outlook.AppointmentItem aptmItem in eventsFiltered)
            {
                result.Add(aptmItem);
            }

            return result;
        }

        public static void CreateCalendarItems(Outlook.MAPIFolder calendar, ListOfEvents newItems)
        {
            foreach (Outlook.AppointmentItem item in newItems)
            {
                Outlook.AppointmentItem newEvent = calendar.Items.Add();
                newEvent.StartTimeZone = item.StartTimeZone;
                newEvent.StartUTC = item.StartUTC;
                newEvent.EndTimeZone = item.EndTimeZone;
                newEvent.EndUTC = item.EndUTC;

                newEvent.Subject = item.Subject;
                newEvent.Location = item.Location;
                newEvent.AllDayEvent = item.AllDayEvent;

                newEvent.ReminderMinutesBeforeStart = item.ReminderMinutesBeforeStart;
                newEvent.ReminderSet = item.ReminderSet;
                newEvent.BusyStatus = item.BusyStatus;

                newEvent.Save();
            }
        }

        public static void DeleteCalendarItems(Outlook.MAPIFolder calendar, DateTime from, DateTime to)
        {
            Outlook.Items tempEvents;
            tempEvents = calendar.Items;
            tempEvents.IncludeRecurrences = true;
            tempEvents.Sort("[Start]");

            string filter = "[Start] >= '" + from.ToString("g") + "'"
                + " AND " + "[Start] <= '" + to.ToString("g") + "'";
            Outlook.Items eventsFiltered = tempEvents.Restrict(filter);

            int count = 0;
            foreach (Outlook.AppointmentItem tmp in eventsFiltered) count++;

            for (int index = count; index >= 1; index--)
            {
                tempEvents[index].Delete();
            }     
        }

        public static void DeleteCalendarItems(Outlook.MAPIFolder calendar)
        {
            for (int index = calendar.Items.Count; index >= 1; index--)
            {
                calendar.Items[index].delete();
            }
        }


        /// <summary>
        /// Resolves the given calendar path and return the corresponding MAPI folder in a local outlook instance
        /// </summary>
        /// <param name="app">local Outlook application</param>
        /// <param name="calendarPath">path of calendar (Profil\Folder\Subfolder\...)</param>
        /// <returns>returns a MAPI folder (calendar)</returns>
        /// 
        public static Outlook.MAPIFolder GetCalendar(Outlook.Application app, string calendarPath)
        {
            string[] names = calendarPath.Split('\\');
            Outlook.MAPIFolder result = app.Session.Folders[names[0]];

            foreach (string name in names.Skip(1))
            {
                result = result.Folders[name];
            }

            return result;
        }
    }
}
