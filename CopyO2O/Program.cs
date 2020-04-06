using System;
using System.Collections.Generic;
using System.Linq;
using Commonfunctions.Logging;
using System.Diagnostics;
using System.Threading.Tasks;
using System.IO;

namespace CopyO2O
{
    class Program
    {
#if DEBUG
        static bool logOutput = true;
#else
        static bool logOutput = false;
#endif
        public static void Log(string value, bool toConsole = false, bool suppressDateTime = false) { Output.Print(value: (!suppressDateTime ? DateTime.Now.ToString() + " - " : "") + value, logEnabled: logOutput, logFile: logFile); if (((logFile != null) || !logOutput) && toConsole) Console.Write(value); }
        public static void LogLn(string value, bool toConsole = false, bool suppressDateTime = false) { Log(value + "\r\n", toConsole, suppressDateTime); }
        public static string logFile = null;

        static void Main(string[] args)
        {
            const string appId = "cb2c2f84-63f0-49d8-9335-79fcc6050654";
            List<string> AppPermissions = new List<string> { "User.Read", "Calendars.ReadWrite", "Contacts.ReadWrite" };

            DateTime from = DateTime.Now.AddMonths(-1);
            from = from.AddHours(-from.Hour).AddMinutes(-from.Minute).AddSeconds(-from.Second).AddMilliseconds(-from.Millisecond);
            DateTime to = from.AddMonths(2).AddHours(23).AddMinutes(59).AddSeconds(59).AddMilliseconds(999);

            int clearpast = 0;
            string calendar_source_Name = "";
            string calendar_destination_Name = "";
            string contacts_source_Name = "";
            string contacts_destination_Name = "";
            string proxy = "";
            bool clrNotExisting = true; //clear every items of target which does not exist on source side
            bool exitLocalOutlookAfterProcessing = true; //if the app has opened a local Outlook instance exit it at the end

            bool SyncCAL() { return (calendar_source_Name != "") && (calendar_destination_Name != ""); }
            bool SyncCON() { return (contacts_source_Name != "") && (contacts_destination_Name != ""); }

            try
            {
                //iterate through all parameters
                foreach (string arg in args)
                {
                    string parameter = arg.Trim().Split(':')[0].ToUpper();
                    string parValue = arg.Trim().Substring(parameter.Length, arg.Trim().Length - parameter.Length).TrimStart(':');

                    switch (parameter)
                    {
                        case "/CAL":
                            if (parValue[0] == '"')
                            {
                                calendar_source_Name = parValue.Split(';')[0].Trim('"');
                                calendar_destination_Name = parValue.Split(';')[1].Trim('"');
                                //if no dest folder was set use the default one
                                if (calendar_destination_Name == "") calendar_destination_Name = null;
                            }
                            else
                            {
                                calendar_source_Name = parValue.Split(';')[0].Trim('\'');
                                calendar_destination_Name = parValue.Split(';')[1].Trim('\'');
                                //if no dest folder was set use the default one
                                if (calendar_destination_Name == "") calendar_destination_Name = null;
                            }
                            break;
                        case "/CON":
                            if (parValue[0] == '"')
                            {
                                contacts_source_Name = parValue.Split(';')[0].Trim('"');
                                contacts_destination_Name = parValue.Split(';')[1].Trim('"');
                                //if no dest folder was set use the default one
                                if (contacts_destination_Name == "") contacts_destination_Name = null;
                            }
                            else
                            {
                                contacts_source_Name = parValue.Split(';')[0].Trim('\'');
                                contacts_destination_Name = parValue.Split(';')[1].Trim('\'');
                                //if no dest folder was set use the default one
                                if (contacts_destination_Name == "") contacts_destination_Name = null;
                            }
                            break;
                        case "/FROM":
                            if (!DateTime.TryParse(parValue, out from))
                                from = DateTime.Today.AddDays(int.Parse(parValue));
                            break;
                        case "/TO":
                            if (!DateTime.TryParse(parValue, out to))
                                to = DateTime.Today.AddDays(int.Parse(parValue));
                            break;
                        case "/CLEAR":
                            clearpast = Math.Abs(int.Parse(parValue));
                            break;
                        case "/CLR":
                            clearpast = Math.Abs(int.Parse(parValue));
                            break;
                        case "/PROXY":
                            proxy = parValue;
                            break;
                        case "/DNE":
                            exitLocalOutlookAfterProcessing = false;
                            break;
                        case "/LOG":
                            logOutput = true;
                            if (parValue != "")
                                logFile = parValue;
                            break;
                    }
                }

                //check if mandatory parameters were set stop execution
                if (SyncCAL() && (calendar_source_Name.Equals("") || !calendar_source_Name.Contains("\\")))
                    throw new Exception("Source calender path not valid.");

                //check if mandatory parameters were set stop execution
                if (SyncCON() && (contacts_source_Name.Equals("") || !contacts_source_Name.Contains("\\")))
                    throw new Exception("Source contact folder path not valid.");

                //check if from-date is lower than to-date
                if (from >= to)
                    throw new Exception("FROM-date (" + from.ToShortDateString() + ") must be lower than TO-date (" + to.ToShortDateString() + ").");

                //if any sync is set
                if (!SyncCAL() && !SyncCON())
                    throw new Exception("At least one sync config must be set.");
            }
            catch (Exception e)
            {
                LogLn("Error: " + e.Message + "\r\n", true);
                LogLn("Parameters:\r\n"
                    + "/CAL:\"<source>\";\"<destination>\" : Calendar source and destination\r\n"
                    + "/CON:\"<source>\";\"<destination>\" : Contacts source and destination\r\n"
                    + "[opt] /from:<date>              : for calendar: First date to sync (DD.MM.YYYY) or relative to today (in days; eg. -10)\r\n"
                    + "[opt] /to:<date>                : for calendar: Last date to sync (DD.MM.YYYY) or relative to today (in days; eg. 8)\r\n"
                    + "[opt] /clear:<days>             : for calendar: Clear <days> in the past (from 'from' back)\r\n"
                    + "[opt] /proxy:<address>          : set if an explicit proxy should be used for connection\r\n"
                    + "[opt] /DNE                      : if the process has started a local Outlook instance suppress the exit\r\n"
                    + "[opt] /log                      : Verbose logging\r\n\r\n"
                    + "Example: CopyO2O /CAL:\"Hans.Mustermann@company.com\\Calendar\";\"Business\" /from:-7 /to:30 /clear:14", true);
                System.Environment.Exit(-1);
            }

            LogLn("Start copy...", true);

            //load the sync cache
            SyncHelpers.LoadSyncCache();
            List<SyncHelpers.SyncInfo> syncWork = new List<SyncHelpers.SyncInfo>();

            //set proxy if set
            if (proxy != "")
            { System.Net.WebRequest.DefaultWebProxy = new System.Net.WebProxy(proxy, true); }

            Outlook.Application outlookApp = null;
            Office365.Graph office365 = null;
            try
            {
                Log("Open Outlook...");
                outlookApp = new Outlook.Application(exitLocalOutlookAfterProcessing);
                LogLn(" Done.", false, true);

                Log("Connect to Office365...");
                office365 = new Office365.Graph(appId, AppPermissions);
                LogLn(" Done.", false, true);

                //if calendar values should be synced
                if (SyncCAL())
                {
                    LogLn("Calendar: '" + calendar_source_Name + "' >> '" + (calendar_destination_Name ?? "DEFAULT") + "'" + " from " + from.ToShortDateString() + " to " + to.ToShortDateString(), true);
                    Log("... ", true);

                    LogLn("", false, true);
                    Log("Get all events of Outlook...");
                    Outlook.Calendar src_calendar = outlookApp.GetCalendar(calendar_source_Name);
                    Events srcEvents = src_calendar.GetItems(from, to);
                    LogLn(" Done. " + srcEvents.Count.ToString() + " found.", false, true);

                    Log("Get all events of O365...");
                    Office365.Calendar o365_dest_calendar = office365.Calendars[calendar_destination_Name ?? "Calendar"]; //Calendar is the default calendar folder
                    Events destEvents = o365_dest_calendar.GetItemsAsync(from, to).Result;
                    LogLn(" Done. " + destEvents.Count.ToString() + " found.", false, true);

                    LogLn("Analyse sync tasks (NEW, MOD, DEL) ...");
                    SyncHelpers.AnalyseSyncJobs(clrNotExisting, syncWork, srcEvents, destEvents, SyncHelpers.ID_type.CalItem);
                    LogLn(" Done.");

                    LogLn("Create events in O365...");
                    int countOfcreateTasks = SyncHelpers.CreateItems(syncWork, srcEvents, o365_dest_calendar, SyncHelpers.ID_type.CalItem);
                    LogLn(" Done. (" + countOfcreateTasks + ")");

                    LogLn("Delete events in O365...");
                    int countOfdeleteTasks = SyncHelpers.RemoveItems(syncWork, o365_dest_calendar, SyncHelpers.ID_type.CalItem);
                    LogLn(" Done. (" + countOfdeleteTasks + ")");

                    //exec only if verbose logging enabled
                    if (logOutput)
                    {
                        Log("Validate count of events in O365...");
                        LogLn(" Done. " + o365_dest_calendar.GetItemsAsync(from.AddDays(-clearpast), to).Result.Count.ToString() + " found.", false, true);
                    }
                    Log("");
                    LogLn(" Done (" + srcEvents.Count.ToString() + "/" + (destEvents.Count + countOfcreateTasks - countOfdeleteTasks).ToString() + ")", true, true);
                }

                //if contacts should be synced
                if (SyncCON())
                {
                    LogLn("Contacts: '" + contacts_source_Name + "' >> '" + (contacts_destination_Name ?? "DEFAULT") + "'", true);
                    Log("... ", true);

                    LogLn("", false, true);
                    Log("Get all contacts of Outlook...");
                    Outlook.ContactFolder src_contactfolder = outlookApp.GetContactFolder(contacts_source_Name);
                    ContactCollectionType srcContacts = src_contactfolder.GetItems();
                    LogLn(" Done. " + srcContacts.Count.ToString() + " found.", false, true);

                    Log("Get all contacts of O365...");
                    Office365.ContactFolder o365_dest_contactfolder = office365.ContactFolders[contacts_destination_Name];
                    ContactCollectionType destContacts = o365_dest_contactfolder.GetContactsAsync().Result;
                    LogLn(" Done. " + destContacts.Count.ToString() + " found.", false, true);

                    LogLn("Analyse sync tasks (NEW, MOD, DEL) ...");
                    SyncHelpers.AnalyseSyncJobs(clrNotExisting, syncWork, srcContacts, destContacts, SyncHelpers.ID_type.Contact);
                    LogLn(" Done.");

                    LogLn("Create contacts in O365...");
                    int countOfcreateTasks = SyncHelpers.CreateItems(syncWork, srcContacts, o365_dest_contactfolder, SyncHelpers.ID_type.Contact);
                    LogLn(" Done. (" + countOfcreateTasks + ")");

                    LogLn("Delete contacts in O365...");
                    int countOfdeleteTasks = SyncHelpers.RemoveItems(syncWork, o365_dest_contactfolder, SyncHelpers.ID_type.Contact);
                    LogLn(" Done. (" + countOfdeleteTasks + ")");

                    //exec only if verbose logging enabled
                    if (logOutput)
                    {
                        Log("Validate count of contacts in O365...");
                        LogLn(" Done. " + o365_dest_contactfolder.GetContactsAsync().Result.Count.ToString() + " found.", false, true);
                    }
                    Log("");
                    LogLn(" Done (" + srcContacts.Count.ToString() + "/" + (destContacts.Count + countOfcreateTasks - countOfdeleteTasks).ToString() + ")", true, true);
                }
            }
            catch (AggregateException ae)
            {
                ae.Handle((e) => { LogLn(" Error occured: " + e.Message, true); return true; });
            }
            catch (System.Net.WebException e)
            {
                LogLn(" Error occured: " + e.Message + "\r\nConnection could not be establised! Proxy?", true);
            }
            catch (Exception e)
            {
                if (e.InnerException != null)
                    LogLn(" Error occured: " + e.InnerException.Message, true);
                else
                    LogLn(" Error occured: " + e.Message, true);
            }
            finally
            {
                //if outlook is still open
                if (outlookApp != null)
                {
                    Log("Close Outlook...");
                    outlookApp.Quit();
                    LogLn(" Done.", false, true);
                }

                //if a connection to office365 is still established
                if (office365 != null)
                {
                    Log("Disconnect Office365...");
                    //office365.Flush(); //finalize all open commands
                    //office365.Disconnect(); //close connection
                    office365 = null;
                    LogLn(" Done.", false, true);
                }

                //store the sync cache to disk
                SyncHelpers.StoreSyncCache();
            }

            LogLn(syncWork.Count.ToString() + " operations processed.", true);

#if DEBUG
            Console.ReadLine();
#endif
        }
    }
}
