using System;
using System.Collections.Generic;
using System.Linq;
using Commonfunctions.Logging;
using System.Diagnostics;

namespace CopyO2O
{
    class Program
    {
#if DEBUG
        static bool logOutput = true;
#else
        static bool logOutput = false;
#endif
        static void Log(string value) { Output.Print(value: value, logEnabled: logOutput); }
        static void LogLn(string value) { Output.PrintLn(value: value, logEnabled: logOutput); }

        static void Main(string[] args)
        {
            const string appId = "cb2c2f84-63f0-49d8-9335-79fcc6050654";
            List<string> AppPermissions = new List<string> { "User.Read", "Calendars.ReadWrite", "Contacts.ReadWrite" };
            string tmpIdsToDelete_FilePath = System.Reflection.Assembly.GetExecutingAssembly().Location + "itd.cache";

            DateTime from = DateTime.Now.AddMonths(-1);
            DateTime to = DateTime.Now.AddMonths(1);
            int clearpast = 0;
            string calendar_source_Name = "";
            string calendar_destination_Name = "";
            string contacts_source_Name = "";
            string contacts_destination_Name = "";
            string proxy = "";

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
                            to = to.AddHours(23).AddMinutes(59).AddSeconds(59).AddMilliseconds(999);
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
                        case "/LOG": logOutput = true; break;
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
                Console.WriteLine("Error: " + e.Message + "\n");
                Console.WriteLine("Parameters:\n"
                    + "/CAL:\"<source>\";\"<destination>\" : Calendar source and destination\n"
                    + "/CON:\"<source>\";\"<destination>\" : Contacts source and destination\n"
                    + "[opt] /from:<date>              : for calendar: First date to sync (DD.MM.YYYY) or relative to today (in days; eg. -10)\n"
                    + "[opt] /to:<date>                : for calendar: Last date to sync (DD.MM.YYYY) or relative to today (in days; eg. 8)\n"
                    + "[opt] /clear:<days>             : for calendar: Clear <days> in the past (from 'from' back)\n"
                    + "[opt] /proxy:<address>          : set if an explicit proxy should be used for connection\n"
                    + "[opt] /log                      : Verbose logging\n\n"
                    + "Example: CopyO2O /CAL:\"Hans.Mustermann@company.com\\Calendar\";\"Business\" /from:-7 /to:30 /clear:14");
                System.Environment.Exit(-1);
            }

            string overview = "Start sync of \n";
            if (SyncCAL())
            {
                overview += "Calendar: '" + calendar_source_Name + "' >> '" + (calendar_destination_Name??"DEFAULT") + "'"
                + " from " + from.ToShortDateString() + " to " + to.ToShortDateString() + "\n";
            }
            if (SyncCON())
            {
                overview += "Contacts: '" + contacts_source_Name + "' >> '" + (contacts_destination_Name??"DEFAULT") + "'";
            }
            Console.WriteLine(overview);

            Outlook.Application outlookApp = null;
            Office365.Calendars outlookCloud_Cals = null;
            Office365.ContactFolders outlookCloud_ConFolds = null;
            try
            {
                Log("Open Outlook...");
                outlookApp = new Outlook.Application();
                LogLn(" Done.");

                //set proxy if necessary
                if (proxy != "")
                { System.Net.WebRequest.DefaultWebProxy = new System.Net.WebProxy(proxy, true); }

                //if calendar values should be synced
                if (SyncCAL())
                {
                    Log("Open Office365...");
                    outlookCloud_Cals = new Office365.Calendars(appId, AppPermissions);
                    Office365.Calendar o365_dest_calendar = outlookCloud_Cals.GetCalendar(calendar_destination_Name??"Calendar"); //Calendar is the default calendar folder
                    LogLn(" Done.");

                    Log("Get all source events...");
                    Outlook.Calendar src_calendar = outlookApp.GetCalendar(calendar_source_Name);
                    Events srcEvents = src_calendar.GetItems(from, to);
                    LogLn(" Done. " + srcEvents.Count.ToString() + " found.");

                    Log("Mark calendar items to delete...");
                    List<KeyValuePair<string, string>> calendarItemsToDelete = new List<KeyValuePair<string, string>>();
                    o365_dest_calendar.GetItemsAsync(from.AddDays(-clearpast), to).Result.ForEach((x) => { calendarItemsToDelete.Add(new KeyValuePair<string, string>(Helpers.Marker_O365, x.Id)); });
                    Helpers.SetIDsToRemove(calendarItemsToDelete, Helpers.ID_type.CalItem, tmpIdsToDelete_FilePath);
                    LogLn(" Done.");

                    Log("Copy events...");
                    List<Office365.Replacement> replacements = new List<Office365.Replacement>(); //new List<Office365.Replacement> { new Office365.Replacement { key = "Location", regex = ".*(Vertex).*", newvalue = "Fasanenweg 9, 70771 Leinfelden-Echterdingen, Deutschland" } };
                    o365_dest_calendar.CreateItems(srcEvents, replacements);
                    LogLn(" Done.");

                    Log("Delete doublettes...");
                    List<string> calendarItemsFromFileToDelete = Helpers.GetIDsToRemove(Helpers.ID_type.CalItem, tmpIdsToDelete_FilePath).Where(x => (x.Key == Helpers.Marker_O365)).Select(x => x.Value).ToList<string>();

                    //if doublettes successfully deleted remove cache file
                    if (o365_dest_calendar.DeleteItems(calendarItemsFromFileToDelete) == true)
                        System.IO.File.Delete(tmpIdsToDelete_FilePath);
                    LogLn(" Done.");

                    //exec only if verbose logging enabled
                    if (logOutput)
                    {
                        Log("Get count of created events...");
                        int destCount = o365_dest_calendar.GetItemsAsync(from.AddDays(-clearpast), to).Result.Count;
                        LogLn(" Done. " + destCount.ToString() + " found.");
                    }
                }

                //if contacts should be synced
                if (SyncCON())
                {
                    Log("Open Office365...");
                    outlookCloud_ConFolds = new Office365.ContactFolders(appId, AppPermissions);
                    Office365.ContactFolder o365_dest_contactfolder = outlookCloud_ConFolds.GetContactFolder(contacts_destination_Name);
                    LogLn(" Done.");

                    Log("Get all source contacts...");
                    Outlook.ContactFolder src_contactfolder = outlookApp.GetContactFolder(contacts_source_Name);
                    ContactCollectionType srcContacts = src_contactfolder.GetItems();
                    LogLn(" Done. " + srcContacts.Count.ToString() + " found.");

                    Log("Mark contacts to delete...");
                    List<KeyValuePair<string, string>> idsToDelete = new List<KeyValuePair<string, string>>();
                    o365_dest_contactfolder.GetContactsAsync().Result.ForEach((x) => { idsToDelete.Add(new KeyValuePair<string, string>(Helpers.Marker_O365, x.Id)); });
                    Helpers.SetIDsToRemove(idsToDelete, Helpers.ID_type.Contact, tmpIdsToDelete_FilePath);
                    LogLn(" Done.");

                    Log("Copy contacts...");
                    o365_dest_contactfolder.AddContacts(srcContacts);
                    LogLn(" Done.");

                    Log("Delete doublettes...");
                    List<string> idsFromFileToDelete = Helpers.GetIDsToRemove(Helpers.ID_type.Contact, tmpIdsToDelete_FilePath).Where(x => (x.Key == Helpers.Marker_O365)).Select(x => x.Value).ToList<string>();

                    //if doublettes successfully deleted remove cache file
                    if (o365_dest_contactfolder.Delete(idsFromFileToDelete) == true)
                        System.IO.File.Delete(tmpIdsToDelete_FilePath);
                    LogLn(" Done.");

                    //exec only if verbose logging enabled
                    if (logOutput)
                    {
                        Log("Get count of created contacts...");
                        int destCount = o365_dest_contactfolder.GetContactsAsync().Result.Count;
                        LogLn(" Done. " + destCount.ToString() + " found.");
                    }
                }
            }
            catch (Exception e)
            {
                if (e.InnerException != null)
                    Console.WriteLine(" Error occured: " + e.InnerException.Message);
                else
                    Console.WriteLine(" Error occured: " + e.Message);
            }
            finally
            {
                Log("Close Outlook...");
                if (outlookApp != null)
                {
                    outlookApp.Quit();
                }
                LogLn(" Done.");

                Log("Disconnect Office365...");
                //outlookCloud.Disconnect(); //not yet implemented or neccessary (?!)
                LogLn(" Done.");
            }

            LogLn("End.");

#if DEBUG
            Console.ReadLine();
#endif
        }
    }
}
