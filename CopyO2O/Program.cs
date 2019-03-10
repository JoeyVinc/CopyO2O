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
        static void Log(string value, bool toConsole = false, bool suppressDateTime = false) { Output.Print(value: (!suppressDateTime ? DateTime.Now.ToString() + " - " : "") + value, logEnabled: logOutput); if (!logOutput && toConsole) Console.Write(value); }
        static void LogLn(string value, bool toConsole = false, bool suppressDateTime = false) { Log(value + "\n", toConsole, suppressDateTime); }

        static void Main(string[] args)
        {
            const string appId = "cb2c2f84-63f0-49d8-9335-79fcc6050654";
            List<string> AppPermissions = new List<string> { "User.Read", "Calendars.ReadWrite", "Contacts.ReadWrite" };

            DateTime from = DateTime.Now.AddMonths(-1);
            DateTime to = DateTime.Now.AddMonths(1);
            int clearpast = 0;
            string calendar_source_Name = "";
            string calendar_destination_Name = "";
            string contacts_source_Name = "";
            string contacts_destination_Name = "";
            string proxy = "";
            bool clrNotExisting = true; //clear every items of target which does not exist in source side

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
                LogLn("Error: " + e.Message + "\n", true);
                LogLn("Parameters:\n"
                    + "/CAL:\"<source>\";\"<destination>\" : Calendar source and destination\n"
                    + "/CON:\"<source>\";\"<destination>\" : Contacts source and destination\n"
                    + "[opt] /from:<date>              : for calendar: First date to sync (DD.MM.YYYY) or relative to today (in days; eg. -10)\n"
                    + "[opt] /to:<date>                : for calendar: Last date to sync (DD.MM.YYYY) or relative to today (in days; eg. 8)\n"
                    + "[opt] /clear:<days>             : for calendar: Clear <days> in the past (from 'from' back)\n"
                    + "[opt] /proxy:<address>          : set if an explicit proxy should be used for connection\n"
                    + "[opt] /log                      : Verbose logging\n\n"
                    + "Example: CopyO2O /CAL:\"Hans.Mustermann@company.com\\Calendar\";\"Business\" /from:-7 /to:30 /clear:14", true);
                System.Environment.Exit(-1);
            }

            LogLn("Start copy...", true);

            //load the sync cache
            Helpers.LoadSyncCache();
            List<Helpers.SyncInfo> syncWork = new List<Helpers.SyncInfo>();

            //set proxy if necessary
            if (proxy != "")
            { System.Net.WebRequest.DefaultWebProxy = new System.Net.WebProxy(proxy, true); }

            Outlook.Application outlookApp = null;
            Office365.MSGraph office365 = null;
            try
            {
                Log("Open Outlook...");
                outlookApp = new Outlook.Application();
                LogLn(" Done.", false, true);

                Log("Connect to Office365...");
                office365 = new Office365.MSGraph(appId, AppPermissions);
                LogLn(" Done.", false, true);

                //if calendar values should be synced
                if (SyncCAL())
                {
                    LogLn("Calendar: '" + calendar_source_Name + "' >> '" + (calendar_destination_Name ?? "DEFAULT") + "'" + " from " + from.ToShortDateString() + " to " + to.ToShortDateString(), true);

                    Log("Get all events of Outlook...");
                    Outlook.Calendar src_calendar = outlookApp.GetCalendar(calendar_source_Name);
                    Events srcEvents = src_calendar.GetItems(from, to);
                    LogLn(" Done. " + srcEvents.Count.ToString() + " found.", false, true);

                    /*srcEvents.ForEach((item) =>
                    {
                        syncWork.Add(new Helpers.SyncInfo { Type = Helpers.ID_type.CalItem, OutlookID = item.OriginId, SyncDirection = Helpers.SyncDirection.ToRight, SyncWorkMethod = Helpers.SyncMethod.CREATE });
                    });*/

                    Log("Get all events of O365...");
                    Office365.Calendar o365_dest_calendar = (new Office365.Calendars(office365)).GetCalendar(calendar_destination_Name ?? "Calendar"); //Calendar is the default calendar folder
                    List<Microsoft.Graph.Event> destEvents = o365_dest_calendar.GetItemsAsync(from, to).Result;
                    LogLn(" Done. " + destEvents.Count.ToString() + " found.", false, true);

                    /*destEvents.ForEach((item) =>
                    {
                        syncWork.Add(new Helpers.SyncInfo { Type = Helpers.ID_type.CalItem, O365ID = item.Id, SyncDirection = Helpers.SyncDirection.ToRight, SyncWorkMethod = Helpers.SyncMethod.DELETE });
                    });*/

                    /*
                    Log("Mark calendar items to delete...");
                    List<KeyValuePair<Helpers.ID_system, string>> calendarItemsToDelete = new List<KeyValuePair<Helpers.ID_system, string>>();
                    o365_dest_calendar.GetItemsAsync(from.AddDays(-clearpast), to).Result.ForEach((x) => { calendarItemsToDelete.Add(new KeyValuePair<Helpers.ID_system, string>(Helpers.ID_system.O365, x.Id)); });
                    Helpers.SetIDsToRemove(calendarItemsToDelete, Helpers.ID_type.CalItem);

                    
                    LogLn(" Done.", false, true);

                    /*
                    Log("Create events in O365...");
                    List<Office365.Replacement> replacements = new List<Office365.Replacement>(); //new List<Office365.Replacement> { new Office365.Replacement { key = "Location", regex = ".*(Vertex).*", newvalue = "Fasanenweg 9, 70771 Leinfelden-Echterdingen, Deutschland" } };
                    o365_dest_calendar.Add(srcEvents, replacements);
                    LogLn(" Done.", false, true);

                    Log("Delete doublettes in O365...");
                    List<string> calendarItemsFromFileToDelete = Helpers.GetIDsToRemove(Helpers.ID_type.CalItem).Where(x => (x.Key == Helpers.ID_system.O365)).Select(x => x.Value).ToList<string>();

                    //if doublettes successfully deleted remove cache file
                    if (o365_dest_calendar.Delete(calendarItemsFromFileToDelete) == true)
                        Helpers.DeleteRemovalLog();
                    LogLn(" Done.", false, true);
                    */

                    //get all ids of outlook which have to be synced
                    List<string> tmpCacheIds = Helpers.syncCache.Select((x) => x.OutlookID).ToList(); //all Ids in cache
                    List<string> tmpOutlookIds = srcEvents.Select((x) => x.OriginId).ToList(); //all Ids on Outlook side
                    List<string> tmpNewOutlookIds = tmpOutlookIds.Except(tmpCacheIds).ToList(); //all Outlook-ids which are NOT contained in sync cache => all NEW outlook ids
                    List<string> tmpModifiedOutlookIds = new List<string>(); //all Outlook-ids which were modified AFTER the last sync
                    {
                        Helpers.syncCache.ForEach(
                            (item) =>
                            {
                                if (srcEvents.Exists(x => (x.OriginId == item.OutlookID)))
                                {
                                    if (srcEvents.Find(x => (x.OriginId == item.OutlookID)).LastModTime > item.LastSyncTime.AddSeconds(10))
                                        tmpModifiedOutlookIds.Add(item.OutlookID);
                                }
                            });
                    }
                    List<string> tmpDeletedOutlookIds = tmpCacheIds.Except(tmpOutlookIds).ToList(); //all Outlook-ids which are synced but does not exist anymore => all DELETED outlook ids

                    tmpCacheIds = Helpers.syncCache.Select((x) => x.O365ID).ToList(); //all Ids in cache
                    List<string> tmpO365Ids = destEvents.Select((x) => x.Id).ToList(); //all Ids on O365 side
                    List<string> tmpNewO365Ids = tmpO365Ids.Except(tmpCacheIds).ToList(); //all O365-ids which are NOT contained in sync cache => all NEW O365 ids
                    List<string> tmpModifiedO365Ids = new List<string>(); //all O365-ids which were modified AFTER the last sync
                    {
                        Helpers.syncCache.ForEach(
                            (item) =>
                            {
                                if (destEvents.Exists(x => (x.Id == item.O365ID)))
                                {
                                    if (destEvents.Find(x => (x.Id == item.O365ID)).LastModifiedDateTime > item.LastSyncTime.AddSeconds(10))
                                        tmpModifiedO365Ids.Add(item.O365ID);
                                }
                            });
                    }
                    List<string> tmpDeletedO365Ids = tmpCacheIds.Except(tmpO365Ids).ToList(); //all ids of sync cache which does NOT exist in O365 anymore => all DELETED o365 ids


                    { //1st OUTLOOK is the leading system, 2nd only sync to O365
                        //add new outlook-elemente to sync list
                        tmpNewOutlookIds.ForEach((id) => syncWork.Add(new Helpers.SyncInfo { OutlookID = id, Type = Helpers.ID_type.CalItem, SyncDirection = Helpers.SyncDirection.ToRight, SyncWorkMethod = Helpers.SyncMethod.CREATE }));

                        //add modified outlook-elemente to sync list
                        tmpModifiedOutlookIds.ForEach((id) =>
                        {
                            syncWork.Add(new Helpers.SyncInfo { OutlookID = id, Type = Helpers.ID_type.CalItem, SyncDirection = Helpers.SyncDirection.ToRight, SyncWorkMethod = Helpers.SyncMethod.CREATE });
                            syncWork.Add(new Helpers.SyncInfo { O365ID = Helpers.syncCache.Find(x => (x.OutlookID == id)).O365ID, Type = Helpers.ID_type.CalItem, SyncDirection = Helpers.SyncDirection.ToRight, SyncWorkMethod = Helpers.SyncMethod.DELETE });
                        });

                        //remove deleted outlook elements from o365
                        tmpDeletedOutlookIds.ForEach((id) => syncWork.Add(new Helpers.SyncInfo { O365ID = Helpers.syncCache.Find(x => (x.OutlookID == id)).O365ID, Type = Helpers.ID_type.CalItem, SyncDirection = Helpers.SyncDirection.ToRight, SyncWorkMethod = Helpers.SyncMethod.DELETE }));

                        //restore from Office 365 deleted outlook-contacts again
                        tmpDeletedO365Ids.ForEach((id) =>
                        {
                            syncWork.Add(new Helpers.SyncInfo { OutlookID = Helpers.syncCache.Find(x => (x.O365ID == id)).OutlookID, Type = Helpers.ID_type.CalItem, SyncDirection = Helpers.SyncDirection.ToRight, SyncWorkMethod = Helpers.SyncMethod.CREATE });
                            Helpers.syncCache.RemoveAll(x => (x.O365ID == id));
                        });

                        //if all items which does not exist in outlook should be removed from O365
                        if (clrNotExisting)
                        { tmpNewO365Ids.ForEach((id) => syncWork.Add(new Helpers.SyncInfo { O365ID = id, Type = Helpers.ID_type.CalItem, SyncDirection = Helpers.SyncDirection.ToRight, SyncWorkMethod = Helpers.SyncMethod.DELETE })); }
                    }

//===========================================

                    Log("Create events in O365...");
                    List<Task> createTasks = new List<Task>();
                    syncWork.Where(x => ((x.SyncWorkMethod == Helpers.SyncMethod.CREATE) && (x.Type == Helpers.ID_type.CalItem))).ToList().ForEach(
                        (item) =>
                        {
                            createTasks.Add(o365_dest_calendar.AddAsync(srcEvents.Find(x => (x.OriginId == item.OutlookID)))
                                .ContinueWith((i) => Helpers.O365_Event_Added(i.Result.Id, item.OutlookID)));
                        });
                    Task.WaitAll(createTasks.ToArray());
                    LogLn(" Done.", false, true);

                    Log("Delete events in O365...");
                    List<Task> deleteTasks = new List<Task>();
                    syncWork.Where(x => ((x.SyncWorkMethod == Helpers.SyncMethod.DELETE) && (x.Type == Helpers.ID_type.CalItem))).ToList().ForEach(
                        (item) =>
                        {
                            deleteTasks.Add(o365_dest_calendar.DeleteAsync(item.O365ID)
                                .ContinueWith((i) => Helpers.O365_Item_Removed(item.O365ID)));
                        });
                    Task.WaitAll(deleteTasks.ToArray());
                    LogLn(" Done.", false, true);

                    //exec only if verbose logging enabled
                    if (logOutput)
                    {
                        Log("Get count of events in O365...");
                        int destCount = o365_dest_calendar.GetItemsAsync(from.AddDays(-clearpast), to).Result.Count;
                        LogLn(" Done. " + destCount.ToString() + " found.", false, true);
                    }

                    LogLn(srcEvents.Count.ToString() + " events copied.", true);
                }

                //if contacts should be synced
                if (SyncCON())
                {
                    LogLn("Contacts: '" + contacts_source_Name + "' >> '" + (contacts_destination_Name ?? "DEFAULT") + "'", true);

                    Log("Get all contacts of Outlook...");
                    Outlook.ContactFolder src_contactfolder = outlookApp.GetContactFolder(contacts_source_Name);
                    ContactCollectionType srcContacts = src_contactfolder.GetItems();
                    LogLn(" Done. " + srcContacts.Count.ToString() + " found.", false, true);

                    Log("Get all contacts of O365...");
                    Office365.ContactFolder o365_dest_contactfolder = (new Office365.ContactFolders(office365)).GetContactFolder(contacts_destination_Name);
                    List<Microsoft.Graph.Contact> destContacts = o365_dest_contactfolder.GetContactsAsync().Result;
                    LogLn(" Done. " + destContacts.Count.ToString() + " found.", false, true);

                    //get all ids of outlook which have to be synced
                    List<string> tmpCacheIds = Helpers.syncCache.Select((x) => x.OutlookID).ToList(); //all Ids in cache
                    List<string> tmpOutlookIds = srcContacts.Select((x) => x.OriginId).ToList(); //all Ids on Outlook side
                    List<string> tmpNewOutlookIds = tmpOutlookIds.Except(tmpCacheIds).ToList(); //all Outlook-ids which are NOT contained in sync cache => all NEW outlook ids
                    List<string> tmpModifiedOutlookIds = new List<string>(); //all Outlook-ids which were modified AFTER the last sync
                    {
                        Helpers.syncCache.ForEach(
                            (item) =>
                            {
                                if (srcContacts.Exists(x => (x.OriginId == item.OutlookID)))
                                {
                                    if (srcContacts.Find(x => (x.OriginId == item.OutlookID)).LastModTime > item.LastSyncTime.AddSeconds(10))
                                        tmpModifiedOutlookIds.Add(item.OutlookID);
                                }
                            });
                    }
                    List<string> tmpDeletedOutlookIds = tmpCacheIds.Except(tmpOutlookIds).ToList(); //all Outlook-ids which are synced but does not exist anymore => all DELETED outlook ids

                    tmpCacheIds = Helpers.syncCache.Select((x) => x.O365ID).ToList(); //all Ids in cache
                    List<string> tmpO365Ids = destContacts.Select((x) => x.Id).ToList(); //all Ids on O365 side
                    List<string> tmpNewO365Ids = tmpO365Ids.Except(tmpCacheIds).ToList(); //all O365-ids which are NOT contained in sync cache => all NEW O365 ids
                    List<string> tmpModifiedO365Ids = new List<string>(); //all O365-ids which were modified AFTER the last sync
                    {
                        Helpers.syncCache.ForEach(
                            (item) =>
                            {
                                if (destContacts.Exists(x => (x.Id == item.O365ID)))
                                {
                                    if (destContacts.Find(x => (x.Id == item.O365ID)).LastModifiedDateTime > item.LastSyncTime.AddSeconds(10))
                                        tmpModifiedO365Ids.Add(item.O365ID);
                                }
                            });
                    }
                    List<string> tmpDeletedO365Ids = tmpCacheIds.Except(tmpO365Ids).ToList(); //all ids of sync cache which does NOT exist in O365 anymore => all DELETED o365 ids


                    { //1st OUTLOOK is the leading system, 2nd only sync to O365
                        //add new outlook-elemente to sync list
                        tmpNewOutlookIds.ForEach((id) => syncWork.Add(new Helpers.SyncInfo { OutlookID = id, Type = Helpers.ID_type.Contact, SyncDirection = Helpers.SyncDirection.ToRight, SyncWorkMethod = Helpers.SyncMethod.CREATE }));

                        //add modified outlook-elemente to sync list
                        tmpModifiedOutlookIds.ForEach((id) =>
                        {
                            syncWork.Add(new Helpers.SyncInfo { OutlookID = id, Type = Helpers.ID_type.Contact, SyncDirection = Helpers.SyncDirection.ToRight, SyncWorkMethod = Helpers.SyncMethod.CREATE });
                            syncWork.Add(new Helpers.SyncInfo { O365ID = Helpers.syncCache.Find(x => (x.OutlookID == id)).O365ID, Type = Helpers.ID_type.Contact, SyncDirection = Helpers.SyncDirection.ToRight, SyncWorkMethod = Helpers.SyncMethod.DELETE });
                        });
                        
                        //remove deleted outlook elements from o365
                        tmpDeletedOutlookIds.ForEach((id) => syncWork.Add(new Helpers.SyncInfo { O365ID = Helpers.syncCache.Find(x => (x.OutlookID == id)).O365ID, Type = Helpers.ID_type.Contact, SyncDirection = Helpers.SyncDirection.ToRight, SyncWorkMethod = Helpers.SyncMethod.DELETE }));

                        //restore from Office 365 deleted outlook-contacts again
                        tmpDeletedO365Ids.ForEach((id) =>
                        {
                            syncWork.Add(new Helpers.SyncInfo { OutlookID = Helpers.syncCache.Find(x => (x.O365ID == id)).OutlookID, Type = Helpers.ID_type.Contact, SyncDirection = Helpers.SyncDirection.ToRight, SyncWorkMethod = Helpers.SyncMethod.CREATE });
                            Helpers.syncCache.RemoveAll(x => (x.O365ID == id));
                        });

                        //if all items which does not exist in outlook should be removed from O365
                        if (clrNotExisting)
                        { tmpNewO365Ids.ForEach((id) => syncWork.Add(new Helpers.SyncInfo { O365ID = id, Type = Helpers.ID_type.Contact, SyncDirection = Helpers.SyncDirection.ToRight, SyncWorkMethod = Helpers.SyncMethod.DELETE })); }
                    }

                    Log("Create contacts in O365...");
                    List<Task> createTasks = new List<Task>();
                    syncWork.Where(x => ((x.SyncWorkMethod == Helpers.SyncMethod.CREATE) && (x.Type == Helpers.ID_type.Contact))).ToList().ForEach(
                        (item) =>
                        {
                            createTasks.Add(o365_dest_contactfolder.AddAsync(srcContacts.Find(x => (x.OriginId == item.OutlookID)))
                                .ContinueWith((i) => Helpers.O365_Contact_Added(i.Result.Id, item.OutlookID)));
                        });
                    Task.WaitAll(createTasks.ToArray());
                    LogLn(" Done.", false, true);

                    Log("Delete contacts in O365...");
                    List<Task> deleteTasks = new List<Task>();
                    syncWork.Where(x => ((x.SyncWorkMethod == Helpers.SyncMethod.DELETE) && (x.Type == Helpers.ID_type.Contact))).ToList().ForEach(
                        (item) =>
                        {
                            deleteTasks.Add(o365_dest_contactfolder.DeleteAsync(item.O365ID)
                                .ContinueWith((i) => Helpers.O365_Item_Removed(item.O365ID)));
                        });
                    Task.WaitAll(deleteTasks.ToArray());
                    LogLn(" Done.", false, true);

                    //exec only if verbose logging enabled
                    if (logOutput)
                    {
                        Log("Get count of contacts in O365...");
                        int destCount = o365_dest_contactfolder.GetContactsAsync().Result.Count;
                        LogLn(" Done. " + destCount.ToString() + " found.", false, true);
                    }

                    LogLn(srcContacts.Count.ToString() + " contacts copied.", true);
                }
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
                Helpers.StoreSyncCache();
            }

            LogLn(syncWork.Count.ToString() + " operations processed.", true);

#if DEBUG
            Console.ReadLine();
#endif
        }
    }
}
