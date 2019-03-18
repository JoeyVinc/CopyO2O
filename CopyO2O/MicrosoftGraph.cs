using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using System.Net.Http.Headers;
using System.Diagnostics;
using Commonfunctions.Convert;
using System.Text.RegularExpressions;


namespace CopyO2O.Office365
{
    public struct Replacement
    {
        public string key;
        public string regex;
        public string newvalue;
    }

    public class Graph
    {
        private PublicClientApplication appClient;
        private GraphServiceClient _graphClient = null;
        private const string Authority = "https://login.microsoftonline.com/common/";
        private List<string> Scope = new List<string> { };

        public bool IsValid { get; } = false; //TRUE if the graph could be successfully connected
        public int MaxItemsPerRequest = 50;
        public int DefaultRetryDelay = 500;

        public Graph(string appId, List<string> scope)
        {
            if (scope.Count == 0) throw new Exception("Scope not defined!");
            Scope = scope;

            try
            {
                //init connection to Graph
                appClient = new PublicClientApplication(appId, Authority, TokenCacheHelper.GetUserCache());
                appClient.RedirectUri = "urn:ietf:wg:oauth:2.0:oob";

                //open connection, request access token, test connection
                User testResult = GetGraphClient().Me.Request().GetAsync().Result;
                IsValid = true;
            }
            catch (Exception e)
            {
                Debug.WriteLine("ERROR MSGraph " + e.Message);
                IsValid = false;
                throw e;
            }

            Calendars = new Calendars(this);
            ContactFolders = new ContactFolders(this);

        }

        private async Task<AuthenticationResult> AuthResultAsync()
        {
            try { return await appClient.AcquireTokenSilentAsync(Scope, appClient.GetAccountsAsync().Result.First()); }
            catch (Exception e)
            {
                Debug.WriteLine("ERROR AuthResultAsync " + e.Message);
                return await appClient.AcquireTokenAsync(Scope);
            }
        }

        public GraphServiceClient GetGraphClient()
        {
            if (_graphClient == null)
            {
                _graphClient = new GraphServiceClient(
                         new DelegateAuthenticationProvider(
                         async (requestMessage) =>
                         {
                             string token = (await AuthResultAsync()).AccessToken;
                             requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);
                         }));
            }

            return _graphClient;
        }

        public Calendars Calendars;
        public ContactFolders ContactFolders;
    }

    public abstract class GraphContainerCollection<ContainerType>
    {
        public ContainerType this[string Name] { get => this.Get(Name); }
        public abstract ContainerType Get(string Name);

        public abstract Task<ContainerType> CreateAsync(string Name);
        public abstract Task DeleteAsync(string Name);
        public abstract Task RenameAsync(string OldName, string NewName);

        public Graph connection = null;

        //constructor
        public GraphContainerCollection(Graph Graph)
        {
            connection = Graph;
        }
    }

    public abstract class GraphContainer<ItemType>
    {
        protected object parent { get; }
        protected string name;
        protected string id;

        public string Name { get => name; }
        public string Id { get => id; }

        public ItemType this[string Id] { get => this.Get(Id); }
        public abstract ItemType Get(string Id);

        //public abstract Task<string> AddAsync(object Item);
        //public abstract Task<bool> RemoveAsync(string Id);

        //constructor
        public GraphContainer(object Parent, string Name, string Id)
        {
            parent = Parent;
            name = Name;
            id = Id;
        }
    }

    public interface IGraphContainer
    {
         string Name { get; }
         string Id { get; }

        Task<string> AddAsync(object Item);
        Task<bool> RemoveAsync(string Id);
    }

    public class Calendars : GraphContainerCollection<Calendar>
    {
        private Dictionary<string, string> _calendarIds = new Dictionary<string, string>();
        public override Calendar Get(string CalendarName)
        {
            return new Calendar(this, CalendarName, this.GetCalendarId(CalendarName));
        }

        public Calendars(Graph Graph) : base(Graph) { }

        private string GetCalendarId(string CalendarName, bool RenewRequest = false)
        {
            if (!RenewRequest)
            {
                //if the calendarname is already known return the corresponding ID
                if (_calendarIds.ContainsKey(CalendarName.ToUpper()))
                    return _calendarIds[CalendarName.ToUpper()];
            }

            List<QueryOption> options = new List<QueryOption>();
            options.Add(new QueryOption("$top", connection.MaxItemsPerRequest.ToString()));

            //request all calendars
            IUserCalendarsCollectionPage calendars = connection.GetGraphClient().Me.Calendars.Request(options).GetAsync().Result;

            bool requestNextPage;
            do
            {
                requestNextPage = (calendars.NextPageRequest != null);

                //search for id of specified calender
                for (int i = 0; i < calendars.Count; i++)
                {
                    //if the calendar was found return the ID
                    if (calendars[i].Name.ToUpper() == CalendarName.ToUpper())
                    {
                        _calendarIds.Add(CalendarName.ToUpper(), calendars[i].Id);
                        return calendars[i].Id;
                    }
                }

                if (requestNextPage)
                    calendars = calendars.NextPageRequest.GetAsync().Result;

            } while (requestNextPage);

            throw new Exception("Calendar named '" + CalendarName + "' could not be found!");
        }

        /// <summary>
        /// create calendar
        /// </summary>
        /// <param name="CalendarName">Name of calendar to delete</param>
        /// <returns>the ID of the created calendar</returns>
        /// 
        public override async Task<Calendar> CreateAsync(string CalendarName)
        {
            return new Calendar(this, null, await this.CreateAsync(CalendarName, CalendarColor.Auto));
        }

        /// <summary>
        /// create calendar
        /// </summary>
        /// <param name="CalendarName">Name of calendar to delete</param>
        /// <returns>the ID of the created calendar</returns>
        /// 
        public async Task<string> CreateAsync(string CalendarName, CalendarColor color = CalendarColor.Auto)
        {
            //if calendar does not exist
            if (GetCalendarId(CalendarName, true) == "")
            {
                try
                {
                    Microsoft.Graph.Calendar newCal = new Microsoft.Graph.Calendar
                    {
                        Name = CalendarName,
                        Color = color
                    };

                    return (await connection.GetGraphClient().Me.Calendars.Request().AddAsync(newCal)).Id;
                }
                catch (Exception e)
                {
                    Debug.WriteLine("ERROR CreateAsync " + e.Message);
                    return "";
                }
            }
            else return "";
        }

        /// <summary>
        /// delete calendar
        /// </summary>
        /// <param name="CalendarName">Name of calendar to delete</param>
        /// <returns>TRUE if successful otherwise FALSE</returns>
        /// 
        public override async Task DeleteAsync(string CalendarName)
        {
            string calendarID = GetCalendarId(CalendarName);

            try
            {
                await RenameAsync(CalendarName, "tmp_" + DateTime.Now.ToString("g") + ".tobedeleted");
                await connection.GetGraphClient().Me.Calendars[calendarID].Request().DeleteAsync();

                _calendarIds.Remove(CalendarName.ToUpper());
            }
            catch (Microsoft.Graph.ServiceException e)
            {
                Debug.WriteLine("ERROR DeleteAsync " + e.Message);

                if (e.StatusCode == System.Net.HttpStatusCode.Conflict)
                {
                    await connection.GetGraphClient().Me.Calendars[calendarID].Request().DeleteAsync();

                    _calendarIds.Remove(CalendarName.ToUpper());
                }
                else
                    throw e;
            }
        }

        /// <summary>
        /// renames a calendar
        /// </summary>
        /// <param name="OldCalendarName">Original/old name of calendar to rename</param>
        /// <param name="NewCalendarName">new name</param>
        /// <param name="NameIsId">TRUE if OldCalendarName is the ID (not the name)</param>
        /// <returns></returns>
        /// 
        public override async Task RenameAsync(string OldCalendarName, string NewCalendarName)
        {
            string calendarID = GetCalendarId(OldCalendarName);

            try
            {
                Microsoft.Graph.Calendar tempCalendar = connection.GetGraphClient().Me.Calendars[calendarID].Request().GetAsync().Result;
                tempCalendar.Name = NewCalendarName;
                await connection.GetGraphClient().Me.Calendars[calendarID].Request().UpdateAsync(tempCalendar);

                _calendarIds.Remove(OldCalendarName.ToUpper());
            }
            catch (Exception e)
            {
                Debug.WriteLine("ERROR RenameAsync " + e.Message);
                throw e;
            }
        }
    }

    public class Calendar : GraphContainer<Microsoft.Graph.Event>, IGraphContainer
    {
        private Dictionary<string, string> deltaSets = new Dictionary<string, string>();

        public Calendar(Calendars Parent, string Name, string Id) : base(Parent, Name, Id) { }

        public override Microsoft.Graph.Event Get(string EventId)
        {
            return (parent as Calendars).connection.GetGraphClient().Me.Calendars[this.Id].Events[EventId].Request().GetAsync().Result;
        }

        public async Task<List<Microsoft.Graph.Event>> getItemsAsync(DateTime From, DateTime To)
        {
            List<Microsoft.Graph.Event> result = new List<Microsoft.Graph.Event>();

            //request all events expanded (calendarview -> resolved recurring events)
            List<QueryOption> options = new List<QueryOption>();
            options.Add(new QueryOption("$top", (parent as Calendars).connection.MaxItemsPerRequest.ToString()));
            options.Add(new QueryOption("startDateTime", DateNTime.ConvertDateTimeToISO8016(From)));
            options.Add(new QueryOption("endDateTime", DateNTime.ConvertDateTimeToISO8016(To)));

            bool retry = false;
            int retryCounter = 0;
            do
            {
                try
                {
                    ICalendarCalendarViewCollectionPage items = await (parent as Calendars).connection.GetGraphClient().Me.Calendars[this.Id].CalendarView.Request(options).GetAsync();
                    bool requestNextPage = false;
                    do
                    {
                        requestNextPage = (items.NextPageRequest != null);
                        foreach (Microsoft.Graph.Event aptmItem in items)
                        {
                            result.Add(aptmItem);
                        }

                        if (requestNextPage)
                            items = await items.NextPageRequest.GetAsync();
                    }
                    while (requestNextPage);

                    return result; //quit and return found objects
                }
                catch (Microsoft.Graph.ServiceException e)
                {
                    Debug.WriteLine(e);
                    retry = false;

                    //if service is busy, too many connection, aso. retry request
                    if (e.StatusCode.ToString() == "429")
                    {
                        int sleeptime = (parent as Calendars).connection.DefaultRetryDelay;

                        if (e.ResponseHeaders.RetryAfter != null)
                            sleeptime = e.ResponseHeaders.RetryAfter.Delta.GetValueOrDefault(new TimeSpan(0, 0, 1)).Milliseconds;

                        System.Threading.Thread.Sleep(sleeptime);
                        retry = true;
                    }
                    //if the specified object was not found in the store => know BUG in O365; could be a false negative!
                    //therefore: Retry one more time after one second.
                    else
                    {
                        System.Threading.Thread.Sleep(1000);
                        retry = true;
                        retryCounter = retryCounter + 98;
                    }

                    retryCounter++;
                }
            } while (retry && (retryCounter <= 100)); //max 100 loops

            return result;
        }

        public async Task<Events> GetItemsAsync(DateTime from, DateTime to)
        {
            Events result = new Events();

            foreach (Microsoft.Graph.Event item in await this.getItemsAsync(from, to))
            {
                Event tmpEvent = new Event();
                tmpEvent.OriginId = item.Id;
                tmpEvent.LastModTime = item.LastModifiedDateTime.GetValueOrDefault().DateTime;

                tmpEvent.Subject = item.Subject;
                tmpEvent.StartDateTime = item.OriginalStart.GetValueOrDefault().DateTime;
                tmpEvent.StartTimeZone = item.OriginalStartTimeZone;
                tmpEvent.StartUTC = item.OriginalStart.GetValueOrDefault().UtcDateTime;
                tmpEvent.EndDateTime = DateTime.Parse(item.End.DateTime);
                tmpEvent.EndTimeZone = item.OriginalEndTimeZone;
                tmpEvent.EndUTC = tmpEvent.EndDateTime;
                tmpEvent.AllDayEvent = item.IsAllDay.GetValueOrDefault(false);

                tmpEvent.Location = item.Location.DisplayName;
                tmpEvent.Subject = item.Subject;
                tmpEvent.Body = item.Body.Content;
                tmpEvent.ReminderMinutesBefore = item.ReminderMinutesBeforeStart.GetValueOrDefault(0);
                tmpEvent.ReminderOn = (item.ReminderMinutesBeforeStart != null);
                tmpEvent.IsPrivate = (item.Sensitivity == Sensitivity.Private);

                switch (item.Importance)
                {
                    case Importance.High: tmpEvent.Importance = Event.ImportanceEnum.High; break;
                    case Importance.Low: tmpEvent.Importance = Event.ImportanceEnum.Low; break;
                    case Importance.Normal: tmpEvent.Importance = Event.ImportanceEnum.Normal; break;
                    default: tmpEvent.Importance = Event.ImportanceEnum.Normal; break;
                }

                switch (item.ShowAs)
                {
                    case FreeBusyStatus.Free : tmpEvent.Status = Event.StatusEnum.Free; break;
                    case FreeBusyStatus.Tentative: tmpEvent.Status = Event.StatusEnum.Tentative; break;
                    case FreeBusyStatus.Busy: tmpEvent.Status = Event.StatusEnum.Busy; break;
                    case FreeBusyStatus.Oof: tmpEvent.Status = Event.StatusEnum.OutOfOffice; break;
                    case FreeBusyStatus.WorkingElsewhere: tmpEvent.Status = Event.StatusEnum.ElseWhere; break;
                    default: tmpEvent.Status = Event.StatusEnum.Tentative; break;
                }

                result.Add(tmpEvent);
            }
            return result;
        }

        public async Task<List<Microsoft.Graph.Event>> GetItemsDeltaAsync(DateTime From, DateTime To, bool InitDelta = true)
        {
            List<Microsoft.Graph.Event> result = new List<Microsoft.Graph.Event>();

            //collect parameters for request
            List<QueryOption> options = new List<QueryOption>();
            options.Add(new QueryOption("odata.maxpagesize", (parent as Calendars).connection.MaxItemsPerRequest.ToString()));

            //if a new delta request should be sent
            if (InitDelta)
            {
                options.Add(new QueryOption("startDateTime", DateNTime.ConvertDateTimeToISO8016(From)));
                options.Add(new QueryOption("endDateTime", DateNTime.ConvertDateTimeToISO8016(To)));
            }
            //if an existing delta request should be resumed
            else
            {
                if (deltaSets.ContainsKey(this.Id) == false)
                    throw new ApplicationException("Error: Init delta run first!");

                options.Add(new QueryOption("$deltatoken", deltaSets[this.Id]));
            }

            //request all events expanded (calendarview -> resolved recurring events)
            IEventDeltaCollectionPage events = await (parent as Calendars).connection.GetGraphClient().Me.Calendars[this.Id].CalendarView.Delta().Request(options).GetAsync();

            //loop to request all pages/data
            bool requestNextPage = false;
            do
            {
                requestNextPage = (events.NextPageRequest != null);
                foreach (Microsoft.Graph.Event aptmItem in events)
                {
                    result.Add(aptmItem);
                }

                if (requestNextPage)
                    events = await events.NextPageRequest.GetAsync();
            }
            while (requestNextPage);

            //save the deltaToken for the next (delta-) request
            if (!deltaSets.ContainsKey(this.Id))
                deltaSets.Add(this.Id, "");

            string deltaToken = System.Net.WebUtility.UrlDecode(events.AdditionalData["@odata.deltaLink"].ToString()).Split(new string[] { "$deltatoken=" }, StringSplitOptions.RemoveEmptyEntries)[1];
            deltaSets[this.Id] = deltaToken;

            return result;
        }

        public async Task<List<Microsoft.Graph.Event>> GetItemsDeltaAsync(string DeltaToken)
        {
            //save the deltaToken for the next (delta-) request
            if (!deltaSets.ContainsKey(this.Id))
                deltaSets.Add(this.Id, "");

            deltaSets[this.Id] = DeltaToken;

            return await this.GetItemsDeltaAsync(DateTime.Now, DateTime.Now, false);
        }

        public async Task<string> AddAsync(object item)
        {
            bool retry = false;

            if (item is Microsoft.Graph.Event)
            {
                do
                {
                    try
                    { return (await (parent as Calendars).connection.GetGraphClient().Me.Calendars[this.Id].Events.Request().AddAsync(item as Microsoft.Graph.Event)).Id; }
                    catch (Microsoft.Graph.ServiceException e)
                    {
                        Debug.WriteLine("ERROR CreateItemAsync " + e.Message);
                        retry = false;

                        //if service is busy, too many connection, aso. retry request
                        if (e.StatusCode.ToString() == "429")
                        {
                            int sleeptime = (parent as Calendars).connection.DefaultRetryDelay;

                            if (e.ResponseHeaders.RetryAfter != null)
                                sleeptime = e.ResponseHeaders.RetryAfter.Delta.GetValueOrDefault(new TimeSpan(0, 0, 1)).Milliseconds;

                            System.Threading.Thread.Sleep(sleeptime);
                            retry = true;
                        }
                    }
                } while (retry == true);
            }
            else if (item is Event)
            {
                return await this.addAsync(item as Event);
            }

            return null;
        }

        private async Task<string> addAsync(Event item)
        {
            Microsoft.Graph.Event newEvent = new Microsoft.Graph.Event();

            newEvent.Start = new Microsoft.Graph.DateTimeTimeZone()
            {
                DateTime = ((DateTime)item.StartDateTime).ToString("s"),
                TimeZone = item.StartTimeZone
            };
            newEvent.End = new Microsoft.Graph.DateTimeTimeZone()
            {
                DateTime = ((DateTime)item.EndDateTime).ToString("s"),
                TimeZone = item.EndTimeZone
            };

            newEvent.Subject = item.Subject;
            newEvent.Location = new Microsoft.Graph.Location() { DisplayName = (item.Location ?? "") };
            newEvent.IsAllDay = item.AllDayEvent;

            newEvent.ReminderMinutesBeforeStart = item.ReminderMinutesBefore;
            newEvent.IsReminderOn = item.ReminderOn;

            switch (item.Status)
            {
                case Event.StatusEnum.Free: newEvent.ShowAs = Microsoft.Graph.FreeBusyStatus.Free; break;
                case Event.StatusEnum.Tentative: newEvent.ShowAs = Microsoft.Graph.FreeBusyStatus.Tentative; break;
                case Event.StatusEnum.Busy: newEvent.ShowAs = Microsoft.Graph.FreeBusyStatus.Busy; break;
                case Event.StatusEnum.OutOfOffice: newEvent.ShowAs = Microsoft.Graph.FreeBusyStatus.Oof; break;
                case Event.StatusEnum.ElseWhere: newEvent.ShowAs = Microsoft.Graph.FreeBusyStatus.WorkingElsewhere; break;
                default: newEvent.ShowAs = Microsoft.Graph.FreeBusyStatus.Free; break;
            }
            
            newEvent.ICalUId = item.OriginId; //set but ignored by MS Graph *hmpf*

            return await this.AddAsync(newEvent);
        }

        /// <summary>
        /// Creates new items - transformed by Replacement settings (RegEx)
        /// </summary>
        /// <param name="items">New Items to create</param>
        /// <param name="ReplacementSettings">Replacement settings which should be used to every item</param>
        /// 
        public void Add(Events items, List<Replacement> ReplacementSettings)
        {
            //if no items should be created
            if (items.Count == 0) return;

            Events newItems = items;

            //for each replacement setting
            foreach (Replacement replItem in ReplacementSettings)
            {
                System.Reflection.FieldInfo keyMember = newItems[0].GetType().GetField(replItem.key);

                //loop through all events to create
                foreach (Event item in newItems)
                {
                    Debug.WriteLine(keyMember.GetValue(item));

                    string oldValue = (string)keyMember.GetValue(item) ?? "";
                    string newValue = Regex.Replace(oldValue, replItem.regex, replItem.newvalue);
                    keyMember.SetValue(item, newValue);

                    Debug.WriteLine(keyMember.GetValue(item));

                }
            }

            List<Task> addTasks = new List<Task>();

            //create transformed items
            foreach (Event item in newItems)
            {
                addTasks.Add(this.AddAsync(item));
            }

            Task.WaitAll(addTasks.ToArray());
        }

        public async Task UpdateAsync(string eventId, Microsoft.Graph.Event updatedItem)
        {
            bool retry = false;

            do
            {
                try
                { await (parent as Calendars).connection.GetGraphClient().Me.Events[eventId].Request().UpdateAsync(updatedItem); }
                catch (Microsoft.Graph.ServiceException e)
                {
                    Debug.WriteLine("ERROR UpdateItemAsync " + e.Message);
                    retry = false;

                    //if service is busy, too many connection, aso. retry request
                    if (e.StatusCode.ToString() == "429")
                    {
                        int sleeptime = (parent as Calendars).connection.DefaultRetryDelay;

                        if (e.ResponseHeaders.RetryAfter != null)
                            sleeptime = e.ResponseHeaders.RetryAfter.Delta.GetValueOrDefault(new TimeSpan(0, 0, 1)).Milliseconds;

                        System.Threading.Thread.Sleep(sleeptime);
                        retry = true;
                    }
                }
            } while (retry == true);
        }

        public async Task<bool> RemoveAsync(string eventId)
        {
            bool retry = false;

            do
            {
                try
                {
                    await (parent as Calendars).connection.GetGraphClient().Me.Events[eventId].Request().DeleteAsync();
                    return true;
                }
                catch (Microsoft.Graph.ServiceException e)
                {
                    Debug.WriteLine(e);
                    retry = false;

                    //if service is busy, too many connection, aso. retry request
                    if (e.StatusCode.ToString() == "429")
                    {
                        int sleeptime = (parent as Calendars).connection.DefaultRetryDelay;

                        if (e.ResponseHeaders.RetryAfter != null)
                            sleeptime = e.ResponseHeaders.RetryAfter.Delta.GetValueOrDefault(new TimeSpan(0, 0, 1)).Milliseconds;

                        System.Threading.Thread.Sleep(sleeptime);
                        retry = true;
                    }
                }
            } while (retry == true);

            //not successfull
            return false;
        }

        public bool Delete(DateTime from, DateTime to)
        {
            List<Task> deleteThreads = new List<Task>();

            try
            {
                List<Microsoft.Graph.Event> items = this.getItemsAsync(from, to).Result;
                foreach (Microsoft.Graph.Event item in items)
                { deleteThreads.Add(this.RemoveAsync(item.Id)); }

                Task.WaitAll(deleteThreads.ToArray());
                return true;
            }
            catch (Microsoft.Graph.ServiceException e)
            {
                Debug.WriteLine("ERROR DeleteItems " + e.Message);
                return false;
            }
        }

        /// <summary>
        /// Delete calendar events which are contained in the IDs list
        /// </summary>
        /// <param name="IDs">List of IDs to delete</param>
        /// <returns>TRUE if successfull otherwise false</returns>
        /// 
        public bool Delete(List<string> IDs)
        {
            List<Task> deleteThreads = new List<Task>();

            try
            {
                foreach (string id in IDs)
                { deleteThreads.Add(this.RemoveAsync(id)); }

                Task.WaitAll(deleteThreads.ToArray());
                return true;
            }
            catch (Microsoft.Graph.ServiceException e)
            {
                Debug.WriteLine("ERROR DeleteItems(List) " + e.Message);
                return false;
            }
        }
    }

    /*public class CalendarItem : GraphItem<Microsoft.Graph.Event>
    {
        public override string Id { get => this.OriginItem.Id; }
        public override string Name { get => this.OriginItem.Subject; }

        public CalendarItem(Microsoft.Graph.Event item) : base(item) { }
    }*/

    public class ContactFolders : GraphContainerCollection<ContactFolder>
    {
        private Dictionary<string, string> _contactFolderIds = new Dictionary<string, string>();

        //constructor
        public ContactFolders(Graph Connection) : base(Connection)
        {
            //read all sub-folders initially
            GetFolders();
        }

        /// <summary>
        /// Get all sub folders which are known - Remark: Not the default folder! *hrgh*
        /// </summary>
        /// <param name="RenewRequest"></param>
        /// 
        private void GetFolders()
        {
            List<QueryOption> options = new List<QueryOption>();
            options.Add(new QueryOption("$top", connection.MaxItemsPerRequest.ToString()));

            //request all contactFolders
            IUserContactFoldersCollectionPage contactFolders = connection.GetGraphClient().Me.ContactFolders.Request(options).GetAsync().Result;

            //if at least one folder found
            if (contactFolders != null)
            {
                this._contactFolderIds.Clear();

                bool requestNextPage;
                do
                {
                    requestNextPage = (contactFolders.NextPageRequest != null);

                    foreach (Microsoft.Graph.ContactFolder folder in contactFolders)
                    {
                        this._contactFolderIds.Add(folder.DisplayName.ToUpper(), folder.Id);
                    }

                    if (requestNextPage)
                        contactFolders = contactFolders.NextPageRequest.GetAsync().Result;

                } while (requestNextPage);
            }
        }

        /// <summary>
        /// Get the corresponding ID of a contact/subfolder (not the default one!)
        /// </summary>
        /// <param name="ContactFolderName">Display name of a contact</param>
        /// <param name="RenewRequest">If true the folders will be new readen from MS Graph; else it will be from cache</param>
        /// <returns>The unique ID of the contact</returns>
        /// 
        private string GetId(string ContactFolderName, bool RenewRequest = false)
        {
            //if nothing set use default folder
            if ((ContactFolderName ?? "") == "")
                return null;

            if (!RenewRequest)
            {
                //if the contactFoldername is already known return the corresponding ID
                if (_contactFolderIds.ContainsKey(ContactFolderName.ToUpper()))
                    return _contactFolderIds[ContactFolderName.ToUpper()];
            }

            GetFolders();
            //if the contactFoldername is already known return the corresponding ID
            if (_contactFolderIds.ContainsKey(ContactFolderName.ToUpper()))
                return _contactFolderIds[ContactFolderName.ToUpper()];

            throw new Exception("Office365: Contact-folder named '" + ContactFolderName + "' could not be found!");
        }

        /// <summary>
        /// Return a ContactFolder object
        /// </summary>
        /// <param name="ContactFolderName">Display name of a contact folder</param>
        /// <returns></returns>
        public override ContactFolder Get(string ContactFolderName)
        {
            return new ContactFolder(this, ContactFolderName, this.GetId(ContactFolderName));
        }

        /// <summary>
        /// Create a Contact Folder (subfolder of the default root)
        /// </summary>
        /// <param name="ContactFolderName">Name of contactFolder to delete</param>
        /// <returns>the ID of the created contactFolder</returns>
        /// 
        public override async Task<ContactFolder> CreateAsync(string ContactFolderName)
        {
            //if contactFolder does not exist
            if (GetId(ContactFolderName, true) == "")
            {
                try
                {
                    Microsoft.Graph.ContactFolder newCal = new Microsoft.Graph.ContactFolder
                    {
                        DisplayName = ContactFolderName,
                    };

                    return this.Get((await connection.GetGraphClient().Me.ContactFolders.Request().AddAsync(newCal)).DisplayName);
                }
                catch (Exception e)
                {
                    Debug.WriteLine("ERROR AddAsync " + e.Message);
                    return null;
                }
            }
            else return null;
        }

        /// <summary>
        /// delete contactFolder
        /// </summary>
        /// <param name="ContactFolderName">Name of contactFolder to delete</param>
        /// <returns>TRUE if successful otherwise FALSE</returns>
        public override async Task DeleteAsync(string ContactFolderName)
        {
            string contactFolderID = GetId(ContactFolderName);

            //if contactFolder exists
            if (contactFolderID != "")
            {
                try
                {
                    await RenameAsync(ContactFolderName, "tmp_" + DateTime.Now.ToString("g") + ".tobedeleted");
                    await connection.GetGraphClient().Me.ContactFolders[contactFolderID].Request().DeleteAsync();

                    _contactFolderIds.Remove(ContactFolderName.ToUpper());
                }
                catch (Microsoft.Graph.ServiceException e)
                {
                    Debug.WriteLine("ERROR DeleteAsync " + e.Message);

                    if (e.StatusCode == System.Net.HttpStatusCode.Conflict)
                    {
                        await connection.GetGraphClient().Me.ContactFolders[contactFolderID].Request().DeleteAsync();

                        _contactFolderIds.Remove(ContactFolderName.ToUpper());
                    }
                    else throw e;
                }
            }
        }

        /// <summary>
        /// renames a contactFolder
        /// </summary>
        /// <param name="OldContactFolderName">Original/old name of contactFolder to rename</param>
        /// <param name="NewContactFolderName">new name</param>
        /// <param name="NameIsId">TRUE if OldContactFolderName is the ID (not the name)</param>
        /// <returns></returns>
        public override async Task RenameAsync(string OldContactFolderName, string NewContactFolderName)
        {
            string contactFolderID = GetId(OldContactFolderName);

            try
            {
                Microsoft.Graph.ContactFolder tempContactFolder = connection.GetGraphClient().Me.ContactFolders[contactFolderID].Request().GetAsync().Result;
                tempContactFolder.DisplayName = NewContactFolderName;
                await connection.GetGraphClient().Me.ContactFolders[contactFolderID].Request().UpdateAsync(tempContactFolder);

                _contactFolderIds.Remove(OldContactFolderName.ToUpper());
            }
            catch (Exception e)
            {
                Debug.WriteLine("ERROR RenameAsync " + e.Message);
                throw e;
            }
        }
    }

    public class ContactFolder : GraphContainer<Microsoft.Graph.Contact>, IGraphContainer
    {
        private bool useDefault = false;

        public string ContactFolderName { get => UseDefault ? null : name; }
        public string ContactFolderId { get => UseDefault ? null : id; }
        public bool UseDefault { get { return useDefault; } set { useDefault = value; name = null; id = null; } }

        public delegate void AddEventHandler(string ID);
        public AddEventHandler Event_Add = null;


        /// <summary>
        /// constructor of ContactFolder. ContactFolder is a SUBfolder for contacts. There is also a default folder (*hrgh*) which does not have a real name.
        /// </summary>
        /// <param name="GraphClient"></param>
        /// <param name="Name">Display name of a folder. If empty or null the default folder will be used.</param>
        /// <param name="Id">ID of the folder</param>
        /// 
        public ContactFolder(ContactFolders Parent, string Name, string Id) : base (Parent, Name, Id)
        {
            //if not default
            if ((Name ?? "") != "")
            {
                name = Name;
                id = Id;
            }
            else //use default folder - which does not have a name *hrgh*
            {
                UseDefault = true;
            }
        }

        /// <summary>
        /// return a specific contact
        /// </summary>
        /// <param name="ItemId"></param>
        /// <returns>Return the contact object</returns>
        /// 
        public override Microsoft.Graph.Contact Get(string ItemId)
        {
            bool retry = false;

            do
            {
                try
                {
                    //if default folder
                    if (this.useDefault)
                        return (parent as ContactFolders).connection.GetGraphClient().Me.Contacts[ItemId].Request().GetAsync().Result;
                    else
                        return (parent as ContactFolders).connection.GetGraphClient().Me.ContactFolders[Id].Contacts[ItemId].Request().GetAsync().Result;
                }
                catch (Microsoft.Graph.ServiceException e)
                {
                    Debug.WriteLine("ERROR ContactFolder THIS " + e.Message);
                    retry = false;

                    //if service is busy, too many connection, aso. retry request
                    if (e.StatusCode.ToString() == "429")
                    {
                        int sleeptime = (parent as ContactFolders).connection.DefaultRetryDelay;

                        if (e.ResponseHeaders.RetryAfter != null)
                            sleeptime = e.ResponseHeaders.RetryAfter.Delta.GetValueOrDefault(new TimeSpan(0, 0, 1)).Milliseconds;

                        System.Threading.Thread.Sleep(sleeptime);
                        retry = true;
                    }
                }
            } while (retry == true);

            return null;
        }

        /// <summary>
        /// Get all contacts in the current folder (or default root).
        /// </summary>
        /// <returns>Returns a list of contacts info</returns>
        /// 
        private async Task<List<Microsoft.Graph.Contact>> getContactsAsync()
        {
            List<Microsoft.Graph.Contact> result = new List<Microsoft.Graph.Contact>();

            List<QueryOption> options = new List<QueryOption>();
            options.Add(new QueryOption("$top", (parent as ContactFolders).connection.MaxItemsPerRequest.ToString()));

            //if default folder
            if (useDefault)
            {
                IUserContactsCollectionPage items = await (parent as ContactFolders).connection.GetGraphClient().Me.Contacts.Request(options).GetAsync();

                //if at least one item found
                if (items != null)
                {
                    bool requestNextPage = false;
                    do
                    {
                        requestNextPage = (items.NextPageRequest != null);
                        foreach (Microsoft.Graph.Contact item in items)
                        {
                            result.Add(item);
                        }

                        if (requestNextPage)
                            items = await items.NextPageRequest.GetAsync();
                    }
                    while (requestNextPage);
                }
            }
            else
            {
                IContactFolderContactsCollectionPage items = await (parent as ContactFolders).connection.GetGraphClient().Me.ContactFolders[Id].Contacts.Request(options).GetAsync();

                //if at least one item found
                if (items != null)
                {
                    bool requestNextPage = false;
                    do
                    {
                        requestNextPage = (items.NextPageRequest != null);
                        foreach (Microsoft.Graph.Contact item in items)
                        {
                            result.Add(item);
                        }

                        if (requestNextPage)
                            items = await items.NextPageRequest.GetAsync();
                    }
                    while (requestNextPage);
                }
            }

            return result;
        }

        /// <summary>
        /// Get all contacts in the current folder (or default root).
        /// </summary>
        /// <returns>Returns a list of contacts info</returns>
        /// 
        public async Task<ContactCollectionType> GetContactsAsync()
        {
            ContactCollectionType result = new ContactCollectionType();

            foreach (Microsoft.Graph.Contact item in await this.getContactsAsync())
            {
                ContactType tmpItem = new ContactType();
                tmpItem.OriginId = item.Id;
                tmpItem.OriginSystem = OriginSystemEnum.Outlook;
                tmpItem.LastModTime = item.LastModifiedDateTime.GetValueOrDefault().DateTime;

                tmpItem.DisplayName = item.DisplayName;
                tmpItem.Title = item.Title;
                tmpItem.Surname = item.Surname;
                tmpItem.MiddleName = item.MiddleName;
                tmpItem.GivenName = item.GivenName;
                tmpItem.AddName = item.AssistantName;
                tmpItem.Company = item.CompanyName;
                //tmpItem.VIP = 
                tmpItem.SaveAs = item.FileAs;
                tmpItem.Notes = item.PersonalNotes;

                result.Add(tmpItem);
            }

            return result;
        }

        /// <summary>
        /// Create a single contact by the specified info.
        /// </summary>
        /// <param name="item">New contact info.</param>
        /// <returns>Return the id of the new created contact.</returns>
        /// 
        public async Task<string> AddAsync(object item)
        {
            if (item is Microsoft.Graph.Contact)
            {
                bool retry = false;

                do
                {
                    try
                    {
                        string result;

                        //if the default root should be used
                        if (useDefault)
                            result = (await (parent as ContactFolders).connection.GetGraphClient().Me.Contacts.Request().AddAsync(item as Microsoft.Graph.Contact)).Id;
                        else
                            result = (await (parent as ContactFolders).connection.GetGraphClient().Me.ContactFolders[Id].Contacts.Request().AddAsync(item as Microsoft.Graph.Contact)).Id;

                        //if a add-event handler is defined
                        if (Event_Add != null)
                        {
                            await Task.Run(() => Event_Add(result));
                        }

                        return result;
                    }
                    catch (Microsoft.Graph.ServiceException e)
                    {
                        Debug.WriteLine("ERROR AddContactAsync " + e.Message);
                        retry = false;

                        //if service is busy, too many connection, aso. retry request
                        if (e.StatusCode.ToString() == "429")
                        {
                            int sleeptime = (parent as ContactFolders).connection.DefaultRetryDelay;

                            if (e.ResponseHeaders.RetryAfter != null)
                                sleeptime = e.ResponseHeaders.RetryAfter.Delta.GetValueOrDefault(new TimeSpan(0, 0, 1)).Milliseconds;

                            System.Threading.Thread.Sleep(sleeptime);
                            retry = true;
                        }
                    }
                } while (retry == true);
            }
            else if (item is ContactType)
            {
                return await addAsync(item as ContactType);
            }
            return null;
        }

        /// <summary>
        /// Create a single contact by the specified info.
        /// </summary>
        /// <param name="item">New contact info.</param>
        /// 
        public async Task<string> addAsync(ContactType item)
        {
            Microsoft.Graph.Contact newItem = new Microsoft.Graph.Contact();
            newItem.FileAs = item.SaveAs;
            newItem.Title = item.Title;
            newItem.DisplayName = item.DisplayName;
            newItem.GivenName = item.GivenName ?? "";
            newItem.MiddleName = item.MiddleName;
            newItem.Surname = item.Surname;
            newItem.Generation = item.AddName;
            newItem.CompanyName = item.Company;
            newItem.Birthday = item.Birthday;
            newItem.PersonalNotes = item.Notes;

            //if at least one messenger address is known
            if (item.IMAddress != null)
                newItem.ImAddresses = new List<string> { item.IMAddress };

            //if at least one category is set
            if (item.Categories != null)
                newItem.Categories = item.Categories.Select(x => x.Name);

            //if at least one private address
            if ((item.PrivateLocation.Street != null) || (item.PrivateLocation.Zip != null) || (item.PrivateLocation.City != null) || (item.PrivateLocation.Country != null))
            {
                newItem.HomeAddress = new PhysicalAddress();
                newItem.HomeAddress.Street = item.PrivateLocation.Street + " " + item.PrivateLocation.Number;
                newItem.HomeAddress.PostalCode = item.PrivateLocation.Zip;
                newItem.HomeAddress.City = item.PrivateLocation.City;
                newItem.HomeAddress.CountryOrRegion = item.PrivateLocation.Country;
            }

            //handle mail adresses
            List<EmailAddress> tmpEMailAddressInfo = new List<EmailAddress>();
            if (item.PrivateMailAddress.Address != null) tmpEMailAddressInfo.Add(new EmailAddress() { Address = item.PrivateMailAddress.Address, Name = item.PrivateMailAddress.Title });
            if (item.BusinessMailAddress.Address != null) tmpEMailAddressInfo.Add(new EmailAddress() { Address = item.BusinessMailAddress.Address, Name = item.BusinessMailAddress.Title });

            if (tmpEMailAddressInfo.Count > 0) newItem.EmailAddresses = tmpEMailAddressInfo;

            //handle phone numbers
            List<String> tmpBusinessPhoneNumberInfo = new List<String>();
            if (item.BusinessPhoneNumber != null) tmpBusinessPhoneNumberInfo.Add(item.BusinessPhoneNumber);
            if (item.BusinessMobileNumber != null) tmpBusinessPhoneNumberInfo.Add(item.BusinessMobileNumber);
            if (tmpBusinessPhoneNumberInfo.Count > 0) newItem.BusinessPhones = tmpBusinessPhoneNumberInfo;

            List<String> tmpPrivatePhoneNumberInfo = new List<String>();
            if (item.PrivatePhoneNumber != null) tmpPrivatePhoneNumberInfo.Add(item.PrivatePhoneNumber);
            if (tmpPrivatePhoneNumberInfo.Count > 0) newItem.HomePhones = tmpPrivatePhoneNumberInfo;

            newItem.MobilePhone = item.PrivateMobileNumber;

            //if at least one business address
            if ((item.BusinessLocation.Street != null) || (item.BusinessLocation.Zip != null) || (item.BusinessLocation.City != null) || (item.BusinessLocation.Country != null))
            {
                newItem.BusinessAddress = new PhysicalAddress();
                newItem.BusinessAddress.Street = item.BusinessLocation.Street + " " + item.BusinessLocation.Number;
                newItem.BusinessAddress.PostalCode = item.BusinessLocation.Zip;
                newItem.BusinessAddress.City = item.BusinessLocation.City;
                newItem.BusinessAddress.CountryOrRegion = item.BusinessLocation.Country;
            }

            string result = await this.AddAsync(newItem);

            if (item.HasPhoto)
            { await this.SetPhoto(result, item.PictureTmpFilename); }

            return result;
        }

        /// <summary>
        /// Create new contacts by a list of standardizes contacts (see "Contacts"-namespace)
        /// </summary>
        /// <param name="items">Standardized contact information</param>
        /// 
        public void AddContacts(ContactCollectionType items)
        {
            List<Task> tasks = new List<Task>();

            foreach (ContactType item in items)
            { tasks.Add(this.AddAsync(item)); }

            Task.WaitAll(tasks.ToArray());
        }

        /// <summary>
        /// Update the information of a contact.
        /// </summary>
        /// <param name="itemId">ID of the contact.</param>
        /// <param name="updatedItem">Info to update/new info</param>
        /// 
        public async Task UpdateAsync(string itemId, Microsoft.Graph.Contact updatedItem)
        {
            bool retry = false;

            do
            {
                try
                { await (parent as ContactFolders).connection.GetGraphClient().Me.Contacts[itemId].Request().UpdateAsync(updatedItem); }
                catch (Microsoft.Graph.ServiceException e)
                {
                    Debug.WriteLine("ERROR UpdateContactAsync " + e.Message);
                    retry = false;

                    //if service is busy, too many connection, aso. retry request
                    if (e.StatusCode.ToString() == "429")
                    {
                        int sleeptime = (parent as ContactFolders).connection.DefaultRetryDelay;

                        if (e.ResponseHeaders.RetryAfter != null)
                            sleeptime = e.ResponseHeaders.RetryAfter.Delta.GetValueOrDefault(new TimeSpan(0, 0, 1)).Milliseconds;

                        System.Threading.Thread.Sleep(sleeptime);
                        retry = true;
                    }
                }
            } while (retry == true);
        }

        public async Task SetPhoto(string itemId, string ImgFile)
        {
            System.IO.FileStream file = new System.IO.FileStream(ImgFile, System.IO.FileMode.Open);

            bool retry = false;

            do
            {
                try
                {
                    await (parent as ContactFolders).connection.GetGraphClient().Me.Contacts[itemId].Photo.Content.Request().PutAsync(file).ContinueWith((i) => { file.Dispose(); });
                }
                catch (Microsoft.Graph.ServiceException e)
                {
                    Debug.WriteLine("ERROR SendRequest " + e.Message);
                    retry = false;

                    //if service is busy, too many connection, aso. retry request
                    if (e.StatusCode.ToString() == "429")
                    {
                        int sleeptime = (parent as ContactFolders).connection.DefaultRetryDelay;

                        if (e.ResponseHeaders.RetryAfter != null)
                            sleeptime = e.ResponseHeaders.RetryAfter.Delta.GetValueOrDefault(new TimeSpan(0, 0, 1)).Milliseconds;

                        System.Threading.Thread.Sleep(sleeptime);
                        retry = true;
                    }
                }
            } while (retry);
        }

        /// <summary>
        /// Delete the specified contact.
        /// </summary>
        /// <param name="itemId">ID of the contact to delete</param>
        /// 
        public async Task<bool> RemoveAsync(string itemId)
        {
            bool retry = false;

            do
            {
                try
                {
                    await (parent as ContactFolders).connection.GetGraphClient().Me.Contacts[itemId].Request().DeleteAsync();
                    return true;
                }
                catch (Microsoft.Graph.ServiceException e)
                {
                    Debug.WriteLine("ERROR DeleteContactAsync " + e.Message);
                    retry = false;

                    //if service is busy, too many connection, aso. retry request
                    if (e.StatusCode.ToString() == "429")
                    {
                        int sleeptime = (parent as ContactFolders).connection.DefaultRetryDelay;

                        if (e.ResponseHeaders.RetryAfter != null)
                            sleeptime = e.ResponseHeaders.RetryAfter.Delta.GetValueOrDefault(new TimeSpan(0, 0, 1)).Milliseconds;

                        System.Threading.Thread.Sleep(sleeptime);
                        retry = true;
                    }
                }
            } while (retry == true);

            //not successfull
            return false;
        }

        /// <summary>
        /// Delete contacts which are contained in the IDs list
        /// </summary>
        /// <param name="IDs">List of IDs to delete</param>
        /// <returns>TRUE if successfull otherwise false</returns>
        /// 
        public bool Delete(List<string> IDs)
        {
            List<Task> deleteThreads = new List<Task>();

            try
            {
                foreach (string id in IDs)
                { deleteThreads.Add(this.RemoveAsync(id)); }

                Task.WaitAll(deleteThreads.ToArray());
                return true;
            }
            catch (Microsoft.Graph.ServiceException e)
            {
                Debug.WriteLine("ERROR DeleteContacts " + e.Message);
                return false;
            }
        }

        /// <summary>
        /// Delete all contacts in the current folder (or default root)
        /// </summary>
        /// <returns>TRUE if successfull otherwise FALSE</returns>
        /// 
        public bool Clear()
        {
            List<Task> deleteThreads = new List<Task>();

            try
            {
                List<Microsoft.Graph.Contact> items = this.getContactsAsync().Result;
                foreach (Microsoft.Graph.Contact item in items)
                { deleteThreads.Add(this.RemoveAsync(item.Id)); }

                Task.WaitAll(deleteThreads.ToArray());
                return true;
            }
            catch (Microsoft.Graph.ServiceException e)
            {
                Debug.WriteLine("ERROR ClearContacts " + e.Message);
                return false;
            }
        }
    }
}
