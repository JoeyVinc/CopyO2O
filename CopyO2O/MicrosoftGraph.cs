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


namespace CopyO2O.Office365
{
    public class MSGraph
    {
        private PublicClientApplication appClient;
        private GraphServiceClient _graphClient = null;
        private const string Authority = "https://login.microsoftonline.com/common/";
        private List<string> Scope = new List<string> { };

        public bool IsValid { get; } = false; //TRUE if the graph could be successfully connected
        public int MaxItemsPerRequest = 10;
        public int DefaultRetryDelay = 500;

        public MSGraph(string appId, List<string> scope)
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

        protected GraphServiceClient GetGraphClient()
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
    }

    public class Calendars : MSGraph
    {
        private Dictionary<string, string> _calendarIds = new Dictionary<string, string>();

        //constructor
        public Calendars(string appId, List<string> AppPermissions) : base(appId, AppPermissions) { }

        private string GetCalendarId(string CalendarName, bool RenewRequest = false)
        {
            if (!RenewRequest)
            {
                //if the calendarname is already known return the corresponding ID
                if (_calendarIds.ContainsKey(CalendarName.ToUpper()))
                    return _calendarIds[CalendarName.ToUpper()];
            }

            List<QueryOption> options = new List<QueryOption>();
            options.Add(new QueryOption("$top", MaxItemsPerRequest.ToString()));

            //request all calendars
            IUserCalendarsCollectionPage calendars = GetGraphClient().Me.Calendars.Request(options).GetAsync().Result;

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

                    return (await GetGraphClient().Me.Calendars.Request().AddAsync(newCal)).Id;
                }
                catch (Exception e)
                {
                    Debug.WriteLine("ERROR CreateAsync " + e.Message);
                    return "";
                }
            }
            else return "";
        }

        public Calendar GetCalendar(string CalendarName)
        {
            return new Calendar(GetGraphClient(), CalendarName, this.GetCalendarId(CalendarName));
        }

        /// <summary>
        /// delete calendar
        /// </summary>
        /// <param name="CalendarName">Name of calendar to delete</param>
        /// <returns>TRUE if successful otherwise FALSE</returns>
        public async Task<bool> DeleteAsync(string CalendarName)
        {
            string calendarID = GetCalendarId(CalendarName);

            //if calendar exists
            if (calendarID != "")
            {
                try
                {
                    await RenameAsync(CalendarName, "tmp_" + DateTime.Now.ToString("g") + ".tobedeleted");
                    await GetGraphClient().Me.Calendars[calendarID].Request().DeleteAsync();

                    _calendarIds.Remove(CalendarName.ToUpper());
                    return true;
                }
                catch (Microsoft.Graph.ServiceException e)
                {
                    Debug.WriteLine("ERROR DeleteAsync " + e.Message);

                    if (e.StatusCode == System.Net.HttpStatusCode.Conflict)
                    {
                        await GetGraphClient().Me.Calendars[calendarID].Request().DeleteAsync();

                        _calendarIds.Remove(CalendarName.ToUpper());
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// renames a calendar
        /// </summary>
        /// <param name="OldCalendarName">Original/old name of calendar to rename</param>
        /// <param name="NewCalendarName">new name</param>
        /// <param name="NameIsId">TRUE if OldCalendarName is the ID (not the name)</param>
        /// <returns></returns>
        public async Task<bool> RenameAsync(string OldCalendarName, string NewCalendarName)
        {
            string calendarID = GetCalendarId(OldCalendarName);

            //if calendar exists
            if (calendarID != "")
            {
                try
                {
                    Microsoft.Graph.Calendar tempCalendar = GetGraphClient().Me.Calendars[calendarID].Request().GetAsync().Result;
                    tempCalendar.Name = NewCalendarName;
                    await GetGraphClient().Me.Calendars[calendarID].Request().UpdateAsync(tempCalendar);

                    _calendarIds.Remove(OldCalendarName.ToUpper());
                    return true;
                }
                catch (Exception e)
                {
                    Debug.WriteLine("ERROR RenameAsync " + e.Message);
                    return false;
                }
            }
            else return false;
        }
    }

    public class Calendar
    {
        private string calendarName = "";
        private string calendarId = "";
        private GraphServiceClient graphService;
        private Dictionary<string, string> deltaSets = new Dictionary<string, string>();
        private int DefaultRetryDelay = 500;
        private int MaxItemsPerRequest = 10;

        public Calendar(GraphServiceClient GraphClient, string Name, string Id)
        {
            graphService = GraphClient;
            calendarName = Name;
            calendarId = Id;
        }

        public async Task<Microsoft.Graph.Event> GetItemAsync(string EventId)
        {
            return await graphService.Me.Calendars[calendarId].Events[EventId].Request().GetAsync();
        }

        public async Task<List<Microsoft.Graph.Event>> GetItemsAsync(DateTime From, DateTime To)
        {
            List<Microsoft.Graph.Event> result = new List<Microsoft.Graph.Event>();

            //request all events expanded (calendarview -> resolved recurring events)
            List<QueryOption> options = new List<QueryOption>();
            options.Add(new QueryOption("$top", MaxItemsPerRequest.ToString()));
            options.Add(new QueryOption("startDateTime", DateNTime.ConvertDateTimeToISO8016(From)));
            options.Add(new QueryOption("endDateTime", DateNTime.ConvertDateTimeToISO8016(To)));

            ICalendarCalendarViewCollectionPage items = await graphService.Me.Calendars[calendarId].CalendarView.Request(options).GetAsync();
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

            return result;
        }

        public async Task<List<Microsoft.Graph.Event>> GetItemsDeltaAsync(DateTime From, DateTime To, bool InitDelta = true)
        {
            List<Microsoft.Graph.Event> result = new List<Microsoft.Graph.Event>();

            //collect parameters for request
            List<QueryOption> options = new List<QueryOption>();
            options.Add(new QueryOption("odata.maxpagesize", MaxItemsPerRequest.ToString()));

            //if a new delta request should be sent
            if (InitDelta)
            {
                options.Add(new QueryOption("startDateTime", DateNTime.ConvertDateTimeToISO8016(From)));
                options.Add(new QueryOption("endDateTime", DateNTime.ConvertDateTimeToISO8016(To)));
            }
            //if an existing delta request should be resumed
            else
            {
                if (deltaSets.ContainsKey(calendarId) == false)
                    throw new ApplicationException("Error: Init delta run first!");

                options.Add(new QueryOption("$deltatoken", deltaSets[calendarId]));
            }

            //request all events expanded (calendarview -> resolved recurring events)
            IEventDeltaCollectionPage events = await graphService.Me.Calendars[calendarId].CalendarView.Delta().Request(options).GetAsync();

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
            if (!deltaSets.ContainsKey(calendarId))
                deltaSets.Add(calendarId, "");

            string deltaToken = System.Net.WebUtility.UrlDecode(events.AdditionalData["@odata.deltaLink"].ToString()).Split(new string[] { "$deltatoken=" }, StringSplitOptions.RemoveEmptyEntries)[1];
            deltaSets[calendarId] = deltaToken;

            return result;
        }

        public async Task<List<Microsoft.Graph.Event>> GetItemsDeltaAsync(string DeltaToken)
        {
            //save the deltaToken for the next (delta-) request
            if (!deltaSets.ContainsKey(calendarId))
                deltaSets.Add(calendarId, "");

            deltaSets[calendarId] = DeltaToken;

            return await this.GetItemsDeltaAsync(DateTime.Now, DateTime.Now, false);
        }

        public async Task<Microsoft.Graph.Event> CreateItemAsync(Microsoft.Graph.Event item)
        {
            bool retry = false;

            do
            {
                try
                { return await graphService.Me.Calendars[calendarId].Events.Request().AddAsync(item); }
                catch (Microsoft.Graph.ServiceException e)
                {
                    Debug.WriteLine("ERROR CreateItemAsync " + e.Message);
                    retry = false;

                    //if service is busy, too many connection, aso. retry request
                    if (e.StatusCode.ToString() == "429")
                    {
                        int sleeptime = DefaultRetryDelay;

                        if (e.ResponseHeaders.RetryAfter != null)
                            sleeptime = e.ResponseHeaders.RetryAfter.Delta.GetValueOrDefault(new TimeSpan(0, 0, 1)).Milliseconds;

                        System.Threading.Thread.Sleep(sleeptime);
                        retry = true;
                    }
                }
            } while (retry == true);

            return null;
        }

        public void CreateItems(Events items)
        {
            List<Task> tasks = new List<Task>();

            foreach (Event item in items)
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
                tasks.Add(this.CreateItemAsync(newEvent));
            }

            Task.WaitAll(tasks.ToArray());
        }

        public async Task UpdateItemAsync(string eventId, Microsoft.Graph.Event updatedItem)
        {
            bool retry = false;

            do
            {
                try
                { await graphService.Me.Events[eventId].Request().UpdateAsync(updatedItem); }
                catch (Microsoft.Graph.ServiceException e)
                {
                    Debug.WriteLine("ERROR UpdateItemAsync " + e.Message);
                    retry = false;

                    //if service is busy, too many connection, aso. retry request
                    if (e.StatusCode.ToString() == "429")
                    {
                        int sleeptime = DefaultRetryDelay;

                        if (e.ResponseHeaders.RetryAfter != null)
                            sleeptime = e.ResponseHeaders.RetryAfter.Delta.GetValueOrDefault(new TimeSpan(0, 0, 1)).Milliseconds;

                        System.Threading.Thread.Sleep(sleeptime);
                        retry = true;
                    }
                }
            } while (retry == true);
        }

        public async Task DeleteItemAsync(string eventId)
        {
            bool retry = false;

            do
            {
                try
                { await graphService.Me.Events[eventId].Request().DeleteAsync(); }
                catch (Microsoft.Graph.ServiceException e)
                {
                    Debug.WriteLine(e);
                    retry = false;

                    //if service is busy, too many connection, aso. retry request
                    if (e.StatusCode.ToString() == "429")
                    {
                        int sleeptime = DefaultRetryDelay;

                        if (e.ResponseHeaders.RetryAfter != null)
                            sleeptime = e.ResponseHeaders.RetryAfter.Delta.GetValueOrDefault(new TimeSpan(0, 0, 1)).Milliseconds;

                        System.Threading.Thread.Sleep(sleeptime);
                        retry = true;
                    }
                }
            } while (retry == true);
        }

        public bool DeleteItems(DateTime from, DateTime to)
        {
            List<Task> deleteThreads = new List<Task>();

            try
            {
                List<Microsoft.Graph.Event> items = this.GetItemsAsync(from, to).Result;
                foreach (Microsoft.Graph.Event item in items)
                { deleteThreads.Add(this.DeleteItemAsync(item.Id)); }

                Task.WaitAll(deleteThreads.ToArray());
                return true;
            }
            catch (Microsoft.Graph.ServiceException e)
            {
                Debug.WriteLine("ERROR DeleteItems " + e.Message);
                return false;
            }
        }
    }

    public class ContactFolders : MSGraph
    {
        private Dictionary<string, string> _contactFolderIds = new Dictionary<string, string>();

        //constructor
        public ContactFolders(string appId, List<string> AppPermissions) : base(appId, AppPermissions)
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
            options.Add(new QueryOption("$top", MaxItemsPerRequest.ToString()));

            //request all contactFolders
            IUserContactFoldersCollectionPage contactFolders = GetGraphClient().Me.ContactFolders.Request(options).GetAsync().Result;

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
        public ContactFolder GetContactFolder(string ContactFolderName)
        {
            return new ContactFolder(GetGraphClient(), ContactFolderName, this.GetId(ContactFolderName));
        }

        /// <summary>
        /// Create a Contact Folder (subfolder of the default root)
        /// </summary>
        /// <param name="ContactFolderName">Name of contactFolder to delete</param>
        /// <returns>the ID of the created contactFolder</returns>
        public async Task<string> AddAsync(string ContactFolderName)
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

                    return (await GetGraphClient().Me.ContactFolders.Request().AddAsync(newCal)).Id;
                }
                catch (Exception e)
                {
                    Debug.WriteLine("ERROR AddAsync " + e.Message);
                    return "";
                }
            }
            else return "";
        }

        /// <summary>
        /// delete contactFolder
        /// </summary>
        /// <param name="ContactFolderName">Name of contactFolder to delete</param>
        /// <returns>TRUE if successful otherwise FALSE</returns>
        public async Task<bool> DeleteAsync(string ContactFolderName)
        {
            string contactFolderID = GetId(ContactFolderName);

            //if contactFolder exists
            if (contactFolderID != "")
            {
                try
                {
                    await RenameAsync(ContactFolderName, "tmp_" + DateTime.Now.ToString("g") + ".tobedeleted");
                    await GetGraphClient().Me.ContactFolders[contactFolderID].Request().DeleteAsync();

                    _contactFolderIds.Remove(ContactFolderName.ToUpper());
                    return true;
                }
                catch (Microsoft.Graph.ServiceException e)
                {
                    Debug.WriteLine("ERROR DeleteAsync " + e.Message);

                    if (e.StatusCode == System.Net.HttpStatusCode.Conflict)
                    {
                        await GetGraphClient().Me.ContactFolders[contactFolderID].Request().DeleteAsync();

                        _contactFolderIds.Remove(ContactFolderName.ToUpper());
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// renames a contactFolder
        /// </summary>
        /// <param name="OldContactFolderName">Original/old name of contactFolder to rename</param>
        /// <param name="NewContactFolderName">new name</param>
        /// <param name="NameIsId">TRUE if OldContactFolderName is the ID (not the name)</param>
        /// <returns></returns>
        public async Task<bool> RenameAsync(string OldContactFolderName, string NewContactFolderName)
        {
            string contactFolderID = GetId(OldContactFolderName);

            //if contactFolder exists
            if (contactFolderID != "")
            {
                try
                {
                    Microsoft.Graph.ContactFolder tempContactFolder = GetGraphClient().Me.ContactFolders[contactFolderID].Request().GetAsync().Result;
                    tempContactFolder.DisplayName = NewContactFolderName;
                    await GetGraphClient().Me.ContactFolders[contactFolderID].Request().UpdateAsync(tempContactFolder);

                    _contactFolderIds.Remove(OldContactFolderName.ToUpper());
                    return true;
                }
                catch (Exception e)
                {
                    Debug.WriteLine("ERROR RenameAsync " + e.Message);
                    return false;
                }
            }
            else return false;
        }
    }

    public class ContactFolder
    {
        private string contactFolderName = "";
        private string contactFolderId = "";
        private bool useDefault = false;

        private GraphServiceClient graphService;
        private int MaxItemsPerRequest = 10;
        private int DefaultRetryDelay = 500;

        public string ContactFolderName { get => UseDefault ? null : contactFolderName; }
        public string ContactFolderId { get => UseDefault ? null : contactFolderId; }
        public bool UseDefault { get { return useDefault; } set { useDefault = value; contactFolderName = null; contactFolderId = null; } }

        /// <summary>
        /// constructor of ContactFolder. ContactFolder is a SUBfolder for contacts. There is also a default folder (*hrgh*) which does not have a real name.
        /// </summary>
        /// <param name="GraphClient"></param>
        /// <param name="Name">Display name of a folder. If empty or null the default folder will be used.</param>
        /// <param name="Id">ID of the folder</param>
        /// 
        public ContactFolder(GraphServiceClient GraphClient, string Name, string Id)
        {
            graphService = GraphClient;

            //if not default
            if ((Name ?? "") != "")
            {
                contactFolderName = Name;
                contactFolderId = Id;
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
        public Microsoft.Graph.Contact this[string ItemId]
        {
            get
            {
                bool retry = false;

                do
                {
                    try
                    {
                        //if default folder
                        if (this.useDefault)
                            return graphService.Me.Contacts[ItemId].Request().GetAsync().Result;
                        else
                            return graphService.Me.ContactFolders[contactFolderId].Contacts[ItemId].Request().GetAsync().Result;
                    }
                    catch (Microsoft.Graph.ServiceException e)
                    {
                        Debug.WriteLine("ERROR ContactFolder THIS " + e.Message);
                        retry = false;

                        //if service is busy, too many connection, aso. retry request
                        if (e.StatusCode.ToString() == "429")
                        {
                            int sleeptime = DefaultRetryDelay;

                            if (e.ResponseHeaders.RetryAfter != null)
                                sleeptime = e.ResponseHeaders.RetryAfter.Delta.GetValueOrDefault(new TimeSpan(0, 0, 1)).Milliseconds;

                            System.Threading.Thread.Sleep(sleeptime);
                            retry = true;
                        }
                    }
                } while (retry == true);

                return null;                
            }
        }

        /// <summary>
        /// Get all contacts in the current folder (or default root).
        /// </summary>
        /// <returns>Returns a list of contacts info</returns>
        /// 
        public async Task<List<Microsoft.Graph.Contact>> GetContactsAsync()
        {
            List<Microsoft.Graph.Contact> result = new List<Microsoft.Graph.Contact>();

            List<QueryOption> options = new List<QueryOption>();
            options.Add(new QueryOption("$top", MaxItemsPerRequest.ToString()));

            //if default folder
            if (useDefault)
            {
                IUserContactsCollectionPage items = await graphService.Me.Contacts.Request(options).GetAsync();

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
                IContactFolderContactsCollectionPage items = await graphService.Me.ContactFolders[contactFolderId].Contacts.Request(options).GetAsync();

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
        /// Create a single contact by the specified info.
        /// </summary>
        /// <param name="item">New contact info.</param>
        /// <returns>Return the new created contact.</returns>
        /// 
        public async Task<Microsoft.Graph.Contact> AddAsync(Microsoft.Graph.Contact item)
        {
            bool retry = false;

            do
            {
                try
                {
                    //if the default root should be used
                    if (useDefault)
                        return await graphService.Me.Contacts.Request().AddAsync(item);
                    else
                        return await graphService.Me.ContactFolders[contactFolderId].Contacts.Request().AddAsync(item);
                }
                catch (Microsoft.Graph.ServiceException e)
                {
                    Debug.WriteLine("ERROR CreateContactAsync " + e.Message);
                    retry = false;

                    //if service is busy, too many connection, aso. retry request
                    if (e.StatusCode.ToString() == "429")
                    {
                        int sleeptime = DefaultRetryDelay;

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
        /// Create new contacts by a list of standardizes contacts (see "Contacts"-namespace)
        /// </summary>
        /// <param name="items">Standardized contact information</param>
        /// 
        public void AddContacts(ContactCollectionType items)
        {
            List<Task> tasks = new List<Task>();
            List<Task> pictureUploads = new List<Task>();

            foreach (ContactType item in items)
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

                tasks.Add(this.AddAsync(newItem).ContinueWith(async (i) =>
                    {
                        if (item.HasPhoto)
                        {
                            System.IO.FileStream file = new System.IO.FileStream(item.PictureTmpFilename, System.IO.FileMode.Open);
                            await graphService.Me.Contacts[i.Result.Id].Photo.Content.Request().PutAsync(file);
                        }
                    }));
            }

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
                { await graphService.Me.Contacts[itemId].Request().UpdateAsync(updatedItem); }
                catch (Microsoft.Graph.ServiceException e)
                {
                    Debug.WriteLine("ERROR UpdateContactAsync " + e.Message);
                    retry = false;

                    //if service is busy, too many connection, aso. retry request
                    if (e.StatusCode.ToString() == "429")
                    {
                        int sleeptime = DefaultRetryDelay;

                        if (e.ResponseHeaders.RetryAfter != null)
                            sleeptime = e.ResponseHeaders.RetryAfter.Delta.GetValueOrDefault(new TimeSpan(0, 0, 1)).Milliseconds;

                        System.Threading.Thread.Sleep(sleeptime);
                        retry = true;
                    }
                }
            } while (retry == true);
        }

        /// <summary>
        /// Delete the specified contact.
        /// </summary>
        /// <param name="itemId">ID of the contact to delete</param>
        /// 
        public async Task DeleteAsync(string itemId)
        {
            bool retry = false;

            do
            {
                try
                { await graphService.Me.Contacts[itemId].Request().DeleteAsync(); }
                catch (Microsoft.Graph.ServiceException e)
                {
                    Debug.WriteLine("ERROR DeleteContactAsync " + e.Message);
                    retry = false;

                    //if service is busy, too many connection, aso. retry request
                    if (e.StatusCode.ToString() == "429")
                    {
                        int sleeptime = DefaultRetryDelay;

                        if (e.ResponseHeaders.RetryAfter != null)
                            sleeptime = e.ResponseHeaders.RetryAfter.Delta.GetValueOrDefault(new TimeSpan(0, 0, 1)).Milliseconds;

                        System.Threading.Thread.Sleep(sleeptime);
                        retry = true;
                    }
                }
            } while (retry == true);
        }

        /// <summary>
        /// Delete all contacts in the current folder (or default root)
        /// </summary>
        /// <returns>TRUE if successfull otherwise FALSE</returns>
        /// 
        public bool DeleteContacts()
        {
            List<Task> deleteThreads = new List<Task>();

            try
            {
                List<Microsoft.Graph.Contact> items = this.GetContactsAsync().Result;
                foreach (Microsoft.Graph.Contact item in items)
                { deleteThreads.Add(this.DeleteAsync(item.Id)); }

                Task.WaitAll(deleteThreads.ToArray());
                return true;
            }
            catch (Microsoft.Graph.ServiceException e)
            {
                Debug.WriteLine("ERROR DeleteContacts " + e.Message);
                return false;
            }
        }
    }
}
