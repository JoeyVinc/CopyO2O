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
                IsValid = false;
                throw e;
            }
        }

        private async Task<AuthenticationResult> AuthResultAsync()
        {
            try { return await appClient.AcquireTokenSilentAsync(Scope, appClient.GetAccountsAsync().Result.First()); }
            catch { return await appClient.AcquireTokenAsync(Scope); }
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

            return "";
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
                catch { return ""; }
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
                catch { return false; }
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

        public int MaxItemsPerRequest = 10;

        public Calendar(GraphServiceClient GraphClient, string Name, string Id)
        {
            graphService = GraphClient;
            calendarName = Name;
            calendarId = Id;
        }

        public async Task<Microsoft.Graph.Event> GetItemAsync(string EventId)
        {
            if (calendarId != "")
                return await graphService.Me.Calendars[calendarId].Events[EventId].Request().GetAsync();
            else return null;
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
            //if calendar exists
            if (calendarId != "")
            {
                try
                {
                    return await graphService.Me.Calendars[calendarId].Events.Request().AddAsync(item);
                }
                catch { return null; }
            }
            else { return null; }
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

            int i = 0;
            while (tasks.Count > 0)
            {
                if (tasks[i].Status == TaskStatus.RanToCompletion)
                    tasks.RemoveAt(i);

                i++;
                if (i >= tasks.Count) i = 0;
            }
        }

        public async Task UpdateItemAsync(string eventId, Microsoft.Graph.Event updatedItem)
        {
            await graphService.Me.Events[eventId].Request().UpdateAsync(updatedItem);
        }

        public async Task DeleteItemAsync(string eventId)
        {
            await graphService.Me.Events[eventId].Request().DeleteAsync();
        }

        public async Task<bool> DeleteItemsAsync(DateTime from, DateTime to)
        {
            //if calendar exists
            if (calendarId != "")
            {
                try
                {
                    List<Microsoft.Graph.Event> items = this.GetItemsAsync(from, to).Result;
                    foreach (Microsoft.Graph.Event item in items)
                    { await this.DeleteItemAsync(item.Id); }

                    return true;
                }
                catch { return false; }
            }
            else return false;
        }
    }

    public class ContactFolders : MSGraph
    {
        private Dictionary<string, string> _contactFolderIds = new Dictionary<string, string>();

        //constructor
        public ContactFolders(string appId, List<string> AppPermissions) : base(appId, AppPermissions) { }

        private string GetContactFolderId(string ContactFolderName, bool RenewRequest = false)
        {
            if (!RenewRequest)
            {
                //if the contactFoldername is already known return the corresponding ID
                if (_contactFolderIds.ContainsKey(ContactFolderName.ToUpper()))
                    return _contactFolderIds[ContactFolderName.ToUpper()];
            }

            List<QueryOption> options = new List<QueryOption>();
            options.Add(new QueryOption("$top", MaxItemsPerRequest.ToString()));

            //request all contactFolders
            IUserContactFoldersCollectionPage contactFolders = GetGraphClient().Me.ContactFolders.Request(options).GetAsync().Result;

            bool requestNextPage;
            do
            {
                requestNextPage = (contactFolders.NextPageRequest != null);

                //search for id of specified calender
                for (int i = 0; i < contactFolders.Count; i++)
                {
                    //if the contactFolder was found return the ID
                    if (contactFolders[i].DisplayName.ToUpper() == ContactFolderName.ToUpper())
                    {
                        _contactFolderIds.Add(ContactFolderName.ToUpper(), contactFolders[i].Id);
                        return contactFolders[i].Id;
                    }
                }

                if (requestNextPage)
                    contactFolders = contactFolders.NextPageRequest.GetAsync().Result;

            } while (requestNextPage);

            return "";
        }

        /// <summary>
        /// create contactFolder
        /// </summary>
        /// <param name="ContactFolderName">Name of contactFolder to delete</param>
        /// <returns>the ID of the created contactFolder</returns>
        public async Task<string> CreateAsync(string ContactFolderName)
        {
            //if contactFolder does not exist
            if (GetContactFolderId(ContactFolderName, true) == "")
            {
                try
                {
                    Microsoft.Graph.ContactFolder newCal = new Microsoft.Graph.ContactFolder
                    {
                        DisplayName = ContactFolderName,
                    };

                    return (await GetGraphClient().Me.ContactFolders.Request().AddAsync(newCal)).Id;
                }
                catch { return ""; }
            }
            else return "";
        }

        public ContactFolder GetContactFolder(string ContactFolderName)
        {
            return new ContactFolder(GetGraphClient(), ContactFolderName, this.GetContactFolderId(ContactFolderName));
        }

        /// <summary>
        /// delete contactFolder
        /// </summary>
        /// <param name="ContactFolderName">Name of contactFolder to delete</param>
        /// <returns>TRUE if successful otherwise FALSE</returns>
        public async Task<bool> DeleteAsync(string ContactFolderName)
        {
            string contactFolderID = GetContactFolderId(ContactFolderName);

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
            string contactFolderID = GetContactFolderId(OldContactFolderName);

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
                catch { return false; }
            }
            else return false;
        }
    }

    public class ContactFolder
    {
        private string contactFolderName = "";
        private string contactFolderId = "";
        private GraphServiceClient graphService;
        private Dictionary<string, string> deltaSets = new Dictionary<string, string>();

        public int MaxItemsPerRequest = 10;

        public ContactFolder(GraphServiceClient GraphClient, string Name, string Id)
        {
            graphService = GraphClient;
            contactFolderName = Name;
            contactFolderId = Id;
        }

        public async Task<Microsoft.Graph.Contact> GetItemAsync(string ItemId)
        {
            if (contactFolderId != "")
                return await graphService.Me.ContactFolders[contactFolderId].Contacts[ItemId].Request().GetAsync();
            else return null;
        }

        public async Task<List<Microsoft.Graph.Contact>> GetItemsAsync()
        {
            List<Microsoft.Graph.Contact> result = new List<Microsoft.Graph.Contact>();

            //request all events expanded (contactFolderview -> resolved recurring events)
            List<QueryOption> options = new List<QueryOption>();
            options.Add(new QueryOption("$top", MaxItemsPerRequest.ToString()));

            IContactFolderContactsCollectionPage items = await graphService.Me.ContactFolders[contactFolderId].Contacts.Request(options).GetAsync();
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

            return result;
        }

        public async Task<Microsoft.Graph.Contact> CreateItemAsync(Microsoft.Graph.Contact item)
        {
            //if contactFolder exists
            if (contactFolderId != "")
            {
                try
                {
                    return await graphService.Me.ContactFolders[contactFolderId].Contacts.Request().AddAsync(item);
                }
                catch { return null; }
            }
            else { return null; }
        }

        public void CreateItems(ContactsType items)
        {
            List<Task> tasks = new List<Task>();
            List<Task> pictureUploads = new List<Task>();

            foreach (ContactType item in items)
            {
                Microsoft.Graph.Contact newItem = new Microsoft.Graph.Contact();
                newItem.FileAs = item.SaveAs;
                newItem.DisplayName = item.DisplayName;
                newItem.GivenName = item.GivenName ?? "";
                newItem.Surname = item.Surname;
                newItem.CompanyName = item.Company;
                newItem.Birthday = item.Birthday;
                newItem.PersonalNotes = item.Notes;

                //if at least one category is set
                if (item.Categories != null)
                    newItem.Categories = item.Categories.Select(x => x.Name);

                //if at least one private address
                if ((item.PrivateLocation.Street != null)||(item.PrivateLocation.Zip != null) || (item.PrivateLocation.City != null) || (item.PrivateLocation.Country != null))
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

                    if (tmpEMailAddressInfo.Count > 0)                        newItem.EmailAddresses = tmpEMailAddressInfo;

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

/*                //add picture if available
                if (item.PictureTmpFilename != null)
                {
                    //newItem.Photo = UpdatePhoto(newItem);
                    graphService.Me.Contacts[9].Photo.Content.Request().p
                }
                */
                tasks.Add(this.CreateItemAsync(newItem));
            }

            int i = 0;
            while (tasks.Count > 0)
            {
                if (tasks[i].Status == TaskStatus.RanToCompletion)
                    tasks.RemoveAt(i);

                i++;
                if (i >= tasks.Count) i = 0;
            }
        }

        public async Task UpdateItemAsync(string itemId, Microsoft.Graph.Contact updatedItem)
        {
            await graphService.Me.Contacts[itemId].Request().UpdateAsync(updatedItem);
        }

        public async Task DeleteItemAsync(string itemId)
        {
            await graphService.Me.Contacts[itemId].Request().DeleteAsync();
        }

        public async Task<bool> DeleteItemsAsync()
        {
            //if contactFolder exists
            if (contactFolderId != "")
            {
                try
                {
                    List<Microsoft.Graph.Contact> items = this.GetItemsAsync().Result;
                    foreach (Microsoft.Graph.Contact item in items)
                    { await this.DeleteItemAsync(item.Id); }

                    return true;
                }
                catch { return false; }
            }
            else return false;
        }
    }

    public class Contact
    {
        //private string contactName = "";
        private string contactId = "";
        private GraphServiceClient graphService;

        public Contact (GraphServiceClient GraphService, string ContactID)
        {
            this.contactId = ContactID;
            this.graphService = GraphService;
        }

        public async Task UpdatePhotoAsync(string FileName)
        {
            ProfilePhoto tmpPhoto = new ProfilePhoto();
            tmpPhoto.AdditionalData.Add("image/jpeg", "poing");
            await graphService.Me.Contacts[contactId].Photo.Request().UpdateAsync(tmpPhoto);
        }
        
    }

}
