using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using OutlookInterop = Microsoft.Office.Interop.Outlook;

namespace CopyO2O.Outlook
{
    public class Calendar
    {
        private OutlookInterop.MAPIFolder mAPI;
        public string Name;

        public Calendar(OutlookInterop.MAPIFolder folder)
        {
            mAPI = folder;
            Name = folder.Name;
        }

        ~Calendar()
        {
            this.Free();
        }

        public void Free()
        {
            mAPI = null;
        }

        public Events GetItems(DateTime from, DateTime to)
        {
            OutlookInterop.Items tempEvents;
            tempEvents = mAPI.Items;
            tempEvents.IncludeRecurrences = true;
            tempEvents.Sort("[Start]");

            string filter = "[Start] >= '" + from.ToString("g") + "'"
                + " AND " + "[Start] <= '" + to.ToString("g") + "'";
            OutlookInterop.Items eventsFiltered = tempEvents.Restrict(filter);

            Events result = new Events();
            foreach (OutlookInterop.AppointmentItem aptmItem in eventsFiltered)
            {
                Event tmpEvent = new Event();
                tmpEvent.OriginId = aptmItem.EntryID + "_" + aptmItem.Start.ToFileTimeUtc().ToString() + aptmItem.Duration.ToString();
                tmpEvent.LastModTime = aptmItem.LastModificationTime;

                tmpEvent.Subject = aptmItem.Subject;
                tmpEvent.StartDateTime = aptmItem.Start;
                tmpEvent.StartTimeZone = aptmItem.StartTimeZone.ID;
                tmpEvent.StartUTC = aptmItem.StartUTC;
                tmpEvent.EndDateTime = aptmItem.End;
                tmpEvent.EndTimeZone = aptmItem.EndTimeZone.ID;
                tmpEvent.EndUTC = aptmItem.EndUTC;
                tmpEvent.AllDayEvent = aptmItem.AllDayEvent;

                tmpEvent.Location = aptmItem.Location;
                tmpEvent.Subject = aptmItem.Subject;
                tmpEvent.Body = aptmItem.Body;
                tmpEvent.ReminderMinutesBefore = aptmItem.ReminderMinutesBeforeStart;
                tmpEvent.ReminderOn = (aptmItem.ReminderSet && (tmpEvent.ReminderMinutesBefore >= 0));
                tmpEvent.IsPrivate = (aptmItem.Sensitivity == OutlookInterop.OlSensitivity.olPrivate);

                switch (aptmItem.Importance)
                {
                    case OutlookInterop.OlImportance.olImportanceHigh: tmpEvent.Importance = Event.ImportanceEnum.High; break;
                    case OutlookInterop.OlImportance.olImportanceLow: tmpEvent.Importance = Event.ImportanceEnum.Low; break;
                    case OutlookInterop.OlImportance.olImportanceNormal: tmpEvent.Importance = Event.ImportanceEnum.Normal; break;
                    default: tmpEvent.Importance = Event.ImportanceEnum.Normal; break;
                }

                switch (aptmItem.BusyStatus)
                {
                    case OutlookInterop.OlBusyStatus.olFree: tmpEvent.Status = Event.StatusEnum.Free; break;
                    case OutlookInterop.OlBusyStatus.olTentative: tmpEvent.Status = Event.StatusEnum.Tentative; break;
                    case OutlookInterop.OlBusyStatus.olBusy: tmpEvent.Status = Event.StatusEnum.Busy; break;
                    case OutlookInterop.OlBusyStatus.olOutOfOffice: tmpEvent.Status = Event.StatusEnum.OutOfOffice; break;
                    case OutlookInterop.OlBusyStatus.olWorkingElsewhere: tmpEvent.Status = Event.StatusEnum.ElseWhere; break;
                    default: tmpEvent.Status = Event.StatusEnum.Tentative; break;
                }

                result.Add(tmpEvent);
            }
            return result;
        }

        public void CreateItems(Events newItems)
        {
            foreach (Event item in newItems)
            {
                OutlookInterop.AppointmentItem newEvent = mAPI.Items.Add();
                newEvent.StartTimeZone = mAPI.Application.TimeZones[item.StartTimeZone];
                newEvent.StartUTC = item.StartUTC.Value;
                newEvent.EndTimeZone = mAPI.Application.TimeZones[item.EndTimeZone];
                newEvent.EndUTC = item.EndUTC.Value;

                newEvent.Subject = item.Subject;
                newEvent.Location = item.Location;
                newEvent.AllDayEvent = item.AllDayEvent;

                newEvent.ReminderMinutesBeforeStart = item.ReminderMinutesBefore;
                newEvent.ReminderSet = item.ReminderOn;

                switch (item.Status)
                {
                    case Event.StatusEnum.Free: newEvent.BusyStatus = OutlookInterop.OlBusyStatus.olFree; break;
                    case Event.StatusEnum.Tentative: newEvent.BusyStatus = OutlookInterop.OlBusyStatus.olTentative; break;
                    case Event.StatusEnum.Busy: newEvent.BusyStatus = OutlookInterop.OlBusyStatus.olBusy; break;
                    case Event.StatusEnum.OutOfOffice: newEvent.BusyStatus = OutlookInterop.OlBusyStatus.olOutOfOffice; break;
                    case Event.StatusEnum.ElseWhere: newEvent.BusyStatus = OutlookInterop.OlBusyStatus.olWorkingElsewhere; break;
                    default: newEvent.BusyStatus = OutlookInterop.OlBusyStatus.olFree; break;
                }

                newEvent.Save();
            }
        }

        public void DeleteItems(DateTime from, DateTime to)
        {
            OutlookInterop.Items tempEvents;
            tempEvents = mAPI.Items;
            tempEvents.IncludeRecurrences = true;
            tempEvents.Sort("[Start]");

            string filter = "[Start] >= '" + from.ToString("g") + "'"
                + " AND " + "[Start] <= '" + to.ToString("g") + "'";
            OutlookInterop.Items eventsFiltered = tempEvents.Restrict(filter);

            int count = 0;
            foreach (OutlookInterop.AppointmentItem tmp in eventsFiltered) count++;

            for (int index = count; index >= 1; index--)
            {
                eventsFiltered[index].Delete();
            }
        }

        public void DeleteAllItems()
        {
            for (int index = mAPI.Items.Count; index >= 1; index--)
            {
                mAPI.Items[index].delete();
            }
        }
    }

    public class ContactFolder
    {
        private OutlookInterop.MAPIFolder mAPI;
        public string Name;

        public ContactFolder(OutlookInterop.MAPIFolder folder)
        {
            mAPI = folder;
            Name = folder.Name;
        }

        ~ContactFolder()
        {
            this.Free();
        }

        public void Free()
        {
            mAPI = null;
        }

        private EMail GetMailStruct(string displayName, string address, string type)
        {
            EMail result = new EMail();

            //if no valid mail data exit without calc
            if ((type ?? "") == "")
                return result;

            result.Title = displayName;
            result.Address = null;

            //if SMTP-address is given
            if (type.ToUpper() == "SMTP")
            {
                result.Address = address;
            }
            //if EXCHANGE-address
            else if (type.ToUpper() == "EX")
            {
                Microsoft.Office.Interop.Outlook.MailItem tmpMail = null;
                try
                {
                    tmpMail = mAPI.Application.CreateItem(OutlookInterop.OlItemType.olMailItem);
                    tmpMail.To = address;
                    if (tmpMail.Recipients.ResolveAll())
                    {
                        foreach (OutlookInterop.Recipient tmpRcip in tmpMail.Recipients)
                        {
                            OutlookInterop.ExchangeUser exchangeUser = tmpRcip.AddressEntry.GetExchangeUser();
                            if (exchangeUser != null)
                                result.Address = exchangeUser.PrimarySmtpAddress;
                        }
                    }
                }
                catch (Exception e)
                {
                    Debug.WriteLine("ERROR GetMailStruct " + e.Message);
                }
                finally
                {
                    if (tmpMail != null)
                        tmpMail.Close(OutlookInterop.OlInspectorClose.olDiscard);
                    tmpMail = null;
                }
            }

            return result;
        }

        public ContactCollectionType GetItems()
        {
            ContactCollectionType result = new ContactCollectionType();
            foreach (OutlookInterop.ContactItem item in mAPI.Items.OfType<OutlookInterop.ContactItem>())
            {
                ContactType tmpItem = new ContactType();
                tmpItem.OriginId = item.EntryID;
                tmpItem.OriginSystem = OriginSystemEnum.Outlook;
                tmpItem.LastModTime = item.LastModificationTime;

                tmpItem.DisplayName = item.FullName;
                tmpItem.Title = item.Title;
                tmpItem.Surname = item.LastName;
                tmpItem.MiddleName = item.MiddleName;
                tmpItem.GivenName = item.FirstName;
                tmpItem.AddName = item.Suffix;
                tmpItem.Company = item.CompanyName;
                tmpItem.VIP = (item.Importance == OutlookInterop.OlImportance.olImportanceHigh);
                tmpItem.SaveAs = item.FileAs;
                if (item.Birthday.Year < 4000) tmpItem.Birthday = item.Birthday;
                if (item.Anniversary.Year < 4000) tmpItem.AnniversaryDay = item.Anniversary;
                tmpItem.Notes = item.Body;
                tmpItem.IMAddress = item.IMAddress;

                //==========================
                //private data
                tmpItem.PrivateMailAddress = GetMailStruct(item.Email1DisplayName, item.Email1Address, item.Email1AddressType);
                //if no valid mail address could be delivered
                if ((tmpItem.PrivateMailAddress.Address ?? "") == "")
                    tmpItem.PrivateMailAddress.Address = (tmpItem.IMAddress ?? "").Contains('@') ? tmpItem.IMAddress : null;

                tmpItem.PrivateMobileNumber = item.MobileTelephoneNumber;
                tmpItem.PrivatePhoneNumber = item.HomeTelephoneNumber;
                tmpItem.PrivateFaxNumber = item.HomeFaxNumber;

                if ((item.HomeAddressStreet != null) || (item.HomeAddressPostalCode != null) || (item.HomeAddressCity != null) || (item.HomeAddressCountry != null))
                {
                    //tmpItem.PrivateLocation = new Address();
                    tmpItem.PrivateLocation.Street = item.HomeAddressStreet?.Split(' ').First();
                    tmpItem.PrivateLocation.Number = tmpItem.PrivateLocation.Street?.Substring(tmpItem.PrivateLocation.Street.Length - 1);
                    tmpItem.PrivateLocation.Zip = item.HomeAddressPostalCode;
                    tmpItem.PrivateLocation.City = item.HomeAddressCity;
                    tmpItem.PrivateLocation.Country = item.HomeAddressCountry;
                }

                //==========================
                //business data
                tmpItem.BusinessMailAddress = GetMailStruct(item.Email2DisplayName, item.Email2Address, item.Email2AddressType);
                //if no valid mail address could be delivered
                if ((tmpItem.BusinessMailAddress.Address ?? "") == "")
                    tmpItem.BusinessMailAddress.Address = (tmpItem.IMAddress ?? "").Contains('@') ? tmpItem.IMAddress : null;

                tmpItem.BusinessMobileNumber = item.Business2TelephoneNumber;
                tmpItem.BusinessPhoneNumber = item.BusinessTelephoneNumber;
                tmpItem.BusinessFaxNumber = item.BusinessFaxNumber;

                if ((item.BusinessAddressStreet != null) || (item.BusinessAddressPostalCode != null) || (item.BusinessAddressCity != null) || (item.BusinessAddressCountry != null))
                {
                    //tmpItem.BusinessLocation = new Address();
                    tmpItem.BusinessLocation.Street = item.BusinessAddressStreet?.Split(' ').First();
                    tmpItem.BusinessLocation.Number = tmpItem.BusinessLocation.Street?.Substring(tmpItem.BusinessLocation.Street.Length - 1);
                    tmpItem.BusinessLocation.Zip = item.BusinessAddressPostalCode;
                    tmpItem.BusinessLocation.City = item.BusinessAddressCity;
                    tmpItem.BusinessLocation.Country = item.BusinessAddressCountry;
                }

                //==========================
                //handle photo
                if (item.HasPicture)
                {
                    OutlookInterop.Attachment tmpPhotofile = item.Attachments["ContactPicture.jpg"];
                    //if a photo is attached
                    if (tmpPhotofile != null)
                    {
                        string tmpFilename = Environment.GetEnvironmentVariable("TEMP").TrimEnd('\\') + '\\' + item.EntryID + ".jpg";

                        tmpPhotofile.SaveAsFile(tmpFilename);
                        tmpItem.PictureTmpFilename = tmpFilename;
                    }
                    Marshal.ReleaseComObject(tmpPhotofile);
                }

                result.Add(tmpItem);
            }
            return result;
        }
    }

    public class Application
    {
        OutlookInterop.Application appInstance;
        private bool outlookAlreadyRunning = false;
        private bool exitOutlookOnQuit = true;
        private List<Calendar> alreadyOpenedCalendars = new List<Calendar>(); //for faster access and controlled destroy of the calendar objects
        private List<ContactFolder> alreadyOpenedContactFolders = new List<ContactFolder>(); //for faster access and controlled destroy of the contactfolder objects

        public Application(bool exitOutlookOnQuit = true)
        {
            this.exitOutlookOnQuit = exitOutlookOnQuit;

            try
            {
                appInstance = (OutlookInterop.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Outlook.Application");
                outlookAlreadyRunning = true;
            }
            catch (Exception e)
            {
                Debug.WriteLine("ERROR OutlookApplication " + e.Message);
                appInstance = new OutlookInterop.Application();
            }
        }

        ~Application()
        {
            this.Quit();
        }

        public void Quit()
        {
            if (appInstance != null)
            {
                appInstance.Session.SendAndReceive(false);

                //if outlook was not already running
                if (!outlookAlreadyRunning && exitOutlookOnQuit)
                {
                    for (int i = 0; i < alreadyOpenedCalendars.Count; i++)
                    { alreadyOpenedCalendars[i].Free(); }

                    appInstance.Quit();
                    appInstance = null;

                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
        }

        public Calendar GetCalendar(string calendarPath)
        {
            Calendar result = alreadyOpenedCalendars.Find(x => x.Name == calendarPath);

            try
            {
                //if the calendar could not be found
                if (result == null)
                {
                    string[] names = calendarPath.Split('\\');
                    OutlookInterop.MAPIFolder tmpFolder = appInstance.Session.Folders[names[0]];

                    foreach (string name in names.Skip(1))
                    {
                        tmpFolder = tmpFolder.Folders[name];
                    }

                    result = new Calendar(tmpFolder);
                    this.alreadyOpenedCalendars.Add(result);
                }
                return result;
            }
            catch (Exception e)
            {
                Debug.WriteLine("ERROR GetCalendar " + e.Message);
                throw new Exception(calendarPath + " could not be found!");
            }
        }

        public ContactFolder GetContactFolder(string contactFolderPath)
        {
            ContactFolder result = alreadyOpenedContactFolders.Find(x => x.Name == contactFolderPath);

            //if the calendar could not be found
            if (result == null)
            {
                string[] names = contactFolderPath.Split('\\');
                OutlookInterop.MAPIFolder tmpFolder = appInstance.Session.Folders[names[0]];

                foreach (string name in names.Skip(1))
                {
                    tmpFolder = tmpFolder.Folders[name];
                }

                result = new ContactFolder(tmpFolder);
                this.alreadyOpenedContactFolders.Add(result);
            }

            return result;
        }
    }
}
