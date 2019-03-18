using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Commonfunctions.Cryptography;

namespace CopyO2O
{
    public class RecurrencePattern
    {
        public enum RecurrenceTypeEnum { Daily, Weekly, Monthly, Yearly }
        public enum DaysEnum { Monday = 1, Tuesday = 2, Wednesday = 3, Thursday = 4, Friday = 5, Saturday = 6, Sunday = 7 }
        public enum MonthsEnum { January = 1, February = 2, March = 3, April = 4, May = 5, June = 6, July = 7, August = 8, September = 9, October = 10, November = 11, December = 12 };

        public RecurrenceTypeEnum RecurrenceType;

        //series occurrence
        public DateTime FirstOccurenceDate = DateTime.MinValue;
        public string TimeZone = "UTC";
        public bool? NeverEnds = null;
        public DateTime? LastOccurenceDate = null;
        public int? Count = null;

        //N-th recurrence
        public int EveryNth; // e.g. every 3rd day or 2nd saturday or 3rd month or 6th year ...
        public int WeekdayInstance; // e.g. every second saturday, every last week of a month

        //in case of weekly recurrence
        public List<DaysEnum> DaysOfWeek = new List<DaysEnum>();

        //in case of monthly recurrence
        public List<int> DaysOfMonth = new List<int>();

        //in case of yearly recurrence
        public List<MonthsEnum> MonthsOfYear = new List<MonthsEnum>();

        //public static DaysEnum GetDayEnum(int day) { return (DaysEnum)day; }
        public static int GetDayInt(DaysEnum day) { return (int)day; }
        public static MonthsEnum GetMonthEnum(int month) { return (MonthsEnum)month; }
        public static int GetMonthInt(MonthsEnum month) { return (int)month; }
    }

    public class Event : SyncElement
    {
        public enum ImportanceEnum { Low, Normal, High }
        public enum StatusEnum { Free, Tentative, Busy, OutOfOffice, ElseWhere }
        public enum TypeEnum { SingleEvent, SeriesMaster, SeriesOccurence, SeriesException }

//        public string OriginId;
//        public DateTime LastModTime;

        public TypeEnum? EventType = null;
        public RecurrencePattern Recurrence = null;

        public string Subject;
        public string Location;
        public string Body;
        public StatusEnum? Status = null;
        public ImportanceEnum? Importance = null;
        public bool IsPrivate = false;
        public bool ReminderOn = false;
        public int ReminderMinutesBefore;

        public bool AllDayEvent;
        public DateTime? StartDateTime;
        public DateTime? StartUTC;
        public string StartTimeZone;
        public DateTime? EndDateTime;
        public DateTime? EndUTC;
        public string EndTimeZone;
        public int Duration { get => (this.EndDateTime - this.StartDateTime).GetValueOrDefault(new TimeSpan(0)).Minutes; }

        protected Event _parent = null;
        public Event Parent { get => _parent; }

        protected bool? _isRemoved = null;
        public bool? IsRemoved { get => _isRemoved; }

        public Events Exceptions { get; set; }
    }

    public class Events : SyncElementCollection<Event> { }
}
