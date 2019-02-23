# CopyO2O
Duplicates items (at the moment calendar items and contacts) from Outlook to Office 365.

Remark: no synchronisation! Items of destination (Office 365) will be deleted without warning before the new values will be copied.

## Current call parameters
```
/CAL:"<source>";"<destination>" : Calendar source and destination
/CON:"<source>";"<destination>" : Contacts source and destination  
[opt] /from:<date>              : for calendar: First date to sync (DD.MM.YYYY) or relative to today (in days; eg. -10)  
[opt] /to:<date>                : for calendar: Last date to sync (DD.MM.YYYY) or relative to today (in days; eg. 8)  
[opt] /clear:<days>             : for calendar: Clear <days> in the past (from 'from' back)  
[opt] /log                      : Verbose logging
```

Example: `CopyO2O /CAL:"Hans.Mustermann@company.com\Calendar";"Business" /from:-7 /to:30 /clear:14`

=> copy calendar items from `Calendar`-folder of the local Outlook postbox `Hans.Mustermann@company.com` to the calendar folder `Business` in Office 365. Copy items from seven days in the past to 30 days in future. During the process clear all calendar items which have a startdate from 21 days to 7 days in the past as well.
Therefore if today is the 15.02. all events starting from the 24th of Jan to today and copy the elements from the 8th of Feb to the 15th of March.
