using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.IO;

namespace CopyO2O
{
    public static class SyncHelpers
    {
        private const char Marker_Splitter = ':';
        public enum ID_type : byte { CalItem, Contact };
        public enum SyncDirection : byte { NA, ToRight, ToLeft };
        public enum SyncMethod : byte { NA, CREATE, MODIFY, DELETE };

        public struct SyncInfo
        {
            public string OutlookID;
            public string O365ID;
            public ID_type Type;
            public DateTime LastSyncTime;
            public SyncDirection SyncDirection;
            public SyncMethod SyncWorkMethod;
        }

        public static List<SyncInfo> syncCache = new List<SyncInfo>();

        private static readonly string syncCache_FilePath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\sync.cache"; //logfile of all copy/sync sessions

        /// <summary>
        /// Store the IDs of the sync in the given file
        /// </summary>
        /// 
        public static void LoadSyncCache()
        {
            System.IO.StreamReader file = null;

            try
            {
                file = new System.IO.StreamReader(syncCache_FilePath);
                string line = null;

                while ((line = file.ReadLine()) != null)
                {
                    string[] parts = line.Split(Marker_Splitter);

                    SyncInfo tmpSyncInfo = new SyncInfo
                    {
                        Type = (ID_type)Enum.Parse(typeof(ID_type), parts[0]),
                        LastSyncTime = DateTime.FromFileTime(long.Parse(parts[1])),
                        OutlookID = parts[3],
                        O365ID = parts[5],
                        SyncDirection = (SyncDirection)Enum.Parse(typeof(SyncDirection), parts[6])
                    };

                    //add the current readen line to the sync cache in memory
                    lock (syncCache)
                    {
                        //only add if IDs are not yet registered
                        if (!syncCache.Exists(x => (x.OutlookID == tmpSyncInfo.OutlookID)) && (!syncCache.Exists(x => (x.O365ID == tmpSyncInfo.O365ID))))
                            syncCache.Add(tmpSyncInfo);
                    }
                }
            }
            catch (System.IO.FileNotFoundException e)
            {
                Program.LogLn(e.Message);
                return;
            }
            finally
            {
                if (file != null) file.Dispose();
            }
        }

        /// <summary>
        /// Store the IDs of the sync in the given file
        /// </summary>
        /// 
        public static void StoreSyncCache()
        {
            System.IO.StreamWriter file = new System.IO.StreamWriter(syncCache_FilePath);
            try
            {
                string line = null;
                foreach (SyncInfo info in syncCache)
                {
                    line =
                     info.Type.ToString() + Marker_Splitter
                     + info.LastSyncTime.ToFileTime() + Marker_Splitter
                     + "OUTLOOK" + Marker_Splitter
                     + info.OutlookID + Marker_Splitter
                     + "O365" + Marker_Splitter
                     + info.O365ID + Marker_Splitter
                     + info.SyncDirection.ToString() + Marker_Splitter 
                     + info.SyncWorkMethod.ToString();

                    file.WriteLine(line);
                }
            }
            finally
            {
                file.Flush();
                file.Dispose();
            }
        }

        public static void O365_Item_Added(string newId, SyncInfo Item)
        {
            lock (syncCache)
            {
                syncCache.Add(new SyncInfo
                {
                    OutlookID = Item.OutlookID,
                    O365ID = newId,
                    Type = Item.Type,
                    LastSyncTime = DateTime.Now,
                    SyncDirection = Item.SyncDirection,
                    SyncWorkMethod = Item.SyncWorkMethod
                });

                Program.LogLn("Item (" + Item.Type.ToString() + ") added: " + Item.OutlookID + " -> " + newId);
            }
        }

        public static void O365_Item_Removed(SyncInfo Item)
        {
            lock (syncCache)
            {
                syncCache.RemoveAll(x => (x.O365ID == Item.O365ID));

                Program.LogLn("Item (" + Item.Type.ToString() + ") removed: " + Item.O365ID);
            }
        }
        /// <summary>
        /// Creates items to O365
        /// </summary>
        /// <param name="syncWork">List of sync jobs to identify the items to create</param>
        /// <param name="o365_graphcontainer">The Office365 container of the items to create (e.g. calendar or contactfolder)</param>
        /// <param name="Type">Type of the items to be processed (e.g. events or contacts)</param>
        /// <returns>Returns count of jobs run.</returns>
        /// 
        public static int CreateItems(List<SyncHelpers.SyncInfo> syncWork, IEnumerable<SyncElement> srcItems, Office365.IGraphContainer o365_graphcontainer, SyncHelpers.ID_type Type)
        {
            List<Task> createTasks = new List<Task>();
            syncWork.Where(x => ((x.SyncWorkMethod == SyncHelpers.SyncMethod.CREATE) && (x.Type == Type))).ToList().ForEach(
                (item) =>
                {
                    createTasks.Add(o365_graphcontainer.AddAsync(srcItems.ToList().Find(x => (x.OriginId == item.OutlookID)))
                        .ContinueWith((i) => { if (i.Result != null) SyncHelpers.O365_Item_Added(i.Result, item); }));
                });
            Task.WaitAll(createTasks.ToArray());
            return createTasks.Count;
        }

        /// <summary>
        /// Removes items from O365
        /// </summary>
        /// <param name="syncWork">List of sync job to identify the items to remove.</param>
        /// <param name="o365_graphcontainer">The Office365 container of the items to delete (e.g. calendar or contactfolder)</param>
        /// <param name="Type">Type of the items to be processed (e.g. events or contacts)</param>
        /// <returns>Returns count of jobs run.</returns>
        /// 
        public static int RemoveItems(List<SyncHelpers.SyncInfo> syncWork, Office365.IGraphContainer o365_graphcontainer, SyncHelpers.ID_type Type)
        {
            List<Task> deleteTasks = new List<Task>();
            syncWork.Where(x => ((x.SyncWorkMethod == SyncHelpers.SyncMethod.DELETE) && (x.Type == Type))).ToList().ForEach(
                (item) =>
                {
                    deleteTasks.Add(o365_graphcontainer.RemoveAsync(item.O365ID)
                        .ContinueWith((i) => { if (i != null) SyncHelpers.O365_Item_Removed(item); }));
                });
            Task.WaitAll(deleteTasks.ToArray());
            return deleteTasks.Count;
        }

        /// <summary>
        /// Analyse and collect all sync jobs for the local Outlook instance.
        /// </summary>
        /// <param name="syncCache">The sync cache stored on disk</param>
        /// <param name="NewOutlookIds">All Ids which were added after the last sync job</param>
        /// <param name="ModifiedOutlookIds">All Ids which were modified after the last sync job</param>
        /// <param name="DeletedOutlookIds">All Ids which were removed after the last sync job</param>
        /// <param name="NewO365Ids">All Ids of Office 365 which were added after the last sync job</param>
        /// <param name="ModifiedO365Ids">All Ids of Office 365 which were modified after the last sync job</param>
        /// <param name="DeletedO365Ids">All Ids of Office 365 which were removed after the last sync job</param>
        /// <param name="Type">Type of the items to be processed (e.g. events or contacts)</param>
        /// <param name="clrNotExisting">If TRUE all items which were identified in O365 but not exist in Outlook will be removed as well</param>
        /// <returns>A list of sync jobs FROM Outlook TO Office365.</returns>
        /// 
        public static List<SyncHelpers.SyncInfo> CollectOutlookSyncTasks(
            List<SyncHelpers.SyncInfo> syncCache,
            List<string> NewOutlookIds, List<string> ModifiedOutlookIds, List<string> DeletedOutlookIds,
            List<string> NewO365Ids, List<string> ModifiedO365Ids, List<string> DeletedO365Ids,
            SyncHelpers.ID_type Type,
            bool clrNotExisting)
        {
            List<SyncHelpers.SyncInfo> result = new List<SyncHelpers.SyncInfo>();

            //1st OUTLOOK is the leading system, 2nd only sync to O365
            //add new outlook-elemente to sync list
            NewOutlookIds.ForEach((id) => result.Add(new SyncHelpers.SyncInfo { OutlookID = id, Type = Type, SyncDirection = SyncHelpers.SyncDirection.ToRight, SyncWorkMethod = SyncHelpers.SyncMethod.CREATE }));

            //add modified outlook-elemente to sync list
            ModifiedOutlookIds.ForEach((id) =>
            {
                result.Add(new SyncHelpers.SyncInfo { OutlookID = id, Type = Type, SyncDirection = SyncHelpers.SyncDirection.ToRight, SyncWorkMethod = SyncHelpers.SyncMethod.CREATE });
                result.Add(new SyncHelpers.SyncInfo { O365ID = syncCache.Find(x => (x.OutlookID == id)).O365ID, Type = Type, SyncDirection = SyncHelpers.SyncDirection.ToRight, SyncWorkMethod = SyncHelpers.SyncMethod.DELETE });
            });

            //remove deleted outlook elements from o365
            DeletedOutlookIds.ForEach((id) => result.Add(new SyncHelpers.SyncInfo { O365ID = syncCache.Find(x => (x.OutlookID == id)).O365ID, Type = Type, SyncDirection = SyncHelpers.SyncDirection.ToRight, SyncWorkMethod = SyncHelpers.SyncMethod.DELETE }));

            //restore from Office 365 deleted outlook-contacts again
            DeletedO365Ids.ForEach((id) =>
            {
                result.Add(new SyncHelpers.SyncInfo { OutlookID = syncCache.Find(x => (x.O365ID == id)).OutlookID, Type = Type, SyncDirection = SyncHelpers.SyncDirection.ToRight, SyncWorkMethod = SyncHelpers.SyncMethod.CREATE });
                SyncHelpers.syncCache.RemoveAll(x => (x.O365ID == id) && (x.Type == Type));
            });

            //if all items which does not exist in outlook should be removed from O365
            if (clrNotExisting)
            { NewO365Ids.ForEach((id) => result.Add(new SyncHelpers.SyncInfo { O365ID = id, Type = Type, SyncDirection = SyncHelpers.SyncDirection.ToRight, SyncWorkMethod = SyncHelpers.SyncMethod.DELETE })); }

            return result;
        }

        /// <summary>
        /// Identifiy all Ids of new items of a system (Outlook or Office 365)
        /// </summary>
        /// <param name="SrcItems">Items of the systems</param>
        /// <param name="SyncCache_Items">Sync cache (history) stored on disk</param>
        /// <param name="System">System to analyse (Outlook or Office 365)</param>
        /// <returns>Returns a list of Ids of new/added items since the last sync run.</returns>
        /// 
        public static List<string> GetIDsOfNewItems(IEnumerable<SyncElement> SrcItems, List<SyncHelpers.SyncInfo> SyncCache_Items, OriginSystemEnum System)
        {
            List<string> result = new List<string>();

            switch (System)
            {
                case OriginSystemEnum.Office365:
                    result.AddRange(SrcItems.Select(x => x.OriginId).Except(SyncCache_Items.Select(x => x.O365ID)));
                    break;
                case OriginSystemEnum.Outlook:
                    result.AddRange(SrcItems.Select(x => x.OriginId).Except(SyncCache_Items.Select(x => x.OutlookID)));
                    break;
            }
            return result;
        }

        /// <summary>
        /// Identifiy all Ids of modified items of a system (Outlook or Office 365)
        /// </summary>
        /// <param name="SrcItems">Items of the systems</param>
        /// <param name="SyncCache_Items">Sync cache (history) stored on disk</param>
        /// <param name="System">System to analyse (Outlook or Office 365)</param>
        /// <returns>Returns a list of Ids of modified items since the last sync run.</returns>
        /// 
        public static List<string> GetIDsOfModifiedItems(IEnumerable<SyncElement> SrcItems, List<SyncHelpers.SyncInfo> SyncCache_Items, OriginSystemEnum System)
        {
            List<string> tmpModifiedIds = new List<string>(); //all ids which were modified AFTER the last sync
            {
                SyncCache_Items.ForEach(
                    (item) =>
                    {
                        try
                        {
                            //if the current item could still be found (is not removed)
                            switch (System)
                            {
                                case OriginSystemEnum.Office365:
                                    if (SrcItems.ToList().Find(x => x.OriginId == item.O365ID).LastModTime > item.LastSyncTime.AddSeconds(10))
                                        tmpModifiedIds.Add(item.O365ID);
                                    break;
                                case OriginSystemEnum.Outlook:
                                    if (SrcItems.ToList().Find(x => x.OriginId == item.OutlookID).LastModTime > item.LastSyncTime.AddSeconds(10))
                                        tmpModifiedIds.Add(item.OutlookID);
                                    break;
                            }
                        }
                        catch (NullReferenceException e)
                        {
                            //not found -> nothing to do
                            Program.LogLn("Item not found: " + item.OutlookID + " > " + item.O365ID);
                        }
                    });
            }

            return tmpModifiedIds;
        }

        /// <summary>
        /// Identifiy all Ids of removed items of a system (Outlook or Office 365)
        /// </summary>
        /// <param name="SrcItems">Items of the systems</param>
        /// <param name="SyncCache_Items">Sync cache (history) stored on disk</param>
        /// <param name="System">System to analyse (Outlook or Office 365)</param>
        /// <returns>Returns a list of Ids of removed items since the last sync run.</returns>
        /// 
        public static List<string> GetIDsOfDeletedItems(IEnumerable<SyncElement> SrcItems, List<SyncHelpers.SyncInfo> SyncCache_Items, OriginSystemEnum System)
        {
            List<string> result = new List<string>();

            switch (System)
            {
                case OriginSystemEnum.Office365:
                    result.AddRange(SyncCache_Items.Select(x => x.O365ID).Except(SrcItems.Select(x => x.OriginId)));
                    break;
                case OriginSystemEnum.Outlook:
                    result.AddRange(SyncCache_Items.Select(x => x.OutlookID).Except(SrcItems.Select(x => x.OriginId)));
                    break;
            }
            return result;
        }
    }
}
