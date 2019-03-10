using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.IO;

namespace CopyO2O
{
    public static class Helpers
    {
        private const char Marker_Splitter = ':';
        public enum ID_type : byte { CalItem, Contact };
        public enum ID_system : byte { O365, OUTLOOK };
        public enum SyncDirection : byte { NA, ToRight, ToLeft };
        public enum SyncMethod : byte { NA, CREATE, MODIFY, DELETE };

        private static readonly string tmpIdsToDelete_FilePath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\tmpSync.cache"; //temporary file to store IDs to delete

        /// <summary>
        /// Read all IDs stored in a file which should be deleted
        /// </summary>
        /// <param name="ItemType">Type of IDs</param>
        /// <param name="FilePath">Path and filename of the file to read</param>
        /// <returns>List of IDs stored in FilePath of type ItemType</returns>
        /// 
        public static List<KeyValuePair<ID_system, string>> GetIDsToRemove(ID_type ItemType)
        {
            List<KeyValuePair<ID_system, string>> result = new List<KeyValuePair<ID_system, string>>();

            System.IO.StreamReader file = new System.IO.StreamReader(tmpIdsToDelete_FilePath);

            //if file could not be found
            if (file == null) return null;

            try
            {
                string line;
                while ((line = file.ReadLine()) != null)
                {
                    Debug.WriteLine("Read line: " + line);

                    String[] splittedLine = line.Split(Marker_Splitter);

                    string system = splittedLine[0];
                    ID_type type = (ID_type)int.Parse(splittedLine[1]);
                    string id = splittedLine[2];

                    switch ((ID_system)int.Parse(system))
                    {
                        case ID_system.O365:
                            result.Add(new KeyValuePair<ID_system, string>(ID_system.O365, id));
                            break;
                        case ID_system.OUTLOOK:
                            result.Add(new KeyValuePair<ID_system, string>(ID_system.OUTLOOK, id));
                            break;
                        default:
                            break;
                    }
                }
            }
            finally
            {
                file.Dispose();
            }
            return result;
        }

        /// <summary>
        /// Stores IDs in a file.
        /// </summary>
        /// <param name="Ids">A list of IDs to store. Key-value pair: 1st system of ID, 2nd ID</param>
        /// <param name="ItemType">Type of ID</param>
        /// <param name="FilePath">Path and filename of the target file</param>
        /// <returns>Return TRUE of successfill; otherwise FALSE</returns>
        /// 
        public static bool SetIDsToRemove(List<KeyValuePair<ID_system, string>> Ids, ID_type ItemType)
        {
            System.IO.StreamWriter file = new System.IO.StreamWriter(tmpIdsToDelete_FilePath);

            //if file could not be opened
            if (file == null) return false;

            try
            {
                file.AutoFlush = true;
                string line = "";

                foreach (KeyValuePair<ID_system, string> id in Ids)
                {
                    switch (id.Key)
                    {
                        case ID_system.O365:
                            line = ((int)ID_system.O365).ToString();
                            break;
                        case ID_system.OUTLOOK:
                            line = ((int)ID_system.OUTLOOK).ToString();
                            break;
                        default: break;
                    }

                    line += Marker_Splitter + ((int)ItemType).ToString() + Marker_Splitter + id.Value;
                    Debug.WriteLine("Write line: " + line);
                    file.WriteLine(line);
                }
            }
            finally
            {
                file.Flush();
                file.Dispose();
            }
            return true;
        }

        public static void DeleteRemovalLog()
        {
            System.IO.File.Delete(tmpIdsToDelete_FilePath);
        }

        /// <summary>
        /// <param name="Source">Source of the item</param>
        /// <param name="SourceID">Original ID</param>
        /// <param name="Target">Destination for the item</param>
        /// <param name="TargetID">ID in the destination system</param>
        /// <param name="Type">Type of the item</param>
        /// </summary>
        /// 
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
                     + ID_system.OUTLOOK.ToString() + Marker_Splitter
                     + info.OutlookID + Marker_Splitter
                     + ID_system.O365.ToString() + Marker_Splitter
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

        public static void O365_Contact_Added(string newO365ID, string OutlookID)
        {
            lock (syncCache)
            {
                syncCache.Add(new SyncInfo
                {
                    OutlookID = OutlookID,
                    O365ID = newO365ID,
                    Type = ID_type.Contact,
                    LastSyncTime = DateTime.Now,
                    SyncDirection = SyncDirection.ToRight,
                    SyncWorkMethod = SyncMethod.CREATE
                });
            }
        }

        public static void O365_Event_Added(string newO365ID, string OutlookID)
        {
            lock (syncCache)
            {
                syncCache.Add(new SyncInfo
                {
                    OutlookID = OutlookID,
                    O365ID = newO365ID,
                    Type = ID_type.CalItem,
                    LastSyncTime = DateTime.Now,
                    SyncDirection = SyncDirection.ToRight,
                    SyncWorkMethod = SyncMethod.CREATE
                });
            }
        }

        public static void O365_Item_Removed(string O365ID)
        {
            lock (syncCache)
            {
                syncCache.RemoveAll(x => (x.O365ID == O365ID));
            }
        }
    }
}
