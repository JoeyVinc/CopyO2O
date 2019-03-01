using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

namespace CopyO2O
{
    public static class Helpers
    {
        private const char Marker_Splitter = ':';
        public enum ID_type : byte { CalItem=1, Contact=2 };

        public const string Marker_O365 = "O365";
        public const string Marker_Outlook = "OUTLOOK";


        public static List<KeyValuePair<string, string>> GetIDsToRemove(ID_type ItemType, string FilePath)
        {
            List<KeyValuePair<string, string>> result = new List<KeyValuePair<string, string>>();

            System.IO.StreamReader file = new System.IO.StreamReader(FilePath);

            //if file could not be found
            if (file == null) return null;

            try
            {
                string line;
                while ((line = file.ReadLine()) != null)
                {
                    Debug.WriteLine("Read line: " + line);

                    String[] splittedLine = line.Split(Marker_Splitter);

                    string target = splittedLine[0];
                    ID_type type = (ID_type)int.Parse(splittedLine[1]);
                    string id = splittedLine[2];

                    switch (target.ToUpper())
                    {
                        case Marker_O365:
                            result.Add(new KeyValuePair<string, string>(Marker_O365, id));
                            break;
                        case Marker_Outlook:
                            result.Add(new KeyValuePair<string, string>(Marker_Outlook, id));
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

        public static bool SetIDsToRemove(List<KeyValuePair<string, string>> Ids, ID_type ItemType, string FilePath)
        {
            System.IO.StreamWriter file = new System.IO.StreamWriter(FilePath);

            //if file could not be opened
            if (file == null) return false;

            try
            {
                file.AutoFlush = true;
                string line = "";

                foreach (KeyValuePair<string, string> id in Ids)
                {
                    switch (id.Key.ToUpper())
                    {
                        case Marker_O365:
                            line = Marker_O365;
                            break;
                        case Marker_Outlook:
                            line = Marker_Outlook;
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
    }
}
