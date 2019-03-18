using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using Commonfunctions.Cryptography;

namespace CopyO2O
{
    public enum OriginSystemEnum { Outlook, Office365 }

    public abstract class SyncElement
    {
        public OriginSystemEnum? OriginSystem;
        public string OriginId;
        public DateTime LastModTime = DateTime.Now;

        public string InternalId => MD5.GetMD5Hash(this.OriginSystem.GetType().ToString() + this.OriginId);

        public override bool Equals(object obj)
        {
            return (this.InternalId.Equals((obj as SyncElement).InternalId) && this.OriginId.Equals((obj as SyncElement).OriginId));
        }

        public override int GetHashCode()
        {
            return MD5.GetMD5Hash(this.InternalId + this.LastModTime.ToString()).GetHashCode();
        }
    }

    public abstract class SyncElementCollection<SyncElementType> : List<SyncElementType>
    {
        public SyncElementCollection() { }
        public SyncElementCollection(List<SyncElementType> items)
        {
            this.AddRange(items);
        }

        public SyncElementType this[string internalId]
        {
            get
            {
                return this.Find(x => (x as SyncElement).InternalId.Equals(internalId));
            }
        }

        public SyncElementType GetEventByOriginId(string Id, OriginSystemEnum origin)
        {
            return this.Find(x => ((x as SyncElement).OriginId.Equals(Id) && (x as SyncElement).OriginSystem.Equals(origin)));
        }

        public new SyncElementCollection<SyncElementType> FindAll(Predicate<SyncElementType> match)
        {
            return (SyncElementCollection<SyncElementType>)base.FindAll(match);
        }

        public bool Remove(string internalId)
        {
            return this.RemoveAll(x => (x as SyncElement).InternalId.Equals(internalId)) > 0;
        }
    }
}
