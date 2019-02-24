using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CopyO2O
{
    public struct Address
    {
        public string Street;
        public string Number;
        public string City;
        public string Zip;
        public string Country;
    }

    public struct Category
    {
        public string Name;
        public System.Drawing.Color Color;
    }

    public struct EMail
    {
        public string Address;
        public string Title;
    }

    public class ContactType
    {
        public Boolean HasPhoto { get { return this.PictureTmpFilename != null; } }

        public string SaveAs;
        public string DisplayName;
        public string Title;
        public string GivenName;
        public string MiddleName;
        public string Surname;
        public string AddName;
        public string Company;
        public DateTime? Birthday;
        public DateTime? AnniversaryDay;
        public Boolean VIP;
        public List<Category> Categories;
        public string Notes;

        public Address PrivateLocation;
        public EMail PrivateMailAddress;
        public string PrivatePhoneNumber;
        public string PrivateMobileNumber;
        public string PrivateFaxNumber;

        public Address BusinessLocation;
        public EMail BusinessMailAddress;
        public string BusinessPhoneNumber;
        public string BusinessMobileNumber;
        public string BusinessFaxNumber;

        public string PictureTmpFilename;
    }

    public class ContactCollectionType : List<ContactType> { }
}
