using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoCompany.Model
{
    public class Companny
    {
        public Companny()
        {
        }
        public Companny(string _MST)
        {
            this.MST = _MST;
        }
        public Companny(string _MST, string typetpl, string stt)
        {
            this.MST = _MST;
            TypeTEMPLATE = typetpl;
            StatusGET = stt;
        }

        public Companny(Companny companny)
        {
            Name = companny.Name;
            NameLanguageOrther = companny.NameLanguageOrther;
            NameShortCut = companny.NameShortCut;
            RepresentativeName = companny.RepresentativeName;
            MST = companny.MST;
            SDT = companny.SDT;
            Type = companny.Type;
            OperationDate = companny.OperationDate;
            LicenseDate = companny.LicenseDate;
            Address = companny.Address;
            Status = companny.Status;
            StatusGET = companny.StatusGET;
            TypeTEMPLATE = companny.TypeTEMPLATE;
        }
        public string NameTo2Word()
        {
            var words = Name.Split();
            var lastTwoWords = string.Join(" ", words.Skip(words.Length - 2)).Trim().ToUpper();
            return lastTwoWords;
        }
        public string AddressToCity()
        {
            var words = Address.Split(char.Parse(","));
            var City = words[words.Count() - 2];
            return City.Trim();
        }
        public string OwnAuthor = "";
        public string Name = "";
        public string NameLanguageOrther = "";
        public string NameShortCut = "";
        public string RepresentativeName = "";//Đại diện
        public string MST = "";
        public string SDT = "";
        public string Type = ""; //Một thành viên
        public string OperationDate = "";//Ngày vận hành
        public string LicenseDate = "";//ngày cấp phép
        public string Address = "";
        public string Status = "";
        public string StatusGET = "";
        public string TypeTEMPLATE = "";

    }
}
