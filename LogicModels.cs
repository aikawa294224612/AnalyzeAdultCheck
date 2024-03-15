using Hl7.Fhir.Model;
using Hl7.Fhir.Rest;
using System.Globalization;

namespace TestAdultCheck

{
    class LogicModels
    { 
        public string ChangeGender(AdministrativeGender? gender)
        {
            string result = null;
            switch (gender)
            {
                case AdministrativeGender.Male:
                    result = "男";
                    break;
                case AdministrativeGender.Female:
                    result = "女";
                    break;
            }
            return result;
        }

        public string GetOrgName(string id)
        {
            string result = null;
            switch (id)
            {
                case "1539335":
                    result = "海端衛生所";
                    break;
                case "3983763":
                    result = "延平衛生所";
                    break;
            }
            return result;
        }

        public string AdToRocEra(string ad)
        {
            DateTime dt = DateTime.Parse(ad);
            CultureInfo culture = new CultureInfo("zh-TW");
            culture.DateTimeFormat.Calendar = new TaiwanCalendar();
            return dt.ToString("yyyMMdd", culture);
        }
    }
}