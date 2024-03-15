using Hl7.Fhir.Model;
using Hl7.Fhir.Rest;
using System.Globalization;
using System.Web;

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
            if(ad == "")
            {
                return "";
            }
            else
            {
                DateTime dt = DateTime.Parse(ad);
                CultureInfo culture = new CultureInfo("zh-TW");
                culture.DateTimeFormat.Calendar = new TaiwanCalendar();
                return dt.ToString("yyyMMdd", culture);
            }
            
        }

        public Bundle? GetNextPages(Bundle results, FhirClient client)
        {   
            string nextLink = null;
            foreach(var link in results.Link)
            {
                if(link.Relation == "next")
                {
                    nextLink = link.Url.Replace("http", "https");
                }
            }
            if(nextLink == null) { 
                return null;
            }
            Uri myUri = new Uri(nextLink);
            string getpages = HttpUtility.ParseQueryString(myUri.Query).Get("_getpages");
            string getpagesoffset = HttpUtility.ParseQueryString(myUri.Query).Get("_getpagesoffset");

            var searchParams = new SearchParams();
            searchParams.Add("_getpages", getpages);
            searchParams.Add("_getpagesoffset", getpagesoffset);  //Important
            searchParams.Add("_count", "50");
            searchParams.Add("_pretty", "true");
            searchParams.Add("_bundletype", "searchset");

            Console.WriteLine("next page: " + nextLink);

            return client.Search<Composition>(searchParams);
        }
    }
}