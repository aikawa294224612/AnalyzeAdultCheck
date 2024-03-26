using Hl7.Fhir.Model;
using Hl7.Fhir.Rest;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Web;

namespace TestAdultCheck

{
    class LogicModels
    {
        static DataService secret = new DataService();

        public string GetOrgId(FhirClient client, string identifier)
        {
            var searchParams = new SearchParams();
            searchParams.Add("identifier", identifier);

            Bundle orgs = client.Search<Organization>(searchParams);

            Organization org = (Organization)orgs.Entry[0].Resource;

            return org.Id;

        }

        public string GetIdentifierValue(Patient pat, string typeCode, bool decryp)
        {
            string value = null;

            foreach (Identifier identifier in pat.Identifier)
            {
                if (identifier.Type.Coding[0].Code == typeCode)
                {
                    value = decryp ? DecryptStringFromBytes_Aes(identifier.Value) : identifier.Value;
                    break;
                }
            }

            return value;
        }

        public string GetIdentifierValue(Organization org, string typeCode)
        {
            string value = null;

            foreach (Identifier identifier in org.Identifier)
            {
                if (identifier.Type.Coding[0].Code == typeCode)
                {
                    value = identifier.Value;
                    break;
                }
            }

            return value;
        }

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

        public string GetOrgName(string iden)
        {
            string result = null;
            switch (iden)
            {
                case "2346130016":
                    result = "海端衛生所";
                    break;
                case "2346100018":
                    result = "池上鄉衛生所";
                    break;
                case "2346010019":
                    result = "台東巿衛生所";
                    break;
                case "2346120010":
                    result = "延平鄉衛生所";
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

        public Bundle? GetNextPages(Bundle results, FhirClient client, string getpagesoffset)
        {   
            string nextLink = null;
            foreach(var link in results.Link)
            {
                if(link.Relation == "next")
                {
                    if (!link.Url.Contains("https"))
                    {
                        nextLink = link.Url.Replace("http", "https");
                    }
                    else
                    {
                        nextLink = link.Url;
                    }
                    
                }
            }
            if(nextLink == null) { 
                return null;
            }
            Uri myUri = new Uri(nextLink);
            string getpages = HttpUtility.ParseQueryString(myUri.Query).Get("_getpages");


            if(getpagesoffset == null)
            {
                getpagesoffset = HttpUtility.ParseQueryString(myUri.Query).Get("_getpagesoffset");
            }

            var searchParams = new SearchParams();
            searchParams.Add("_getpages", getpages);
            searchParams.Add("_getpagesoffset", getpagesoffset);  //Important
            searchParams.Add("_count", "50");
            searchParams.Add("_pretty", "true");
            searchParams.Add("_bundletype", "searchset");

            Console.WriteLine("next page: " + nextLink);

            return client.Search(searchParams);
        }

        public string DecryptStringFromBytes_Aes(string cipherText)
        {
            string key = secret.secret;

            byte[] keyBytes = Encoding.UTF8.GetBytes(key.Substring(0, 16));
            byte[] cipherBytes = Convert.FromBase64String(cipherText);

            using (Aes aes = Aes.Create())
            {
                aes.Mode = CipherMode.ECB;
                aes.KeySize = 128;
                aes.Key = keyBytes;
                aes.Padding = PaddingMode.PKCS7;

                using (MemoryStream ms = new MemoryStream())
                {
                    using (CryptoStream cs = new CryptoStream(ms, aes.CreateDecryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(cipherBytes, 0, cipherBytes.Length);
                        cs.Close();
                    }
                    byte[] decryptedBytes = ms.ToArray();
                    return Encoding.UTF8.GetString(decryptedBytes);
                }
            }
        }
    }
}