using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using HtmlAgilityPack;

namespace UpdatePatientsBankAccountInfo.Utils
{
    public class MelliBankHttpClient
    {
        private readonly HttpClient _httpClient;
        private readonly Dictionary<string, string> _fields;
        private const string AccountNoKey = "ctl00$MainContent$IBANSkinControl$AccountNumberTextBox";
        private const string Url = "https://bmi.ir/fa/shaba/";
        
        public MelliBankHttpClient()
        {
            _httpClient = new HttpClient(new HttpClientHandler {UseCookies = true});
            _httpClient.DefaultRequestHeaders.Add("User-Agent",
                "Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:x.x.x) Gecko/20041107 Firefox/x.x");
            
            var htmlDoc = new HtmlDocument();

            var resp = _httpClient.GetAsync(Url).Result;
            htmlDoc.LoadHtml(resp.Content.ReadAsStringAsync().Result);
            
            _fields = new Dictionary<string, string>
            {
                {"__VIEWSTATE", ""},
                {"__VIEWSTATEGENERATOR", ""},
                {"__EVENTVALIDATION", ""}
            };
            foreach (var key in _fields.Keys.ToList())
            {
                _fields[key] = htmlDoc.DocumentNode.SelectSingleNode($"//input[@name='{key}']").Attributes["value"]
                    .Value;
            }

            _fields.Add("__EVENTTARGET", "ctl00$MainContent$IBANSkinControl$GetShabaLinkButton");
            _fields.Add(AccountNoKey, "");
            _fields.Add("ctl00$MainContent$IBANSkinControl$AccountTypeRadioButtonList", "0");
        }

        public string GetAccountOwnerName(string accountNo)
        {
            _fields[AccountNoKey] = accountNo;
            
            var htmlDoc = new HtmlDocument();
            var res = _httpClient.PostAsync(Url, new FormUrlEncodedContent(_fields)).Result;
            htmlDoc.LoadHtml(res.Content.ReadAsStringAsync().Result);
            var parts = htmlDoc.DocumentNode.SelectSingleNode(
                    "//span[@id='ctl00_MainContent_IBANSkinControl_ShabaOwnerLabel']")?
                .InnerText.Split(' ').Where(x => !string.IsNullOrWhiteSpace(x));
            
            return parts != null ? string.Join(' ', parts) : null;
        }
    }
}