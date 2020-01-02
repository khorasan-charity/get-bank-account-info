using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Mime;
using System.Text.RegularExpressions;
using Dapper;
using HtmlAgilityPack;
using RestSharp;

namespace UpdateMahakPatientsBankAccountInfo
{
    class Program
    {
        static void Main(string[] args)
        {
            using var conn =
                new OdbcConnection(
                    "Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\\Users\\Majid\\Desktop\\mahak.accdb;");
            conn.Open();
            var handler = new HttpClientHandler() {UseCookies = true};
            using var httpClient = new HttpClient(handler);
            httpClient.DefaultRequestHeaders.Add("User-Agent",
                "Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:x.x.x) Gecko/20041107 Firefox/x.x");
            const string url = "https://bmi.ir/fa/shaba/";
            var htmlDoc = new HtmlDocument();

            var resp = httpClient.GetAsync(url).Result;
            htmlDoc.LoadHtml(resp.Content.ReadAsStringAsync().Result);
            var cookies = resp.Headers.GetValues("Set-Cookie").ToList();
            const string accountNoKey = "ctl00$MainContent$IBANSkinControl$AccountNumberTextBox";
            var fields = new Dictionary<string, string>
            {
                {"__VIEWSTATE", ""},
                {"__VIEWSTATEGENERATOR", ""},
                {"__EVENTVALIDATION", ""}
            };
            foreach (var key in fields.Keys.ToList())
            {
                fields[key] = htmlDoc.DocumentNode.SelectSingleNode($"//input[@name='{key}']").Attributes["value"]
                    .Value;
            }

            fields.Add("__EVENTTARGET", "ctl00$MainContent$IBANSkinControl$GetShabaLinkButton");
            fields.Add(accountNoKey, "");
            fields.Add("ctl00$MainContent$IBANSkinControl$AccountTypeRadioButtonList", "0");

            var accounts = conn.Query<string>("SELECT TOP 20 [Melli_Acc_No] FROM [MainTable] WHERE Melli_Acc_No<>'';");
            
            foreach (var account in accounts)
            {
                var numbers = Regex.Split(account, @"\D+").Where(x => x.Length >= 5).ToList();
                if (numbers.Count > 1)
                    Console.WriteLine("WARNING! more than one account number");
                foreach (var value in numbers)
                {
                    fields[accountNoKey] = value;
                    
                    var res = httpClient.PostAsync(url, new FormUrlEncodedContent(fields)).Result;
                    htmlDoc.LoadHtml(res.Content.ReadAsStringAsync().Result);
                    var name = "";
                    var parts = htmlDoc.DocumentNode.SelectSingleNode(
                            "//span[@id='ctl00_MainContent_IBANSkinControl_ShabaOwnerLabel']")?
                        .InnerText.Split(' ').Where(x => !string.IsNullOrWhiteSpace(x));
                    if (parts != null)
                        name = string.Join(' ', parts);
                    Console.WriteLine(name);
                }
            }
            
        }
    }
}