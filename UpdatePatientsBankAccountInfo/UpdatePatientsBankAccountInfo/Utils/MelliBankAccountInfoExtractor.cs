using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Linq;
using System.Net.Http;
using System.Reflection.Metadata;
using System.Text;
using System.Text.RegularExpressions;
using Dapper;
using HtmlAgilityPack;

namespace UpdatePatientsBankAccountInfo.Utils
{
    public class MelliBankAccountInfoExtractor
    {
        public void Do(string dbFilePath, string tableFieldNames, IEnumerable<string> patientFileIds)
        {
            using var conn =
                new OdbcConnection($"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={dbFilePath};");
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

            var patients = conn.Query<object>($"SELECT [Melli_Acc_No], {tableFieldNames} FROM [MainTable] " +
                                              $"WHERE [File_No] IN ({string.Join(',', patientFileIds)});");
            
            using var file = System.IO.File.Create("report.csv");
            using var logWriter = new System.IO.StreamWriter(file, Encoding.UTF8);
            foreach (IDictionary<string, object> patient in patients)
            {
                var numbers = Regex.Split(patient.Values.First().ToString(), @"\D+").Where(x => x.Length >= 5).ToList();
                if (numbers.Count > 1)
                    Console.WriteLine("WARNING! more than one account number");

                var ownerName = "";
                fields[accountNoKey] = numbers[0];

                var res = httpClient.PostAsync(url, new FormUrlEncodedContent(fields)).Result;
                htmlDoc.LoadHtml(res.Content.ReadAsStringAsync().Result);
                var parts = htmlDoc.DocumentNode.SelectSingleNode(
                        "//span[@id='ctl00_MainContent_IBANSkinControl_ShabaOwnerLabel']")?
                    .InnerText.Split(' ').Where(x => !string.IsNullOrWhiteSpace(x));
                if (parts != null)
                    ownerName = string.Join(' ', parts);
                
                var csvRow = "";
                var index = 0;
                foreach (KeyValuePair<string, object> kv in patient)
                {
                    if (index++ == 0)
                        continue;
                    csvRow += kv.Value != null ? kv.Value.ToString() + "," : ",";
                }

                csvRow += ownerName;

                logWriter.WriteLine(csvRow);
            }
            logWriter.Close();
        }
    }
}