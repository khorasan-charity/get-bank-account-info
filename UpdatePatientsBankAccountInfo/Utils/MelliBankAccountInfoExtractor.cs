using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection.Metadata;
using System.Text;
using System.Text.RegularExpressions;
using Dapper;
using HtmlAgilityPack;
using NPOI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using UpdatePatientsBankAccountInfo.Controllers;

namespace UpdatePatientsBankAccountInfo.Utils
{
    public class MelliBankAccountInfoExtractor
    {
        public void Do(HomeModel model)
        {
            using var conn =
                new OdbcConnection($"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={model.DbFilePath};");
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

            var fileIds = ReadFileIdsFromExcel(model.ExcelFilePath);
            var patients = conn.Query<Patient>(
                "SELECT [File_No] AS FileNo, [Melli_Acc_No] AS AccountNo, [Melli_Acc_Owner_Name] AS AccountOwnerName FROM [MainTable] " +
                $"WHERE [File_No] IN ({string.Join(',', fileIds)});").ToList();


            var index = 1;
            foreach (var patient in patients)
            {
                var accountNumbers = Regex.Split(patient.AccountNo, @"\D+").Where(x => x.Length >= 5).ToList();
                if (accountNumbers.Count > 1)
                    Console.WriteLine("WARNING! more than one account number");
                
                fields[accountNoKey] = accountNumbers[0];

                var res = httpClient.PostAsync(url, new FormUrlEncodedContent(fields)).Result;
                htmlDoc.LoadHtml(res.Content.ReadAsStringAsync().Result);
                var parts = htmlDoc.DocumentNode.SelectSingleNode(
                        "//span[@id='ctl00_MainContent_IBANSkinControl_ShabaOwnerLabel']")?
                    .InnerText.Split(' ').Where(x => !string.IsNullOrWhiteSpace(x));
                if (parts != null)
                    patient.RealAccountOwnerName = string.Join(' ', parts);
                
                Console.WriteLine($"{index} of {patients.Count} ( {index * 100 / patients.Count}% )");
            }
            
            var patientByFileId = patients.ToDictionary(x => x.FileNo);
            
            var excel = new XSSFWorkbook(model.ExcelFilePath);
            var sheet = excel.GetSheetAt(0);
            for (var i = 1;; i++)
            {
                var row = sheet.GetRow(i);
                var cell = row?.GetCell(3);
                var value = GetCellValueAsString(cell);
                if (value == null)
                    break;
                
                if (patientByFileId.ContainsKey(value))
                row.CopyCell(2, 5).SetCellValue("سلاممممم");
                row.CopyCell(2, 6).SetCellValue("خداحااافظ");
            }
            
            SaveExcel(excel, ".\\file2.xlsx");
        }

        private List<string> ReadFileIdsFromExcel(string filename)
        {
            var fileIds = new List<string>();
            var excel = new XSSFWorkbook(filename);
            var sheet = excel.GetSheetAt(0);
            for (var i = 1;; i++)
            {
                var row = sheet.GetRow(i);
                var cell = row?.GetCell(3);
                var value = GetCellValueAsString(cell);
                if (value == null)
                    break;
                fileIds.Add(value.Trim());
            }
            excel.Close();
            return fileIds;
        }

        private string GetCellValueAsString(ICell cell)
        {
            if (IsCellEmpty(cell))
                return null;
            
            return cell.CellType switch
            {
                CellType.Numeric => cell.NumericCellValue.ToString(CultureInfo.InvariantCulture),
                CellType.String => cell.StringCellValue,
                _ => null
            };
        }

        private bool IsCellEmpty(ICell cell)
        {
            return cell == null
                   || cell.CellType == CellType.Blank
                   || (cell.CellType == CellType.String && string.IsNullOrWhiteSpace(cell.StringCellValue))
                   || (cell.CellType == CellType.Numeric && Math.Abs(cell.NumericCellValue) < 1);
        }

        private static void SaveExcel(XSSFWorkbook excel, string filename)
        {
            using var file = File.OpenWrite(filename);
            excel.Write(file);
        }
    }
}