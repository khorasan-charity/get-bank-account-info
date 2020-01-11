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
            var fileIds = ReadFileIdsFromExcel(model.ExcelFilePath);
            var patients = GetPatientInfo(model.DbFilePath, fileIds);
            var httpClient = new MelliBankHttpClient();

            var index = 0;
            foreach (var patient in patients)
            {
                index++;
                Console.WriteLine($"{index} of {patients.Count} ( {index * 100 / patients.Count}% )");
                
                if (string.IsNullOrWhiteSpace(patient.AccountNo))
                    continue;
                
                patient.RealAccountOwnerName = httpClient.GetAccountOwnerName(patient.AccountNo);
            }

            FillExcelAndSaveAs(model.ExcelFilePath, patients, ".\\result.xlsx");
            model.IsDone = true;
        }

        private void FillExcelAndSaveAs(string excelFilePath, IEnumerable<Patient> patients, string resultFilename)
        {
            var patientByFileId = patients.ToDictionary(x => x.FileNo);
            
            var excel = new XSSFWorkbook(excelFilePath);
            var sheet = excel.GetSheetAt(0);
            for (var i = 1;; i++)
            {
                var row = sheet.GetRow(i);
                var cell = row?.GetCell(3);
                var value = GetCellValueAsString(cell);
                if (value == null)
                    break;

                if (!patientByFileId.ContainsKey(value))
                    continue;
                
                row.CopyCell(2, 5).SetCellValue(patientByFileId[value].AccountOwnerName);
                row.CopyCell(2, 6).SetCellValue(patientByFileId[value].RealAccountOwnerName);
                row.CopyCell(2, 8).SetCellValue(patientByFileId[value].AccountNo);
            }
            
            SaveExcel(excel, resultFilename);
        }

        private static List<Patient> GetPatientInfo(string dbFilePath, List<string> fileIds)
        {
            using var conn =
                new OdbcConnection($"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={dbFilePath};");
            conn.Open();
            
            var patients = conn.Query<Patient>(
                "SELECT [File_No] AS FileNo, [Melli_Acc_No] AS AccountNo, [Melli_Acc_Owner_Name] AS AccountOwnerName FROM [MainTable] " +
                $"WHERE [File_No] IN ({string.Join(',', fileIds)});").ToList();
            
            conn.Close();

            foreach (var patient in patients)
            {
                if (string.IsNullOrWhiteSpace(patient.AccountNo))
                    continue;
                
                var accountNumbers = Regex.Split(patient.AccountNo, @"\D+").Where(x => x.Length >= 5).ToList();
                if (accountNumbers.Count > 1)
                    Console.WriteLine("WARNING! more than one account number");
                
                if (string.IsNullOrWhiteSpace(patient.AccountOwnerName))
                    patient.AccountOwnerName = patient.AccountNo.Replace(accountNumbers[0], "").Trim();
                patient.AccountNo = accountNumbers[0];
            }
            
            return patients;
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