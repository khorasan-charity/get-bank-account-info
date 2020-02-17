using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection.Metadata;
using System.Text;
using System.Text.RegularExpressions;
using Dapper;
using GetBankAccountInfo.Controllers;
using HtmlAgilityPack;
using NPOI;
using NPOI.HSSF.Util;
using NPOI.OpenXmlFormats.Dml.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace GetBankAccountInfo.Utils
{
    public class MelliBankAccountInfoExtractor
    {
        public int ProgressTotal { get; set; }
        public int ProgressValue { get; set; }
        public bool IsBusy { get; set; }

        public void Do1(string inputExcel, string outputExcel)
        {
            IsBusy = true;
            
            var patients = ReadPatientsFromExcel(inputExcel);
            ProgressTotal = patients.Count;
            ProgressValue = 0;
            var httpClient = new MelliBankHttpClient();
            
            foreach (var patient in patients)
            {
                ProgressValue++;
                Console.WriteLine($"{ProgressValue} of {ProgressTotal} ( {ProgressValue * 100 / ProgressTotal}% )");
                
                if (string.IsNullOrWhiteSpace(patient.AccountNo))
                    continue;
                
                patient.RealAccountOwnerName = httpClient.GetAccountOwnerName(patient.AccountNo);
            }

            SavePatientsAsExcel(inputExcel, 3, outputExcel, patients);
            
            IsBusy = false;
        }
        
        public void Do2(HomeModel model)
        {
            IsBusy = true;
            
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
            
            IsBusy = false;
        }

        private void SavePatientsAsExcel(string inputFilename, int resultColumnNo,
            string outputFilename, List<Patient> patients)
        {
            var patientByFileId = patients.ToDictionary(x => x.FileNo);
            
            var excel = new XSSFWorkbook(inputFilename);
            var sheet = excel.GetSheetAt(0);
            SetCellValue(sheet, 0, resultColumnNo, "نام استعلام شده از بانک");
            
            for (var i = 1;; i++)
            {
                var row = sheet.GetRow(i);
                var fileId = GetCellValueAsString(row?.GetCell(0));
                if (fileId == null)
                    break;

                if (!patientByFileId.ContainsKey(fileId))
                    continue;
                
                if(GetCellValueAsString(row.GetCell(4)) != patientByFileId[fileId].AccountNo)
                    continue;
                
                SetCellValue(sheet, i, resultColumnNo, patientByFileId[fileId].RealAccountOwnerName);
                // if (patientByFileId[fileId].RealAccountOwnerName == patientByFileId[fileId].AccountOwnerName)
                // {
                //     //setCellStyle(excel, destCell);
                //     //destCell.CellStyle = matchStyle;
                //     //destCell.CellStyle.FillPattern = FillPattern.SolidForeground;
                //     //destCell.CellStyle.FillForegroundColor = new XSSFColor(Color.LightGreen).Indexed;
                // }
            }
            
            SaveAndCloseExcel(excel, outputFilename);
        }

        private static void SetCellValue(ISheet sheet, int rowNo, int columnNo, string value)
        {
            var row = sheet.GetRow(rowNo) ?? sheet.CreateRow(rowNo); 
            var cell = row.GetCell(columnNo) ?? row.CreateCell(columnNo);
            cell.SetCellType(CellType.String);
            cell.SetCellValue(value);
        }

        private void setCellStyle(XSSFWorkbook workbook, ICell cell)
        {
            var fCellStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            
            fCellStyle.FillForegroundXSSFColor = new XSSFColor(Color.Green);
            fCellStyle.FillPattern = FillPattern.SolidForeground;

            fCellStyle.VerticalAlignment = VerticalAlignment.Center;
            fCellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
          
            cell.CellStyle = fCellStyle;
        }

        private List<Patient> ReadPatientsFromExcel(string filename)
        {
            var patients = new List<Patient>();
            var excel = new XSSFWorkbook(filename);
            var sheet = excel.GetSheetAt(0);
            for (var i = 1;; i++)
            {
                var row = sheet.GetRow(i);
                var fileId = GetCellValueAsString(row?.GetCell(0));
                if (fileId == null)
                    break;
                patients.Add(new Patient
                {
                    FileNo = fileId,
                    AccountNo = GetCellValueAsString(row?.GetCell(4)),
                    AccountOwnerName = GetCellValueAsString(row?.GetCell(2)),
                    RealAccountOwnerName = ""
                });
            }
            excel.Close();
            return patients;
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
            
            SaveAndCloseExcel(excel, resultFilename);
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
                CellType.String => cell.StringCellValue.Trim(),
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

        private static void SaveAndCloseExcel(XSSFWorkbook excel, string filename)
        {
            using var file = File.OpenWrite(filename);
            excel.Write(file);
            excel.Close();
        }
    }
}