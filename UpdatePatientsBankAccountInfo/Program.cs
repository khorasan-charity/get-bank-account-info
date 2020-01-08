using System;
using System.IO;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Hosting;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using UpdatePatientsBankAccountInfo.Controllers;
using UpdatePatientsBankAccountInfo.Utils;


namespace UpdatePatientsBankAccountInfo
{
    public class Program
    {
        public static void Main(string[] args)
        {
            new MelliBankAccountInfoExtractor().Do(new HomeModel
            {
                BankDelay = 0,
                DbFilePath = ".\\mahak.accdb",
                ExcelFilePath = ".\\file.xlsx"
            });
            CreateHostBuilder(args).Build().Run();
        }

        public static IHostBuilder CreateHostBuilder(string[] args) =>
            Host.CreateDefaultBuilder(args)
                .ConfigureWebHostDefaults(webBuilder => { webBuilder.UseStartup<Startup>(); });
    }
}