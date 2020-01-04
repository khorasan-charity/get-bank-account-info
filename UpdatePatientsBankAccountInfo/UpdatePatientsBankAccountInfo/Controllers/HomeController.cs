using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Extensions.Logging;
using UpdatePatientsBankAccountInfo.Models;
using UpdatePatientsBankAccountInfo.Utils;

namespace UpdatePatientsBankAccountInfo.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View(new HomeModel
            {
                DbFilePath = ".\\محک.accdb",
                FieldNames = "[File_No], [Name_First] + ' ' + [Name_Last], [Melli_Acc_No], [Melli_Acc_Owner_Name]",
                FileIds = ""
            });
        }

        [HttpPost]
        public IActionResult Report(HomeModel model)
        {
            var extractor = new MelliBankAccountInfoExtractor();
            extractor.Do(model.DbFilePath, model.FieldNames, model.FileIds.Split(Environment.NewLine));
            return View("Index", model);
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel {RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier});
        }
    }
}