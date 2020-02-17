using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using GetBankAccountInfo.Models;
using GetBankAccountInfo.Utils;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Extensions.Logging;

namespace GetBankAccountInfo.Controllers
{
    [Route("/api/[controller]")]
    [ApiController]
    public class BankController : ControllerBase
    {
        private readonly ILogger<BankController> _logger;
        private readonly MelliBankAccountInfoExtractor _bankAccountInfoExtractor;

        public BankController(ILogger<BankController> logger, MelliBankAccountInfoExtractor bankAccountInfoExtractor)
        {
            _logger = logger;
            _bankAccountInfoExtractor = bankAccountInfoExtractor;
        }

        [HttpGet("progress")]
        public ActionResult<dynamic> GetProgress()
        {
            return new
            {
                Value = _bankAccountInfoExtractor.ProgressValue,
                Total = _bankAccountInfoExtractor.ProgressTotal,
                Percent = _bankAccountInfoExtractor.ProgressValue * 100 / Math.Max((double)_bankAccountInfoExtractor.ProgressTotal, 1) 
            };
        }

        [HttpPost("upload1")]
        public dynamic Upload1(IFormFile file)
        {
            try
            {
                const string inputFileName = "wwwroot/tmp/input.xlsx";
                Directory.CreateDirectory("wwwroot/tmp");
                System.IO.File.Delete(inputFileName);
                using var f = System.IO.File.OpenWrite(inputFileName);
                file.OpenReadStream().CopyTo(f);
                f.Close();
                _bankAccountInfoExtractor.Do1(inputFileName, "wwwroot/tmp/result.xlsx");
                return new
                {
                    Url = $"{this.Request.Scheme}://{this.Request.Host}{this.Request.PathBase}/tmp/result.xlsx"
                };
            }
            catch (Exception e)
            {
                return new
                {
                    Error = e.Message
                };
            }
        }
    }
}