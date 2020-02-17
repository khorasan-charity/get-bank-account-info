using System;
using System.Collections.Generic;
using System.Diagnostics;
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
                Percent = _bankAccountInfoExtractor.ProgressValue * 100 / (double)_bankAccountInfoExtractor.ProgressTotal 
            };
        }

        [HttpPost("upload-1")]
        public IActionResult Upload1(IFormFile file)
        {
            _bankAccountInfoExtractor.Do1(file.FileName, "result.xlsx");
            return Ok();
        }
    }
}