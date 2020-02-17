namespace GetBankAccountInfo.Controllers
{
    public class HomeModel
    {
        public string DbFilePath { get; set; }
        public string ExcelFilePath { get; set; }
        public int BankDelay { get; set; }
        public bool IsDone { get; set; }
    }
}