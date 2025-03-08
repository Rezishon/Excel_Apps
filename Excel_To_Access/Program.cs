#undef Testing
using OfficeOpenXml;

namespace Excel_To_Access;

class Program
{
    static void Main(string[] args)
    {
        #region Setting

        string mainFilePath = @"/home/rezishon/storage/Project/Excel_Apps/Excel_To_Access/Data/14031217.xlsx";
        string resultFilePath = @"/home/rezishon/storage/Project/Excel_Apps/Excel_To_Access/Data/result.xlsx";
        string resultFileSheetName = "NewSheet";
        Dictionary<string, string> knownColumn = new Dictionary<string, string>()
        {
            { "ساعت ثبت بارنامه", "registeringtime" },
            { "کارت هوشمند وسیله", "trucksmartcardno" },
            {"تاریخ صدور", "registeringdate" },
        };

        #endregion

        #region EPPLUS config

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        #endregion

        #region Main file init

        using var mainPackage = new ExcelPackage(new FileInfo(mainFilePath));
        var mainWorksheet = mainPackage.Workbook.Worksheets[0];
        int mainFileColumnCount = mainWorksheet.Dimension.Columns;
        int mainFileRowCount = mainWorksheet.Dimension.Rows;

        #endregion
    }
}
