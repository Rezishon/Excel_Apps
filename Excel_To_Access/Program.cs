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

        #region Result file init

        using var newPackage = new ExcelPackage();
        var newWorksheet = newPackage.Workbook.Worksheets.Add(resultFileSheetName);
        int newWorksheetColumn = 1;

        #endregion

        #region New file building

        // Loops over columns
        for (int column = 1; column <= mainFileColumnCount; column++)
        {
            // Finding specific columns base on app setting
            if (knownColumn.ContainsKey(mainWorksheet.Cells[1, column].GetCellValue<string>() ?? string.Empty))
            {

                // Set result file header
                newWorksheet.Cells[1, newWorksheetColumn].Value = knownColumn[mainWorksheet.Cells[1, column].GetCellValue<string>()];

                // Set result file rows for each column
                for (int row = 2; row <= mainFileRowCount; row++)
                {
                    newWorksheet.Cells[row, newWorksheetColumn].Value = mainWorksheet.Cells[row, column].Value;
                }

                // Move result file column to the next one
                newWorksheetColumn++;
            }
        }

        #endregion

        #region Save result file

        // Replace Old file if exist
        if (File.Exists(resultFilePath))
        {
            File.Copy(resultFilePath, $"{resultFilePath}.del");
            File.Delete(resultFilePath);
        }

        // Save the result file (the new file)
        var fileInfo = new FileInfo(resultFilePath);
        newPackage.SaveAs(fileInfo);

        #endregion

    }
}
