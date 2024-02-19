using System;
using System.IO;
using OfficeOpenXml;

namespace ExcelFileCategorization
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string directoryPath = @"S:\New folder (2)\"; // Specify your directory path
            string fileExtension = ".xlsx"; // Specify the file extension

            ExcelPackage combinedData = new ExcelPackage();

            // Get all Excel files in the specified directory
            string[] excelFiles = Directory.GetFiles(directoryPath, $"*{fileExtension}");
        }
    }
}
