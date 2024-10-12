using System.Text;
using ClosedXML.Excel;
namespace Excel_To_SqlQ;

class Program
{
    static void Main(string[] args)
    {
        using var workbook = new XLWorkbook("./all1.xlsx");
        var worksheet = workbook.Worksheet(1);
    }
}
