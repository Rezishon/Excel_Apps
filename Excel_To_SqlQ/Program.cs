using System.Text;
using ClosedXML.Excel;
namespace Excel_To_SqlQ;

class Program
{
    static void Main(string[] args)
    {
        using var workbook = new XLWorkbook("./all1.xlsx");
        var worksheet = workbook.Worksheet(1);

        Processes processes = new Processes(worksheet, "Costumer");

    }
}

internal class Processes(IXLWorksheet worksheet, string TableName)
{
    public List<string> Headers { get; set; } = new List<string>();
    public IXLWorksheet Worksheet { get; set; } = worksheet;
    public List<string> RowsData { get; set; } = new List<string>();

    public void Luncher()
    {
    }

    }
}
