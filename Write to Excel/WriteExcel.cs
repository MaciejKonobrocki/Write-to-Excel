
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Windows.Forms;

namespace Write_to_Excel
{
    class WriteExcel
    {
        public static void writeExcel()
        {
            string filePath = "C:\\Users\\asdf\\Desktop\\NOWY.xlsx";
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb;
            Worksheet ws;

            wb = excel.Workbooks.Open(filePath);
            ws= wb.Worksheets[1];

            Range cellRange = ws.Range["A1:D1"];
            string[] things = new[] { "a", "b", "c", "d" };
            cellRange.set_Value(XlRangeValueDataType.xlRangeValueDefault, things);

            wb.SaveAs("C:\\Users\\asdf\\Desktop\\NOWY.xlsx");
            wb.Close();
        }

    }
}