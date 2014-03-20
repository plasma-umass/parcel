using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ParcelTest
{
    public class MockWorkbook : IDisposable
    {
        Excel.Application app;
        Excel.Workbook wb;
        Excel.Sheets ws;

        public MockWorkbook()
        {
            // worksheet indices; watch out! the second index here is the NUMBER of elements, NOT the max value!
            var e = Enumerable.Range(1, 10);

            // new Excel instance
            app = new Excel.Application();

            // create new workbook
            wb = app.Workbooks.Add();

            // get a reference to the worksheet array
            // By default, workbooks have three blank worksheets.
            ws = wb.Worksheets;

            // add some worksheets
            foreach (int i in e)
            {
                ws.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            }
        }
        public Excel.Application GetApplication() { return app; }
        public Excel.Workbook GetWorkbook() { return wb; }
        public Excel.Sheets GetWorksheets() { return ws; }
        public Excel.Worksheet GetWorksheet(int idx) { return (Excel.Worksheet)ws[idx]; }
        public static int TestGetRanges(string formula)
        {
            // mock workbook object
            var mwb = new MockWorkbook();
            Excel.Workbook wb = mwb.GetWorkbook();
            Excel.Worksheet ws = mwb.GetWorksheet(1);

            var ranges = ExcelParserUtility.GetReferencesFromFormula(formula, wb, ws);

            return ranges.Count();
        }
        public void Dispose()
        {
            try
            {
                wb.Close(false, Type.Missing, Type.Missing);
                app.Quit();
                Marshal.ReleaseComObject(ws);
                Marshal.ReleaseComObject(wb);
                Marshal.ReleaseComObject(app);
                ws = null;
                wb = null;
                app = null;
            }
            catch
            {
            }
        }
    }
}
