using System;
using System.Linq;
using FSCore = Microsoft.FSharp.Core;

namespace ParcelTest
{
    public class MockWorkbook
    {
        private string _path;
        private string _workbook_name;
        private string[] _sheet_names;

        public MockWorkbook(string path, string workbook_name, string[] sheet_names)
        {
            _path = path;
            _workbook_name = workbook_name;
            _sheet_names = sheet_names;
        }
        public AST.Defaults defaultsForSheet(int idx)
        {
            return new AST.Defaults(_path, _workbook_name, _sheet_names[idx + 1]);
        }
        public string Path { 
            get { return _path; }
        }
        public string WorkbookName
        {
            get { return _workbook_name; }
        }
        public string worksheetName(int idx) {
            // mimic an Excel one-based array
            return _sheet_names[idx + 1];
        }
        public static int testGetRanges(string formula)
        {
            // mock workbook object
            var mwb = standardMockWorkbook();
            var ranges = Parcel.rangeReferencesFromFormula(
                formula,
                mwb.Path,
                mwb.WorkbookName,
                mwb.worksheetName(1),
                false
                );

            return ranges.Count();
        }

        public static MockWorkbook standardMockWorkbook()
        {
            return new MockWorkbook(
                System.IO.Directory.GetCurrentDirectory(),
                "MockWorkbook.xls",
                new [] { "sheet1", "sheet2", "sheet3" }
                );
        }
    }
}
