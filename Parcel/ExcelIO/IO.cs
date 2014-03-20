using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Globalization;

namespace ExcelIO
{
    public class InvalidURLZoneException : Exception { }

    public class TrustedPathAlreadySetException : Exception { }

    public class AppWrap : IDisposable
    {
        public class WorkbookWrap : IDisposable
        {
            private Excel.Workbook _wb;
            private List<Excel.Worksheet> _ws = new List<Excel.Worksheet>();
            private List<Excel.Range> _rs = new List<Excel.Range>();
            private Regex _formula_filter = new Regex("^=", RegexOptions.Compiled);

            public WorkbookWrap(Excel.Workbook wb)
            {
                _wb = wb;
            }
            public IEnumerable<string> GetFormulas() {
                var fs = new List<string>();
                foreach (Excel.Worksheet ws in _wb.Worksheets)
                {
                    // track touched workbooks
                    _ws.Add(ws);

                    // track touched ranges
                    Excel.Range ur = ws.UsedRange;
                    _rs.Add(ur);

                    // array read formulas
                    object[,] formulas = ur.Formula;

                    // grab formula strings
                    // danger: one-based array
                    for (int i = 1; i <= formulas.GetLength(0); i++)
                    {
                        for (int j = 1; j <= formulas.GetLength(1); j++)
                        {
                            String s = (String)formulas[i, j];
                            if (!String.IsNullOrEmpty(s) && _formula_filter.IsMatch(s))
                            {
                                fs.Add(s);
                            }
                        }
                    }
                }
                return fs;
            }
            public void Dispose()
            {
                // dispose of ranges
                foreach (Excel.Range r in _rs)
                {
                    Marshal.ReleaseComObject(r);
                }
                _rs = null;

                // dispose of worksheets
                foreach (Excel.Worksheet ws in _ws)
                {
                    Marshal.ReleaseComObject(ws);
                    
                }
                _ws = null;

                // dispose of workbooks
                _wb.Close();
                Marshal.ReleaseComObject(_wb);
                _wb = null;

                // force collection
                GC.Collect();
            }
        }

        private Excel.Application _app;
        public AppWrap()
        {
            _app = new Excel.Application();
        }
        public WorkbookWrap OpenWorkbook(string file) {
            // clear zone data
            IO.DeleteZoneData(file);

            return new WorkbookWrap(IO.OpenWorkbook(file, _app));
        }
        public void Dispose()
        {
            _app.Quit();
        }
    }

    public class IO
    {
        public static readonly string KEY_PREFIX = @"HKEY_CURRENT_USER\";
        public static readonly string KEY_BASE = @"Software\Microsoft\Office\14.0\Excel\Security\Trusted Locations\";

        public enum URLZone
        {
            URLZONE_INVALID = -1,
            URLZONE_PREDEFINED_MIN = 0,
            URLZONE_LOCAL_MACHINE = 0,
            URLZONE_INTRANET,
            URLZONE_TRUSTED,
            URLZONE_INTERNET,
            URLZONE_UNTRUSTED,
            URLZONE_PREDEFINED_MAX = 999,
            URLZONE_USER_MIN = 1000,
            URLZONE_USER_MAX = 10000
        }

        // All of the following private enums are poorly documented
        private enum XlCorruptLoad
        {
            NormalLoad = 0,
            RepairFile = 1,
            ExtractData = 2
        }

        private enum XlUpdateLinks
        {
            Yes = 2,
            No = 0
        }

        private enum XlPlatform
        {
            Macintosh = 1,
            Windows = 2,
            MSDOS = 3
        }

        [DllImport(@"..\..\..\Parcel\Debug\TrustedZoneHandler.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Unicode)]
        private static extern void EraseZoneData(string filename);

        [DllImport(@"..\..\..\Parcel\Debug\TrustedZoneHandler.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Unicode)]
        private static extern uint ReadZoneData(string filename);

        [DllImport(@"..\..\..\Parcel\Debug\TrustedZoneHandler.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Unicode)]
        private static extern uint SetZoneData(string filename, uint url_zone);

        public static URLZone GetZoneData(string filename)
        {
            // ReadZoneData requires an absolute path
            return (URLZone)ReadZoneData(System.IO.Path.GetFullPath(filename));
        }

        public static void DeleteZoneData(string filename)
        {
            // EraseZoneData requires an absolute path
            EraseZoneData(System.IO.Path.GetFullPath(filename));
        }

        public static void UpdateZoneData(string filename, URLZone urlz)
        {
            if (urlz < URLZone.URLZONE_PREDEFINED_MIN || urlz > URLZone.URLZONE_USER_MAX)
            {
                throw new InvalidURLZoneException();
            }

            // EraseZoneData requires an absolute path
            SetZoneData(System.IO.Path.GetFullPath(filename), (uint)urlz);
        }

        /// <summary>
        /// Returns the next available trusted location ID.
        /// </summary>
        /// <returns>An integer representing the trusted location ID. Throws an exception if the key already exists.</returns>
        public static int NextTrustedLocationID(string path)
        {
            var cpath = System.IO.Path.GetFullPath(path);
            Regex r = new Regex("Location([0-9]+)", RegexOptions.Compiled);
            var max = 0;
            var no_locations = true;
            var already_defined = false;

            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(KEY_BASE))
            {
                foreach (var location in key.GetSubKeyNames())
                {
                    no_locations = false;

                    // extract numeric ID
                    var id = System.Convert.ToInt32(r.Match(location).Groups[1].Value);

                    // check to see if the location is already defined
                    var trusted_path = System.IO.Path.GetFullPath((String)Registry.GetValue(KEY_PREFIX + KEY_BASE + "Location" + id, "Path", "empty"));
                    if (cpath == trusted_path)
                    {
                        throw new TrustedPathAlreadySetException();
                    }

                    // update max
                    if (id > max)
                    {
                        max = id;
                    }
                }
            }

            if (no_locations)
            {
                return max;
            }
            else
            {
                return max + 1;
            }
        }

        /// <summary>
        /// Adds a trusted location to the Windows registry.
        /// </summary>
        /// <param name="path">An absolute path string.</param>
        /// <param name="allow_subfolders">Subfolders of the path are trusted when true.</param>
        public static void AddTrustedLocation(string path, bool allow_subfolders, string description)
        {
            int id;
            try
            {
                id = NextTrustedLocationID(path);
            }
            catch (TrustedPathAlreadySetException)
            {
                return;
            }

            var dt = DateTime.Now;
            Registry.SetValue(KEY_PREFIX + KEY_BASE + "Location" + id, "Path", path);
            Registry.SetValue(KEY_PREFIX + KEY_BASE + "Location" + id, "AllowSubfolders", allow_subfolders, RegistryValueKind.DWord);
            Registry.SetValue(KEY_PREFIX + KEY_BASE + "Location" + id, "Date", dt.ToString("g", CultureInfo.CreateSpecificCulture("en-US")));
            Registry.SetValue(KEY_PREFIX + KEY_BASE + "Location" + id, "Description", description);
        }

        public static Excel.Workbook OpenWorkbook(string filename, Excel.Application app)
        {
            // we need to disable all alerts, e.g., password prompts, etc.
            app.DisplayAlerts = false;

            // disable macros
            app.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;

            // This call is stupid.  See:
            // http://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.workbooks.open%28v=office.11%29.aspx
            app.Workbooks.Open(filename,                    // FileName (String)
                               XlUpdateLinks.Yes,           // UpdateLinks (XlUpdateLinks enum)
                               true,                        // ReadOnly (Boolean)
                               Missing.Value,               // Format (int?)
                               "thisisnotapassword",        // Password (String)
                               Missing.Value,               // WriteResPassword (String)
                               true,                        // IgnoreReadOnlyRecommended (Boolean)
                               Missing.Value,               // Origin (XlPlatform enum)
                               Missing.Value,               // Delimiter; if the filetype is txt (String)
                               Missing.Value,               // Editable; not what you think (Boolean)
                               false,                       // Notify (Boolean)
                               Missing.Value,               // Converter(int)
                               false,                       // AddToMru (Boolean)
                               Missing.Value,               // Local; really "use my locale?" (Boolean)
                               XlCorruptLoad.RepairFile);   // CorruptLoad (XlCorruptLoad enum)

            return app.Workbooks[1];
        }

        /// <summary>
        /// Enumerates spreadsheets found at the given path.  Recursively
        /// descends through all directories in the directory hierarchy.
        /// </summary>
        /// <param name="path">The starting point in the search.</param>
        /// <param name="path_filters">Only include matches from directories in the path matching this pattern.</param>
        /// <returns>IEnumerable collection of spreadsheet names</returns>
        public static IEnumerable<string> EnumerateSpreadsheets(string path, Regex path_filter)
        {
            var xlfiles = Directory.EnumerateFiles(path, "*.xls", SearchOption.AllDirectories);
            return xlfiles.Where(xlfile => path_filter.IsMatch(xlfile));
        }
    }
}
