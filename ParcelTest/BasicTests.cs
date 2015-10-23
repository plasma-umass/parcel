using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.FSharp.Core;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ParcelTest
{
    [TestClass]
    public class BasicTests
    {
        FSharpOption<string> NONESTR = FSharpOption<string>.None;

        private AST.Address makeAddressForA1(string col, int row, MockWorkbook mwb)
        {
            return AST.Address.NewFromA1(
                row,
                col,
                FSharpOption<string>.Some(mwb.GetWorksheet(1).Name),
                FSharpOption<string>.Some(mwb.GetWorkbook().Name),
                mwb.MaybeGetPath()
            );
        }

        private AST.COMRef makeCOMRefForFormula(string formula, MockWorkbook mwb)
        {
            var addr = AST.Address.NewFromR1C1(
                1,
                1,
                FSharpOption<string>.Some(mwb.GetWorksheet(1).Name),
                FSharpOption<string>.Some(mwb.GetWorkbook().Name),
                mwb.MaybeGetPath());
            var rng = new AST.Range(addr, addr);
            Excel.Range com = rng.GetCOMObject(mwb.GetApplication());
            Excel.Worksheet ws = com.Worksheet;
            Excel.Workbook wb = ws.Parent;
            string wsname = ws.Name;
            string wbname = wb.Name;
            var path = new Microsoft.FSharp.Core.FSharpOption<string>(wb.Path);
            int width = com.Columns.Count;
            int height = com.Rows.Count;
            var frm = FSharpOption<string>.Some(formula);

            AST.COMRef c = new AST.COMRef(rng.getUniqueID(), wb, ws, com, path, wbname, wsname, frm, width, height);
            
            return c;
        }


        [TestMethod]
        public void StandardAddress()
        {
            String s = "A3";

            AST.Reference r = ExcelParser.SimpleReferenceParser(s);
            AST.Reference correct = new AST.ReferenceAddress(NONESTR, NONESTR, NONESTR, AST.Address.NewFromA1(3, "A", NONESTR, NONESTR, NONESTR));
            Assert.AreEqual(r, correct);
        }

        [TestMethod]
        public void StandardRange()
        {
            String s = "A3:B22";

            AST.Reference r = ExcelParser.SimpleReferenceParser(s);
            AST.Reference correct = new AST.ReferenceRange(NONESTR,
                                                           NONESTR,
                                                           NONESTR,
                                                           new AST.Range(AST.Address.NewFromA1(3, "A", NONESTR, NONESTR, NONESTR),
                                                                         AST.Address.NewFromA1(22, "B", NONESTR, NONESTR, NONESTR)
                                                                        )
                                                          );
            Assert.AreEqual(r, correct);
        }

        [TestMethod]
        public void IFStatement()
        {
            using (var mwb = new MockWorkbook())
            {
                String s = "=IF(SUM(A1:A5) = 10, \"yes\", \"no\")";

                IEnumerable<AST.Range> rngs = new List<AST.Range>();
                try
                {
                    rngs = ExcelParserUtility.GetRangeReferencesFromFormulaRaw(s, mwb.MaybeGetPath(), mwb.GetWorkbook(), mwb.GetWorksheet(1), false);
                }
                catch (ExcelParserUtility.ParseException e)
                {
                    Assert.Fail(String.Format("\"{0}\" should parse.", s));
                }

                Assert.AreEqual(1, rngs.Count());

                var addr = rngs.First().A1Local();
                Assert.AreEqual("A1:A5", addr);
            }
        }

        [TestMethod]
        public void UnOpAndBinOp()
        {
            using (var mwb = new MockWorkbook())
            {
                String formula = "=(+E6+E7)*0.28";
                var cr = makeCOMRefForFormula(formula, mwb);

                IEnumerable<AST.Address> addrs = new List<AST.Address>();
                try
                {
                    addrs = ExcelParserUtility.GetSingleCellReferencesFromFormula(cr, false);
                }
                catch (ExcelParserUtility.ParseException e)
                {
                    Assert.Fail(String.Format("\"{0}\" should parse.", formula));
                }

                Assert.AreEqual(2, addrs.Count());

                var a1 = makeAddressForA1("E", 6, mwb);
                var a2 = makeAddressForA1("E", 7, mwb);

                Assert.AreEqual(true, addrs.Contains(a1));
                Assert.AreEqual(true, addrs.Contains(a2));
            }
        }

        [TestMethod]
        public void BrutalEUSESTest()
        {
            var failures = new System.Collections.Concurrent.ConcurrentQueue<string>();

            using (var mwb = new MockWorkbook())
            {
                var formulas = System.IO.File.ReadAllLines(@"..\..\TestData\formulas_distinct.txt");
                System.Threading.Tasks.Parallel.ForEach(formulas, f =>
                {
                    try
                    {
                        ExcelParserUtility.ParseFormula(f, mwb.MaybeGetPath(), mwb.GetWorkbook(), mwb.GetWorksheet(1));
                        }
                        catch (Exception e)
                        {
                            if (e is AST.IndirectAddressingNotSupportedException)
                            {
                                // OK   
                            }
                            else if (e is ExcelParserUtility.ParseException)
                            {
                                System.Diagnostics.Debug.WriteLine("Fail: " + f);
                                failures.Enqueue(f);
                            }
                        }
                    });
                }

                Assert.AreEqual(0, failures.Count);
                if (failures.Count > 0)
                {
                    String.Join("\n", failures);
                }
            }

        [TestMethod]
        public void IndirectReferences()
        {
            using (var mwb = new MockWorkbook())
            {
                String formula = "=TRANSPOSE(INDIRECT(ADDRESS(1,4,3,1,Menus!$K$10)):INDIRECT(ADDRESS(1,256,3,1,Menus!$K$10)))";
                try
                {
                    ExcelParserUtility.ParseFormula(formula, mwb.MaybeGetPath(), mwb.GetWorkbook(), mwb.GetWorksheet(1));
                }
                catch (Exception e)
                {
                    if (e is AST.IndirectAddressingNotSupportedException)
                    {
                        // we pass
                    }
                    else
                    {
                        Assert.Fail(e.Message);
                    }
                }
            }
        }

        [TestMethod]
        public void WorksheetNameQuoteEscaping()
        {
            using (var mwb = new MockWorkbook())
            {
                String formula1 = "=Dan!H45";
                String formula2 = "='Dan Stuff'!H45";
                String formula3 = "='Dan''s Stuff'!H45";
                String formula4 = "='Dan's Stuff'!H45";

                // first, the good ones
                try
                {
                    ExcelParserUtility.ParseFormula(formula1, mwb.MaybeGetPath(), mwb.GetWorkbook(), mwb.GetWorksheet(1));
                    ExcelParserUtility.ParseFormula(formula2, mwb.MaybeGetPath(), mwb.GetWorkbook(), mwb.GetWorksheet(1));
                    ExcelParserUtility.ParseFormula(formula3, mwb.MaybeGetPath(), mwb.GetWorkbook(), mwb.GetWorksheet(1));
                }
                catch (ExcelParserUtility.ParseException e)
                {
                    Assert.Fail(e.Message);
                }

                // a bad one
                try
                {
                    ExcelParserUtility.ParseFormula(formula4, mwb.MaybeGetPath(), mwb.GetWorkbook(), mwb.GetWorksheet(1));
                }
                catch (ExcelParserUtility.ParseException e)
                {
                    // OK
                }
            }
        }
    }
}
