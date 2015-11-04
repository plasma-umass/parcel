using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Linq;

namespace ParcelTest
{
    [TestClass]
    public class BasicTests
    {
        private AST.Address makeAddressForA1(string col, int row, MockWorkbook mwb)
        {
            return AST.Address.fromA1(
                row,
                col,
                mwb.worksheetName(1),
                mwb.WorkbookName,
                mwb.Path
            );
        }
        
        [TestMethod]
        public void standardAddress()
        {
            var mwb = MockWorkbook.standardMockWorkbook();
            var e = mwb.envForSheet(1);
            String s = "A3";

            AST.Reference r = Parcel.simpleReferenceParser(s, e);
            AST.Reference correct = new AST.ReferenceAddress(e, AST.Address.fromA1(3, "A", e.WorksheetName, e.WorkbookName, e.Path));
            Assert.AreEqual(r, correct);
        }

        [TestMethod]
        public void standardRange()
        {
            var mwb = MockWorkbook.standardMockWorkbook();
            var e = mwb.envForSheet(1);
            String s = "A3:B22";

            AST.Reference r = Parcel.simpleReferenceParser(s, e);
            AST.Reference correct = new AST.ReferenceRange(e,
                                                           new AST.Range(makeAddressForA1("A", 3, mwb),
                                                                         makeAddressForA1("B", 22, mwb))
                                                          );
            Assert.AreEqual(r, correct);
        }

        [TestMethod]
        public void ifStatement()
        {
            var mwb = MockWorkbook.standardMockWorkbook();
            string s = "=IF(SUM(A1:A5) = 10, \"yes\", \"no\")";

            IEnumerable<AST.Range> rngs = new List<AST.Range>();
            try
            {
                rngs = Parcel.rangeReferencesFromFormula(s, mwb.Path, mwb.WorkbookName, mwb.worksheetName(1), false);
            }
            catch (Parcel.ParseException e)
            {
                Assert.Fail(String.Format("\"{0}\" should parse.", s));
            }

            Assert.AreEqual(1, rngs.Count());

            var addr = rngs.First().A1Local();
            Assert.AreEqual("A1:A5", addr);
        }

        [TestMethod]
        public void unOpAndBinOp()
        {
            var mwb = MockWorkbook.standardMockWorkbook();
            string formula = "=(+E6+E7)*0.28";

            IEnumerable<AST.Address> addrs = new List<AST.Address>();
            try
            {
                addrs = Parcel.addrReferencesFromFormula(formula, mwb.Path, mwb.WorkbookName, mwb.worksheetName(1), false);
            }
            catch (Parcel.ParseException e)
            {
                Assert.Fail(String.Format("\"{0}\" should parse.", formula));
            }

            Assert.AreEqual(2, addrs.Count());

            var a1 = makeAddressForA1("E", 6, mwb);
            var a2 = makeAddressForA1("E", 7, mwb);

            Assert.AreEqual(true, addrs.Contains(a1));
            Assert.AreEqual(true, addrs.Contains(a2));
        }

        [TestMethod]
        [Ignore]    // temporarily ignore test
        public void brutalEUSESTest()
        {
            var failures = new System.Collections.Concurrent.ConcurrentQueue<string>();

            var mwb = MockWorkbook.standardMockWorkbook();

            var formulas = System.IO.File.ReadAllLines(@"..\..\TestData\formulas_distinct.txt");

            System.Threading.Tasks.Parallel.ForEach(formulas, f =>
            {
                try
                {
                    Parcel.parseFormula(f, mwb.Path, mwb.WorkbookName, mwb.worksheetName(1));
                }
                catch (Exception e)
                {
                    if (e is AST.IndirectAddressingNotSupportedException)
                    {
                        // OK   
                    }
                    else if (e is Parcel.ParseException)
                    {
                        System.Diagnostics.Debug.WriteLine("Fail: " + f);
                        failures.Enqueue(f);
                    }
                }
            });

            Assert.AreEqual(0, failures.Count());
            if (failures.Count > 0)
            {
                String.Join("\n", failures);
            }
        }

        [TestMethod]
        public void indirectReferences()
        {
            var mwb = MockWorkbook.standardMockWorkbook();
            String formula = "=TRANSPOSE(INDIRECT(ADDRESS(1,4,3,1,Menus!$K$10)):INDIRECT(ADDRESS(1,256,3,1,Menus!$K$10)))";
            try
            {
                Parcel.parseFormula(formula, mwb.Path, mwb.WorkbookName, mwb.worksheetName(1));
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

        [TestMethod]
        public void worksheetNameQuoteEscaping()
        {
            var mwb = MockWorkbook.standardMockWorkbook();
            String formula1 = "=Dan!H45";
            String formula2 = "='Dan Stuff'!H45";
            String formula3 = "='Dan''s Stuff'!H45";
            String formula4 = "='Dan's Stuff'!H45";

            // first, the good ones
            try
            {
                Parcel.parseFormula(formula1, mwb.Path, mwb.WorkbookName, mwb.worksheetName(1));
                Parcel.parseFormula(formula2, mwb.Path, mwb.WorkbookName, mwb.worksheetName(1));
                Parcel.parseFormula(formula3, mwb.Path, mwb.WorkbookName, mwb.worksheetName(1));
            }
            catch (Parcel.ParseException e)
            {
                Assert.Fail(e.Message);
            }

            // a bad one
            try
            {
                Parcel.parseFormula(formula4, mwb.Path, mwb.WorkbookName, mwb.worksheetName(1));
            }
            catch (Parcel.ParseException e)
            {
                // OK
            }
        }

        [TestMethod]
        public void crossWorkbookAddrExtraction()
        {
            var mwb = MockWorkbook.standardMockWorkbook();
            var xmwb = new MockWorkbook("C:\\FINRES\\FIRMAS\\FORCASTS\\MODELS\\", "models.xls", new[] { "Forecast Assumptions" });

            var f1 = "=L66*('C:\\FINRES\\FIRMAS\\FORCASTS\\MODELS\\[models.xls]Forecast Assumptions'!J27)^0.25";
            var f1a1 = makeAddressForA1("L", 66, mwb);
            var f1a2 = makeAddressForA1("J", 27, xmwb);

            // extract
            try
            {
                var addrs = Parcel.addrReferencesFromFormula(f1, mwb.Path, mwb.WorkbookName, mwb.worksheetName(1), false);
                Assert.IsTrue(addrs.Contains(f1a1));
                Assert.IsTrue(addrs.Contains(f1a2));
            } catch (Parcel.ParseException e)
            {
                Assert.Fail(e.Message);
            }
        }
    }
}
