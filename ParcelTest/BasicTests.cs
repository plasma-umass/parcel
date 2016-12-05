﻿using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Linq;
using ArgList = Microsoft.FSharp.Collections.FSharpList<AST.Expression>;
using ExprOpt = Microsoft.FSharp.Core.FSharpOption<AST.Expression>;
using Expr = AST.Expression;

namespace ParcelTest
{
    [TestClass]
    public class BasicTests
    {
        [TestMethod]
        public void standardAddress()
        {
            var mwb = MockWorkbook.standardMockWorkbook();
            var e = mwb.envForSheet(1);
            String s = "A3";

            AST.Reference r = Parcel.simpleReferenceParser(s, e);
            AST.Reference correct = new AST.ReferenceAddress(e, AST.Address.fromA1withMode(3, "A", AST.AddressMode.Relative, AST.AddressMode.Relative, e.WorksheetName, e.WorkbookName, e.Path));
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
                                                           new AST.Range(Utility.makeAddressForA1("A", 3, e),
                                                                         Utility.makeAddressForA1("B", 22, e))
                                                          );
            Assert.AreEqual(r, correct);
        }

        [TestMethod]
        public void relativeAddress()
        {
            var addrstr = "A1";
            var worksheet = "Worksheet";
            var workbook = "Workbook";
            var path = "";

            var addr = AST.Address.FromA1String(addrstr, worksheet, workbook, path);

            Assert.IsTrue(addr.RowMode == AST.AddressMode.Relative);
            Assert.IsTrue(addr.ColMode == AST.AddressMode.Relative);
        }

        [TestMethod]
        public void absoluteAddress()
        {
            var addrstr = "$A$1";
            var worksheet = "Worksheet";
            var workbook = "Workbook";
            var path = "";

            var addr = AST.Address.FromA1String(addrstr, worksheet, workbook, path);

            Assert.IsTrue(addr.RowMode == AST.AddressMode.Absolute);
            Assert.IsTrue(addr.ColMode == AST.AddressMode.Absolute);
        }

        [TestMethod]
        public void halfAbsoluteAddress()
        {
            var addrstr = "A$1";
            var worksheet = "Worksheet";
            var workbook = "Workbook";
            var path = "";

            var addr = AST.Address.FromA1String(addrstr, worksheet, workbook, path);

            Assert.IsTrue(addr.RowMode == AST.AddressMode.Absolute);
            Assert.IsTrue(addr.ColMode == AST.AddressMode.Relative);
        }

        [TestMethod]
        public void halfSharkAlligatorHalfMan()
        {
            var addrstr = "$A1";
            var worksheet = "Worksheet";
            var workbook = "Workbook";
            var path = "";

            var addr = AST.Address.FromA1String(addrstr, worksheet, workbook, path);

            Assert.IsTrue(addr.RowMode == AST.AddressMode.Relative);
            Assert.IsTrue(addr.ColMode == AST.AddressMode.Absolute);
        }

        [TestMethod]
        public void basicIfExpression()
        {
            var mwb = MockWorkbook.standardMockWorkbook();
            var e = mwb.envForSheet(1);
            string f = "=IF(TRUE, 1, 2)";

            ExprOpt asto = Parcel.parseFormula(f, e.Path, e.WorkbookName, e.WorksheetName);

            Expr[] a = {
                Expr.NewReferenceExpr(new AST.ReferenceBoolean(e, true)),
                Expr.NewReferenceExpr(new AST.ReferenceConstant(e, 1.0)),
                Expr.NewReferenceExpr(new AST.ReferenceConstant(e, 2.0))
            };
            ArgList args = Utility.makeFSList<AST.Expression>(a);
            Expr correct = Expr.NewReferenceExpr(new AST.ReferenceFunction(e, "IF", args, AST.Arity.NewFixed(3)));

            try
            {
                Expr ast = asto.Value;
                Assert.AreEqual(ast, correct);
            } catch (NullReferenceException nre)
            {
                Assert.Fail("Parse error: " + nre.Message);
            }
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
            catch (Parcel.ParseException)
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
            catch (Parcel.ParseException)
            {
                Assert.Fail(String.Format("\"{0}\" should parse.", formula));
            }

            Assert.AreEqual(2, addrs.Count());

            var a1 = Utility.makeAddressForA1("E", 6, mwb.envForSheet(1));
            var a2 = Utility.makeAddressForA1("E", 7, mwb.envForSheet(1));

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
            catch (Parcel.ParseException)
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
            var f1a1 = Utility.makeAddressForA1("L", 66, mwb.envForSheet(1));
            var f1a2 = Utility.makeAddressForA1("J", 27, xmwb.envForSheet(1));

            // extract
            try
            {
                var addrs = Parcel.addrReferencesFromFormula(f1, mwb.Path, mwb.WorkbookName, mwb.worksheetName(1), false);
                Assert.IsTrue(addrs.Contains(f1a1));
                Assert.IsTrue(addrs.Contains(f1a2));
                Assert.IsTrue(addrs.Length == 2);
            }
            catch (Parcel.ParseException e)
            {
                Assert.Fail(e.Message);
            }
        }

        [TestMethod]
        public void missingWorkbookAddrExtraction()
        {
            var mwb = new MockWorkbook("C:\\FOOBAR", "workbook.xls", new[] { "budget" });
            var f = "=budget!A43";
            var addr = Utility.makeAddressForA1("A", 43, mwb.envForSheet(1));

            // extract
            try
            {
                var addrs = Parcel.addrReferencesFromFormula(f, mwb.Path, mwb.WorkbookName, mwb.worksheetName(1), false);
                Assert.IsTrue(addrs.Contains(addr));
                Assert.IsTrue(addrs.Length == 1);
            }
            catch (Parcel.ParseException e)
            {
                Assert.Fail(e.Message);
            }
        }

        [TestMethod]
        public void crossWorksheetRangeExtraction()
        {
            var mwb = new MockWorkbook("C:\\FOOBAR", "workbook.xls", new[] { "sheet1", "One Country Charts", "One Country Data" });
            var f = "=IF('One Country Charts'!F17=\"\",VLOOKUP('One Country Charts'!F11,'One Country Data'!M5:O187,3,FALSE),VLOOKUP('One Country Charts'!F17,'One Country Data'!M5:O187,3,FALSE))";

            var data_env = mwb.envForSheet(3);

            var rng = new AST.Range(Utility.makeAddressForA1("M", 5, data_env), Utility.makeAddressForA1("O", 187, data_env));

            // extract
            try
            {
                var rngs = Parcel.rangeReferencesFromFormula(f, mwb.Path, mwb.WorkbookName, mwb.worksheetName(1), false);
                Assert.IsTrue(rngs.Contains(rng));
                Assert.IsTrue(rngs.Length == 1);
            }
            catch (Parcel.ParseException e)
            {
                Assert.Fail(e.Message);
            }
        }

        [TestMethod]
        public void multipleRangeExtraction()
        {
            var mwb = new MockWorkbook("C:\\FOOBAR", "workbook.xls", new[] { "sheet1", "Calculations", "Status" });
            var f = "=IF(Status!G11=\"stand-alone hub\",SUMIF(Calculations!B7:P7,\"include\",Calculations!B163:P163),IF(Status!G11=\"remote\",SUMIF(Calculations!B7:P7,\"include\",Calculations!B184:P184),SUMIF(Calculations!B7:P7,\"include\",Calculations!B205:P205)))";

            var calc_env = mwb.envForSheet(2);

            var rng1 = new AST.Range(Utility.makeAddressForA1("B", 7, calc_env), Utility.makeAddressForA1("P", 7, calc_env));
            var rng2 = new AST.Range(Utility.makeAddressForA1("B", 163, calc_env), Utility.makeAddressForA1("P", 163, calc_env));
            var rng3 = new AST.Range(Utility.makeAddressForA1("B", 184, calc_env), Utility.makeAddressForA1("P", 184, calc_env));
            var rng4 = new AST.Range(Utility.makeAddressForA1("B", 205, calc_env), Utility.makeAddressForA1("P", 205, calc_env));

            // extract
            try
            {
                var rngs = Parcel.rangeReferencesFromFormula(f, mwb.Path, mwb.WorkbookName, mwb.worksheetName(1), false);
                Assert.IsTrue(rngs.Contains(rng1));
                Assert.IsTrue(rngs.Contains(rng2));
                Assert.IsTrue(rngs.Contains(rng3));
                Assert.IsTrue(rngs.Contains(rng4));
                Assert.IsTrue(rngs.Length == 4);
            }
            catch (Parcel.ParseException e)
            {
                Assert.Fail(e.Message);
            }
        }

        [TestMethod]
        public void testHLOOKUP()
        {
            var mwb = new MockWorkbook("C:\\FOOBAR", "workbook.xls", new[] { "sheet1", "Calculations", "Status" });

            var f = "=HLOOKUP(J9,$C$45:$AH$46,2,TRUE)";

            // parse
            try
            {
                var ast = Parcel.parseFormula(f, mwb.Path, mwb.WorkbookName, mwb.worksheetName(1));
                Assert.IsTrue(ExprOpt.get_IsSome(ast));
                var expr = (Expr.ReferenceExpr)ast.Value;
                var formula = (AST.ReferenceFunction)expr.Item;

                Assert.IsTrue(formula.FunctionName == "HLOOKUP");
            }
            catch (Parcel.ParseException e)
            {
                Assert.Fail(e.Message);
            }

        }

        [TestMethod]
        public void testCOUNTA()
        {
            // a parse that failed in the wild:
            var mwb = new MockWorkbook("C:\\FOOBAR", "workbook.xls", new[] { "sheet1", "Calculations", "Status" });
            var f = "=COUNTA(A6:A33)";

            // parse
            try
            {
                var ast = Parcel.parseFormula(f, mwb.Path, mwb.WorkbookName, mwb.worksheetName(1));
                Assert.IsTrue(ExprOpt.get_IsSome(ast));
                var expr = (Expr.ReferenceExpr)ast.Value;
                var formula = (AST.ReferenceFunction)expr.Item;

                Assert.IsTrue(formula.FunctionName == "COUNTA");
            }
            catch (Parcel.ParseException e)
            {
                Assert.Fail(e.Message);
            }
        }

        [TestMethod]
        public void longestMatchFirstTest()
        {
            string[] isa = { "COLUMN", "FALSE", "NA", "NOW", "PI", "RAND", "ROW", "SHEET", "SHEETS", "TODAY", "TRUE" };
            string[] esa = { "COLUMN", "SHEETS", "FALSE", "SHEET", "TODAY", "RAND", "TRUE", "NOW", "ROW", "NA", "PI" };
            string[] asa = Grammar.lmf(isa);
            Assert.IsTrue(esa.SequenceEqual(asa));
        }

        [TestMethod]
        public void testMAXVarArgs()
        {
            // a parse that failed in the wild:
            var mwb = new MockWorkbook("C:\\FOOBAR", "workbook.xls", new[] { "sheet1", "Calculations", "Status" });
            var f = "=MAX(B6-10,0)";

            // parse
            try
            {
                var ast = Parcel.parseFormula(f, mwb.Path, mwb.WorkbookName, mwb.worksheetName(1));
                Assert.IsTrue(ExprOpt.get_IsSome(ast));
                var expr = (Expr.ReferenceExpr)ast.Value;
                var formula = (AST.ReferenceFunction)expr.Item;

                Assert.IsTrue(formula.FunctionName == "MAX");
                Assert.AreEqual(formula.ArgumentList.Length, 2);
                Assert.IsTrue(formula.ArgumentList.First().IsBinOpExpr);
                Assert.IsTrue(formula.ArgumentList.Last().IsReferenceExpr);
            }
            catch (Parcel.ParseException e)
            {
                Assert.Fail(e.Message);
            }
        }

        [TestMethod]
        public void testTODAY()
        {
            // a parse that failed in the wild:
            var mwb = new MockWorkbook("C:\\FOOBAR", "workbook.xls", new[] { "sheet1", "Calculations", "Status" });
            var f = "=TODAY()";

            // parse
            try
            {
                var ast = Parcel.parseFormula(f, mwb.Path, mwb.WorkbookName, mwb.worksheetName(1));
                Assert.IsTrue(ExprOpt.get_IsSome(ast));
                var expr = (Expr.ReferenceExpr)ast.Value;
                var formula = (AST.ReferenceFunction)expr.Item;

                Assert.IsTrue(formula.FunctionName == "TODAY");
                Assert.AreEqual(formula.ArgumentList.Length, 0);
            }
            catch (Parcel.ParseException e)
            {
                Assert.Fail(e.Message);
            }
        }

        [TestMethod]
        public void testTODAY2()
        {
            // a parse that should fail:
            var mwb = new MockWorkbook("C:\\FOOBAR", "workbook.xls", new[] { "sheet1", "Calculations", "Status" });
            var f = "=TODAY";

            // parse
            try
            {
                var ast = Parcel.parseFormula(f, mwb.Path, mwb.WorkbookName, mwb.worksheetName(1));
                Assert.Fail("Should not parse.");
            }
            catch (Parcel.ParseException e)
            {
                // pass
            }
        }
    }
}
