﻿using System;
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

                IEnumerable<Excel.Range> rngs = new List<Excel.Range>();
                try
                {
                    rngs = ExcelParserUtility.GetReferencesFromFormula(s, mwb.GetWorkbook(), mwb.GetWorksheet(1));
                }
                catch (ExcelParserUtility.ParseException e)
                {
                    Assert.Fail(String.Format("\"{0}\" should parse.", s));
                }

                Assert.AreEqual(1, rngs.Count());

                var addr = rngs.First().Address;
                Assert.AreEqual("$A$1:$A$5", addr);
            }
        }

        [TestMethod]
        public void UnOpAndBinOp()
        {
            using (var mwb = new MockWorkbook())
            {
                String formula = "=(+E6+E7)*0.28";

                IEnumerable<AST.Address> addrs = new List<AST.Address>();
                try
                {
                    addrs = ExcelParserUtility.GetSingleCellReferencesFromFormula(formula, mwb.GetWorkbook(), mwb.GetWorksheet(1));
                }
                catch (ExcelParserUtility.ParseException e)
                {
                    Assert.Fail(String.Format("\"{0}\" should parse.", formula));
                }

                Assert.AreEqual(2, addrs.Count());

                var a1 = AST.Address.AddressFromCOMObject(mwb.GetWorksheet(1).get_Range("E6"), mwb.GetWorkbook());
                var a2 = AST.Address.AddressFromCOMObject(mwb.GetWorksheet(1).get_Range("E7"), mwb.GetWorkbook());

                Assert.AreEqual(true, addrs.Contains(a1));
                Assert.AreEqual(true, addrs.Contains(a2));
            }
        }

        [TestMethod]
        public void BrutalEUSESTest()
        {
            using (var mwb = new MockWorkbook())
            {
                var formulas = System.IO.File.ReadAllLines(@"..\..\TestData\formulas_distinct.txt");
                int count = 0;
                foreach (String f in formulas)
                {
                    try
                    {
                        ExcelParserUtility.ParseFormula(f, "", mwb.GetWorkbook(), mwb.GetWorksheet(1));
                        // show progress
                        System.Diagnostics.Debug.WriteLine(String.Format("{0}", count++));
                    }
                    catch (ExcelParserUtility.ParseException e)
                    {
                        Assert.Fail(String.Format("\"{0}\" should parse.", f));
                    }
                }
            }
        }
    }
}