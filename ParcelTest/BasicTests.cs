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
            }
        }
    }
}
