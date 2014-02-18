using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.FSharp.Core;

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
    }
}
