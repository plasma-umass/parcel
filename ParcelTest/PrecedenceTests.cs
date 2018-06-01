using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExprOpt = Microsoft.FSharp.Core.FSharpOption<AST.Expression>;
using Expr = AST.Expression;

namespace ParcelTest
{
    [TestClass]
    public class PrecedenceTests
    {
        [TestMethod]
        public void MultiplicationVsAdditionPrecedenceTest()
        {
            var mwb = MockWorkbook.standardMockWorkbook();
            var e = mwb.envForSheet(1);

            var f = "=2*3+1";

            ExprOpt asto = Parcel.parseFormula(f, e.Path, e.WorkbookName, e.WorksheetName);

            Expr correct =
                Expr.NewBinOpExpr(
                    "+",
                    Expr.NewBinOpExpr(
                        "*",
                        Expr.NewReferenceExpr(new AST.ReferenceConstant(e, 2.0)),
                        Expr.NewReferenceExpr(new AST.ReferenceConstant(e, 3.0))
                    ),
                    Expr.NewReferenceExpr(new AST.ReferenceConstant(e, 1.0))
                );

            try
            {
                Expr ast = asto.Value;
                Assert.AreEqual(correct, ast);
            }
            catch (NullReferenceException nre)
            {
                Assert.Fail("Parse error: " + nre.Message);
            }
        }
    }
}
