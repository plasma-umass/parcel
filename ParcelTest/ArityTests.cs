using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ArgList = Microsoft.FSharp.Collections.FSharpList<AST.Expression>;
using ExprOpt = Microsoft.FSharp.Core.FSharpOption<AST.Expression>;
using Expr = AST.Expression;

namespace ParcelTest
{
    [TestClass]
    public class ArityTests
    {
        [TestMethod]
        public void Arity2Test()
        {
            var mwb = MockWorkbook.standardMockWorkbook();
            var e = mwb.envForSheet(1);

            var f = "=SUMX2MY2(A4:A10,B4:B10)";

            ExprOpt asto = Parcel.parseFormula(f, e.Path, e.WorkbookName, e.WorksheetName);

            Expr[] a = {
                Expr.NewReferenceExpr(new AST.ReferenceRange(e, Utility.makeRangeForA1("A4:A10", e))),
                Expr.NewReferenceExpr(new AST.ReferenceRange(e, Utility.makeRangeForA1("B4:B10", e)))
            };
            ArgList args = Utility.makeFSList<AST.Expression>(a);
            Expr correct = Expr.NewReferenceExpr(new AST.ReferenceFunction(e, "SUMX2MY2", args, AST.Arity.NewFixed(2)));

            try
            {
                Expr ast = asto.Value;
                Assert.AreEqual(ast, correct);
            }
            catch (NullReferenceException nre)
            {
                Assert.Fail("Parse error: " + nre.Message);
            }
        }

        [TestMethod]
        public void Arity2Test2()
        {
            var mwb = MockWorkbook.standardMockWorkbook();
            var e = mwb.envForSheet(1);

            var f = "=SUMX2MY2(A$4:A$10,B$4:B$10)";

            ExprOpt asto = Parcel.parseFormula(f, e.Path, e.WorkbookName, e.WorksheetName);

            Expr[] a = {
                Expr.NewReferenceExpr(new AST.ReferenceRange(e, Utility.makeRangeForA1("A$4:A$10", e))),
                Expr.NewReferenceExpr(new AST.ReferenceRange(e, Utility.makeRangeForA1("B$4:B$10", e)))
            };
            ArgList args = Utility.makeFSList<AST.Expression>(a);
            Expr correct = Expr.NewReferenceExpr(new AST.ReferenceFunction(e, "SUMX2MY2", args, AST.Arity.NewFixed(2)));

            try
            {
                Expr ast = asto.Value;
                Assert.AreEqual(ast, correct);
            }
            catch (NullReferenceException nre)
            {
                Assert.Fail("Parse error: " + nre.Message);
            }
        }
    }
}
