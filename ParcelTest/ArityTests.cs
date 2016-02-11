using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

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

            //AST.Range range = new AST.Range(
            //    AST.Address.fromA1(1, "A", e.WorkbookName, e.WorkbookName, e.Path),
            //    AST.Address.fromA1(1, "B", e.WorkbookName, e.WorkbookName, e.Path)
            //    );

            AST.Reference r = Parcel.simpleReferenceParser(f, e);
            //AST.Reference correct = new AST.ReferenceRange(e, range);
            //Assert.AreEqual(r, correct);
        }
    }
}
