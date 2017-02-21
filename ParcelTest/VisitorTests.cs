using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ParcelTest
{
    [TestClass]
    public class VisitorTests
    {
        [TestMethod]
        public void ExtractConstantsTest()
        {
            var mwb = MockWorkbook.standardMockWorkbook();
            var e = mwb.envForSheet(1);
            var f = "=4/(3.2*SUM(A$4:A$10,B$4:B$10)-22)";

            var constants = Parcel.constantsFromFormula(f, e.Path, e.WorkbookName, e.WorksheetName, ignore_parse_errors: false);

            Assert.IsTrue(Array.Exists(constants, c => c.Equals(new AST.ReferenceConstant(e, 3.2))));
            Assert.IsTrue(Array.Exists(constants, c => c.Equals(new AST.ReferenceConstant(e, 4))));
            Assert.IsTrue(Array.Exists(constants, c => c.Equals(new AST.ReferenceConstant(e, 22))));
            Assert.IsTrue(constants.Length == 3);
        }
    }
}
