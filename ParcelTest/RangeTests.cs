using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PairList = Microsoft.FSharp.Collections.FSharpList<System.Tuple<AST.Address, AST.Address>>;
using AddrPair = System.Tuple<AST.Address, AST.Address>;

namespace ParcelTest
{
    [TestClass]
    public class RangeTests
    {
        [TestMethod]
        public void RangeCase1Test()
        {
            var mwb = MockWorkbook.standardMockWorkbook();
            var e = mwb.envForSheet(1);
            String s = "A1:A2";

            AST.Reference r = Parcel.simpleReferenceParser(s, e);
            AST.Reference correct =
                new AST.ReferenceRange(
                    e,
                    new AST.Range(
                        AST.Address.fromA1(1, "A", e.WorksheetName, e.WorkbookName, e.Path),
                        AST.Address.fromA1(2, "A", e.WorksheetName, e.WorkbookName, e.Path)
                    )
                );
            Assert.AreEqual(r, correct);
        }

        [TestMethod]
        public void RangeCase2Test()
        {
            var mwb = MockWorkbook.standardMockWorkbook();
            var e = mwb.envForSheet(1);
            String s = "A1,A2";

            AST.Reference r = Parcel.simpleReferenceParser(s, e);
            AST.Reference correct =
                new AST.ReferenceRange(
                    e,
                    new AST.Range(
                        new AST.Range(
                            AST.Address.fromA1(1, "A", e.WorksheetName, e.WorkbookName, e.Path),
                            AST.Address.fromA1(1, "A", e.WorksheetName, e.WorkbookName, e.Path)
                        ),
                        new AST.Range(
                            AST.Address.fromA1(2, "A", e.WorksheetName, e.WorkbookName, e.Path),
                            AST.Address.fromA1(2, "A", e.WorksheetName, e.WorkbookName, e.Path)
                        )
                    )
                );
            Assert.AreEqual(r, correct);
        }

        [TestMethod]
        public void RangeCase3Test()
        {
            var mwb = MockWorkbook.standardMockWorkbook();
            var e = mwb.envForSheet(1);
            String s = "A1,A2:A3";

            AST.Reference r = Parcel.simpleReferenceParser(s, e);
            AST.Reference correct =
                new AST.ReferenceRange(
                    e,
                    new AST.Range(
                        new AST.Range(
                            AST.Address.fromA1(1, "A", e.WorksheetName, e.WorkbookName, e.Path),
                            AST.Address.fromA1(1, "A", e.WorksheetName, e.WorkbookName, e.Path)
                        ),
                        new AST.Range(
                            AST.Address.fromA1(2, "A", e.WorksheetName, e.WorkbookName, e.Path),
                            AST.Address.fromA1(3, "A", e.WorksheetName, e.WorkbookName, e.Path)
                        )
                    )
                );
            Assert.AreEqual(r, correct);
        }

        [TestMethod]
        public void RangeCase4Test()
        {
            var mwb = MockWorkbook.standardMockWorkbook();
            var e = mwb.envForSheet(1);
            String s = "A1:A2,A3";

            AST.Reference r = Parcel.simpleReferenceParser(s, e);
            AST.Reference correct =
                new AST.ReferenceRange(
                    e,
                    new AST.Range(
                        new AST.Range(
                            AST.Address.fromA1(1, "A", e.WorksheetName, e.WorkbookName, e.Path),
                            AST.Address.fromA1(2, "A", e.WorksheetName, e.WorkbookName, e.Path)
                        ),
                        new AST.Range(
                            AST.Address.fromA1(3, "A", e.WorksheetName, e.WorkbookName, e.Path),
                            AST.Address.fromA1(3, "A", e.WorksheetName, e.WorkbookName, e.Path)
                        )
                    )
                );
            Assert.AreEqual(r, correct);
        }

        [TestMethod]
        public void RangeCase5Test()
        {
            var mwb = MockWorkbook.standardMockWorkbook();
            var e = mwb.envForSheet(1);
            String s = "A1:A2,A3:A4";

            AST.Reference r = Parcel.simpleReferenceParser(s, e);
            AST.Reference correct =
                new AST.ReferenceRange(
                    e,
                    new AST.Range(
                        new AST.Range(
                            AST.Address.fromA1(1, "A", e.WorksheetName, e.WorkbookName, e.Path),
                            AST.Address.fromA1(2, "A", e.WorksheetName, e.WorkbookName, e.Path)
                        ),
                        new AST.Range(
                            AST.Address.fromA1(3, "A", e.WorksheetName, e.WorkbookName, e.Path),
                            AST.Address.fromA1(4, "A", e.WorksheetName, e.WorkbookName, e.Path)
                        )
                    )
                );
            Assert.AreEqual(r, correct);
        }

        [TestMethod]
        public void RangeCase6Test()
        {
            var mwb = MockWorkbook.standardMockWorkbook();
            var e = mwb.envForSheet(1);
            String s = "A1,A2,A3,A4,A5";

            AddrPair[] addrpairs = {
                                    new AddrPair(
                                        AST.Address.fromA1(5, "A", e.WorksheetName, e.WorkbookName, e.Path),
                                        AST.Address.fromA1(5, "A", e.WorksheetName, e.WorkbookName, e.Path)
                                    ),
                                    new AddrPair(
                                        AST.Address.fromA1(4, "A", e.WorksheetName, e.WorkbookName, e.Path),
                                        AST.Address.fromA1(4, "A", e.WorksheetName, e.WorkbookName, e.Path)
                                    ),
                                    new AddrPair(
                                        AST.Address.fromA1(3, "A", e.WorksheetName, e.WorkbookName, e.Path),
                                        AST.Address.fromA1(3, "A", e.WorksheetName, e.WorkbookName, e.Path)
                                    ),
                                    new AddrPair(
                                        AST.Address.fromA1(2, "A", e.WorksheetName, e.WorkbookName, e.Path),
                                        AST.Address.fromA1(2, "A", e.WorksheetName, e.WorkbookName, e.Path)
                                    ),
                                    new AddrPair(
                                        AST.Address.fromA1(1, "A", e.WorksheetName, e.WorkbookName, e.Path),
                                        AST.Address.fromA1(1, "A", e.WorksheetName, e.WorkbookName, e.Path)
                                    )
                                   };

            AST.Reference r = Parcel.simpleReferenceParser(s, e);
            AST.Reference correct = new AST.ReferenceRange(e, new AST.Range(addrpairs));
            Assert.AreEqual(r, correct);
        }

        [TestMethod]
        public void RangeCase7Test()
        {
            var mwb = MockWorkbook.standardMockWorkbook();
            var e = mwb.envForSheet(1);
            String s = "A1,A2:A3,A4";

            AddrPair[] addrpairs = {
                                    new AddrPair(
                                        AST.Address.fromA1(1, "A", e.WorksheetName, e.WorkbookName, e.Path),
                                        AST.Address.fromA1(1, "A", e.WorksheetName, e.WorkbookName, e.Path)
                                    ),
                                    new AddrPair(
                                        AST.Address.fromA1(2, "A", e.WorksheetName, e.WorkbookName, e.Path),
                                        AST.Address.fromA1(3, "A", e.WorksheetName, e.WorkbookName, e.Path)
                                    ),
                                    new AddrPair(
                                        AST.Address.fromA1(4, "A", e.WorksheetName, e.WorkbookName, e.Path),
                                        AST.Address.fromA1(4, "A", e.WorksheetName, e.WorkbookName, e.Path)
                                    )
                                   };

            AST.Reference r = Parcel.simpleReferenceParser(s, e);
            AST.Reference correct = new AST.ReferenceRange(e, new AST.Range(addrpairs));
            Assert.AreEqual(r, correct);
        }

        [TestMethod]
        public void rangeCase8Test()
        {
            var mwb = MockWorkbook.standardMockWorkbook();
            var e = mwb.envForSheet(1);

            String s = "A1:B1";

            AST.Range range = new AST.Range(
                AST.Address.fromA1(1, "A", e.WorkbookName, e.WorkbookName, e.Path),
                AST.Address.fromA1(1, "B", e.WorkbookName, e.WorkbookName, e.Path)
                );

            AST.Reference r = Parcel.simpleReferenceParser(s, e);
            AST.Reference correct = new AST.ReferenceRange(e, range);
            Assert.AreEqual(r, correct);
        }
    }
}
