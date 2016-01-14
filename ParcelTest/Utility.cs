using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ParcelTest
{
    static class Utility
    {
        public static AST.Address makeAddressForA1(string col, int row, AST.Env env)
        {
            return AST.Address.fromA1(
                row,
                col,
                env.WorksheetName,
                env.WorkbookName,
                env.Path
            );
        }
    }
}
