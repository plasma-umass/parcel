using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FSC = Microsoft.FSharp.Collections;

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

        private static FSC.FSharpList<T> mkFSL<T>(T[] arr, int i, FSC.FSharpList<T> fsl)
        {
            if (i >= 0)
            {
                return mkFSL(arr, i - 1, new FSC.FSharpList<T>(arr[i], fsl));
            } else
            {
                return fsl;
            }
        }

        public static FSC.FSharpList<T> makeFSList<T>(T[] arr)
        {
            return mkFSL(arr, arr.Length - 1, FSC.FSharpList<T>.Empty);
        }
    }
}
