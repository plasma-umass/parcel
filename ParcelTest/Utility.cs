using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FSC = Microsoft.FSharp.Collections;

namespace ParcelTest
{
    public static class Utility
    {
        public static AST.Range makeRangeForA1(string rng, AST.Env e)
        {
            var addrs = rng.Split(':');

            return new AST.Range(
                makeAddressForA1(addrs[0], e),
                makeAddressForA1(addrs[1], e)
            );
        }

        public static AST.Address makeAddressForA1(string addr, AST.Env e)
        {
            var r = new System.Text.RegularExpressions.Regex(@"(\$?)([A-Z]+)(\$?)([0-9]+)");

            var matches = r.Match(addr);

            var colMode = String.IsNullOrEmpty(matches.Groups[1].Value) ? AST.AddressMode.Relative : AST.AddressMode.Absolute;
            var col = matches.Groups[2].Value;
            var rowMode = String.IsNullOrEmpty(matches.Groups[3].Value) ? AST.AddressMode.Relative : AST.AddressMode.Absolute;
            var row = Convert.ToInt32(matches.Groups[4].Value);

            return AST.Address.fromA1withMode(row, col, rowMode, colMode, e.WorksheetName, e.WorkbookName, e.Path);
        }

        public static AST.Address makeAddressForA1(string col, int row, AST.Env env)
        {
            return AST.Address.fromA1withMode(
                row,
                col,
                AST.AddressMode.Relative,
                AST.AddressMode.Relative,
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
