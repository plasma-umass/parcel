open System
open System.IO
open System.Text.RegularExpressions
open Microsoft.Office.Interop.Excel
open ExcelIO

let get_formulas(file: string)(a: AppWrap) : seq<string> =
    try
        using(a.OpenWorkbook(file)) (fun (wb) ->
            wb.GetFormulas()
        )
    with
    | :? Exception as e ->
        Console.Error.WriteLine("Exception: {0}", e.Message)
        Seq.empty

[<EntryPoint>]
let main argv = 
    using(new AppWrap()) (fun app ->
        let path = @"C:\EUSES\"
        let r = Regex(@"\\processed\\", RegexOptions.Compiled)

        Seq.iter (fun (file: string) ->
            Console.Error.WriteLine("Opening {0}", file)
            Seq.iter (fun (formula: string) -> Console.WriteLine(formula)) (get_formulas file app)
            Console.Error.WriteLine("Closing {0}", file)
        ) (IO.EnumerateSpreadsheets(path, r))
    )
    0 // return an integer exit code
