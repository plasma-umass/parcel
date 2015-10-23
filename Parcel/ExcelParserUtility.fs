module ExcelParserUtility
    open FParsec
    type Workbook = Microsoft.Office.Interop.Excel.Workbook
    type Worksheet = Microsoft.Office.Interop.Excel.Worksheet
    type XLRange = Microsoft.Office.Interop.Excel.Range

    // using C#-style exceptions so that the can be handled by C# code
    type ParseException(formula: string, reason: string) =
        inherit System.Exception(formula)
        new(formula: string) = ParseException(formula, "")

    let PuntedFunction(fnname: string) : bool =
        match fnname with
        | "INDEX" -> true
        | "HLOOKUP" -> true
        | "VLOOKUP" -> true
        | "LOOKUP" -> true
        | "OFFSET" -> true
        | _ -> false

    let rec GetRangeReferenceRanges(ref: AST.ReferenceRange) : AST.Range list = [ref.Range]

    and GetFunctionRanges(ref: AST.ReferenceFunction) : AST.Range list =
        if PuntedFunction(ref.FunctionName) then
            []
        else
            List.map (fun arg -> GetExprRanges(arg)) ref.ArgumentList |> List.concat
        
    and GetExprRanges(expr: AST.Expression) : AST.Range list =
        match expr with
        | AST.ReferenceExpr(r) -> GetRanges(r)
        | AST.BinOpExpr(op, e1, e2) -> GetExprRanges(e1) @ GetExprRanges(e2)
        | AST.UnaryOpExpr(op, e) -> GetExprRanges(e)
        | AST.ParensExpr(e) -> GetExprRanges(e)

    and GetRanges(ref: AST.Reference) : AST.Range list =
        match ref with
        | :? AST.ReferenceRange as r -> GetRangeReferenceRanges(r)
        | :? AST.ReferenceAddress -> []
        | :? AST.ReferenceNamed -> []   // TODO: symbol table lookup
        | :? AST.ReferenceFunction as r -> GetFunctionRanges(r)
        | :? AST.ReferenceConstant -> []
        | :? AST.ReferenceString -> []
        | _ -> failwith "Unknown reference type."

    let GetRangeReferencesFromFormulaRaw(formula: string, path: string option, workbook: Workbook, worksheet: Worksheet, ignore_parse_errors: bool) : seq<AST.Range> =
        try
            match ExcelParser.ParseFormula(formula, path, workbook, worksheet),ignore_parse_errors with
            | Some(tree),_ ->
                let refs = GetExprRanges(tree)
                List.map (fun (r: AST.Range) ->
                            // temporarily bail if r refers to object in a different workbook
                            // TODO: we should open the other workbook and continue the analysis
                            let paths = Seq.filter (fun pth -> pth = path) (r.GetPathNames())
                            if Seq.length(paths) = 0 then
                                None
                            else
                                Some(r)
                         ) refs |> List.choose id |> Seq.ofList
            | None,false -> raise (ParseException(formula))
            | None,true -> Seq.empty    // just ignore parse exceptions for now
        with
        // right now, we recognize indirect addresses but do not correctly dereference them
        | :? AST.IndirectAddressingNotSupportedException -> seq[]

    let GetRangeReferencesFromFormula(cr: AST.COMRef, ignore_parse_errors: bool) : seq<AST.Range> =
        GetRangeReferencesFromFormulaRaw(cr.Formula, cr.Path, cr.Workbook, cr.Worksheet, ignore_parse_errors)

    // single-cell variants:
    let rec GetSCExprRanges(expr: AST.Expression) : AST.Address list =
        match expr with
        | AST.ReferenceExpr(r) -> GetSCRanges(r)
        | AST.BinOpExpr(op, e1, e2) -> GetSCExprRanges(e1) @ GetSCExprRanges(e2)
        | AST.UnaryOpExpr(op, e) -> GetSCExprRanges(e)
        | AST.ParensExpr(e) -> GetSCExprRanges(e)

    and GetSCRanges(ref: AST.Reference) : AST.Address list =
        match ref with
        | :? AST.ReferenceRange -> []
        | :? AST.ReferenceAddress as r -> GetSCAddressReferenceRanges(r)
        | :? AST.ReferenceNamed -> []   // TODO: symbol table lookup
        | :? AST.ReferenceFunction as r -> GetSCFunctionRanges(r)
        | :? AST.ReferenceConstant -> []
        | :? AST.ReferenceString -> []
        | _ -> failwith "Unknown reference type."

    and GetSCAddressReferenceRanges(ref: AST.ReferenceAddress) : AST.Address list = [ref.Address]

    and GetSCFunctionRanges(ref: AST.ReferenceFunction) : AST.Address list =
        if PuntedFunction(ref.FunctionName) then
            []
        else
            List.map (fun arg -> GetSCExprRanges(arg)) ref.ArgumentList |> List.concat

    let rec GetFormulaNamesFromExpr(ast: AST.Expression): string list =
        match ast with
        | AST.ReferenceExpr(r) -> GetFormulaNamesFromReference(r)
        | AST.BinOpExpr(op, e1, e2) -> GetFormulaNamesFromExpr(e1) @ GetFormulaNamesFromExpr(e2)
        | AST.UnaryOpExpr(op, e) -> GetFormulaNamesFromExpr(e)
        | AST.ParensExpr(e) -> GetFormulaNamesFromExpr(e)

    and GetFormulaNamesFromReference(ref: AST.Reference): string list =
        match ref with
        | :? AST.ReferenceFunction as r -> [r.FunctionName]
        | _ -> []

    let GetSCFormulaNames(formula: string, path: string, ws: Worksheet, wb: Workbook, ignore_parse_errors: bool) =
        let app = wb.Application
        let path' = Some(System.IO.Path.GetDirectoryName(wb.FullName))
        match ExcelParser.ParseFormula(formula, path', wb, ws),ignore_parse_errors with
        | Some(ast),_ -> GetFormulaNamesFromExpr(ast) |> Seq.ofList
        | None,false -> raise (ParseException formula)
        | None,true -> Seq.empty    // just ignore parse exceptions for now

    let GetSingleCellReferencesFromFormula(cr: AST.COMRef, ignore_parse_errors: bool) : seq<AST.Address> =
        match ExcelParser.ParseFormula(cr.Formula, cr.Path, cr.Workbook, cr.Worksheet),ignore_parse_errors with
        | Some(tree),_ ->
            // temporarily bail if a refers to object in a different workbook
            // TODO: we should open the other workbook and continue the analysis
            let refs = GetSCExprRanges(tree) |> Seq.ofList
            Seq.filter (fun (a: AST.Address) ->
                let a_path = a.A1Path()
                Some(a_path) = cr.Path
            ) refs
        | None,false -> raise (ParseException cr.Formula)
        | None,true -> Seq.empty    // just ignore parse exceptions for now

    // This is the one you want to call from user code
    let ParseFormula(str, path, wb, ws): AST.Expression =
        match run ExcelParser.Formula str with
        | Success(formula, _, _) ->
            ExcelParser.ExprAddrResolve formula path wb ws
            formula
        | Failure(errorMsg, _, _) -> raise (ParseException(str, errorMsg))