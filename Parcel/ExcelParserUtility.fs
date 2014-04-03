module ExcelParserUtility
    open FParsec
    open SpreadsheetAST

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

    let rec GetRangeReferenceRanges(ref: ReferenceRange) : Range list = [ref.Range]

    and GetFunctionRanges(ref: ReferenceFunction) : Range list =
        if PuntedFunction(ref.FunctionName) then
            []
        else
            List.map (fun arg -> GetExprRanges(arg)) ref.ArgumentList |> List.concat
        
    and GetExprRanges(expr: Expression) : Range list =
        match expr with
        | ReferenceExpr(r) -> GetRanges(r)
        | BinOpExpr(op, e1, e2) -> GetExprRanges(e1) @ GetExprRanges(e2)
        | UnaryOpExpr(op, e) -> GetExprRanges(e)
        | ParensExpr(e) -> GetExprRanges(e)

    and GetRanges(ref: Reference) : Range list =
        match ref with
        | :? ReferenceRange as r -> GetRangeReferenceRanges(r)
        | :? ReferenceAddress -> []
        | :? ReferenceNamed -> []   // TODO: symbol table lookup
        | :? ReferenceFunction as r -> GetFunctionRanges(r)
        | :? ReferenceConstant -> []
        | :? ReferenceString -> []
        | _ -> failwith "Unknown reference type."

    let GetReferencesFromFormula(formula: string, wbname: string, wsname: string) : seq<Range> =
        let wbpath = System.IO.Path.GetDirectoryName(wbname)
        try
            match ExcelParser.ParseFormula(formula, wbpath, wbname, wsname) with
            | Some(tree) ->
                let refs = GetExprRanges(tree)
                List.map (fun (r: Range) ->
                            // temporarily bail if r refers to object in a different workbook
                            // TODO: we should open the other workbook and continue the analysis
                            let paths = Seq.filter (fun pth -> pth = wbpath) (r.GetPathNames())
                            if Seq.length(paths) = 0 then
                                None
                            else
                                Some(r)
                         ) refs |> List.choose id |> Seq.ofList
            | None -> raise (ParseException(formula))
        with
        | :? IndirectAddressingNotSupportedException -> seq[]

    // single-cell variants:

    let rec GetSCExprRanges(expr: Expression) : Address list =
        match expr with
        | ReferenceExpr(r) -> GetSCRanges(r)
        | BinOpExpr(op, e1, e2) -> GetSCExprRanges(e1) @ GetSCExprRanges(e2)
        | UnaryOpExpr(op, e) -> GetSCExprRanges(e)
        | ParensExpr(e) -> GetSCExprRanges(e)

    and GetSCRanges(ref: Reference) : Address list =
        match ref with
        | :? ReferenceRange -> []
        | :? ReferenceAddress as r -> GetSCAddressReferenceRanges(r)
        | :? ReferenceNamed -> []   // TODO: symbol table lookup
        | :? ReferenceFunction as r -> GetSCFunctionRanges(r)
        | :? ReferenceConstant -> []
        | :? ReferenceString -> []
        | _ -> failwith "Unknown reference type."

    and GetSCAddressReferenceRanges(ref: ReferenceAddress) : Address list = [ref.Address]

    and GetSCFunctionRanges(ref: ReferenceFunction) : Address list =
        if PuntedFunction(ref.FunctionName) then
            []
        else
            List.map (fun arg -> GetSCExprRanges(arg)) ref.ArgumentList |> List.concat

    let rec GetFormulaNamesFromExpr(ast: Expression): string list =
        match ast with
        | ReferenceExpr(r) -> GetFormulaNamesFromReference(r)
        | BinOpExpr(op, e1, e2) -> GetFormulaNamesFromExpr(e1) @ GetFormulaNamesFromExpr(e2)
        | UnaryOpExpr(op, e) -> GetFormulaNamesFromExpr(e)
        | ParensExpr(e) -> GetFormulaNamesFromExpr(e)

    and GetFormulaNamesFromReference(ref: Reference): string list =
        match ref with
        | :? ReferenceFunction as r -> [r.FunctionName]
        | _ -> []

    let GetSCFormulaNames(formula: string, path: string, wsname: string, wbname: string) =
        let path' = System.IO.Path.GetDirectoryName(wbname)
        match ExcelParser.ParseFormula(formula, path', wbname, wsname) with
        | Some(ast) -> GetFormulaNamesFromExpr(ast) |> Seq.ofList
        | None -> raise (ParseException formula)

    let GetSingleCellReferencesFromFormula(formula: string, wbname: string, wsname: string) : seq<Address> =
        let path = System.IO.Path.GetDirectoryName(wbname)
        match ExcelParser.ParseFormula(formula, path, wbname, wsname) with
        | Some(tree) ->
            // temporarily bail if a refers to object in a different workbook
            // TODO: we should open the other workbook and continue the analysis
            let refs = GetSCExprRanges(tree) |> Seq.ofList
            Seq.filter (fun (a: Address) ->
                let a_path = a.A1Path()
                a_path = path
            ) refs
        | None -> raise (ParseException formula)

    // This is the one you want to call from user code
    let ParseFormula(fstr, path, wb, ws): Expression =
        match run ExcelParser.Formula fstr with
        | Success(formula, _, _) ->
            ExcelParser.ExprAddrResolve formula path wb ws
            formula
        | Failure(errorMsg, _, _) -> raise (ParseException(fstr, errorMsg))

    // You can also call this one from user code if you
    // already have an Address object
    let ParseFormulaWithAddress(fstr: string, addr: Address): Expression =
        match run ExcelParser.Formula fstr with
        | Success(formula, _, _) ->
            ExcelParser.ExprAddrResolve formula (addr.A1Path()) (addr.A1Workbook()) (addr.A1Worksheet())
            formula
        | Failure(errorMsg, _, _) -> raise (ParseException(fstr, errorMsg))