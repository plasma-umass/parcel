module Parcel
    open FParsec

    (*
     * TYPES
     *)
    type ParseException(formula: string, reason: string) =
        inherit System.Exception(formula)
        new(formula: string) = ParseException(formula, "")

    (*
     * PRIVATE IMPLEMENTATIONS
     *)

    // wrapper for success/failure
    let private test p str =
        match runParserOnString p (AST.Defaults("", "", "")) "" str with
        | Success(result, _, _)     -> printfn "Success: %A" result
        | Failure(errorMsg, _, _)   -> printfn "Failure: %s" errorMsg

    let private getAddress(formula: string)(path: string)(wbname: string)(wsname: string): AST.Address =
        match runParserOnString (Grammar.AddrR1C1 .>> eof) (AST.Defaults(path, wbname, wsname)) "" formula with
        | Success(addr, _, _)       -> addr
        | Failure(errorMsg, _, _)   -> failwith errorMsg

    let private getRange(formula: string)(path: string)(wbname: string)(wsname: string): AST.Range option =
        match runParserOnString (Grammar.RangeR1C1 .>> eof) (AST.Defaults(path, wbname, wsname)) "" formula with
        | Success(range, _, _)      -> Some(range)
        | Failure(errorMsg, _, _)   -> None

    let private getReference(formula: string)(path: string)(wbname: string)(wsname: string): AST.Reference option =
        match runParserOnString (Grammar.Reference .>> eof) (AST.Defaults(path, wbname, wsname)) "" formula with
        | Success(reference, _, _)  -> Some(reference)
        | Failure(errorMsg, _, _)   -> None

    let private isNumeric(str: string): bool =
        match run (pfloat .>> eof) str with
        | Success(number, _, _)     -> true
        | Failure(errorMsg, _, _)   -> false

    (*
     * PUBLIC API
     *)
    let parseFormula(formula: string)(path: string)(wbname: string)(wsname: string): AST.Expression option =
        match runParserOnString Grammar.Formula (AST.Defaults(path, wbname, wsname)) "" formula with
        | Success(formula, _, _) -> Some(formula)
        | Failure(errorMsg, _, _) -> None

    // The parser REPL calls this; note that the
    // Formula parser looks for EOF
    let consoleParser(s: string) = test Grammar.Formula s

    // Call this for simple address parsing
    let simpleReferenceParser(s: string) : AST.Reference =
        match runParserOnString Grammar.Reference (AST.Defaults("", "", "")) "" s with
        | Success(result, _, _) -> result
        | Failure(errorMsg, _, _) -> failwith ("String \"" + s + "\" does not appear to be a Reference:\n" + errorMsg)

    let rangeReferencesFromFormula(formula: string, path: string, workbook: string, worksheet: string, ignore_parse_errors: bool) : seq<AST.Range> =
        try
            match (parseFormula formula path workbook worksheet),ignore_parse_errors with
            | Some(tree),_ ->
                let refs = RangeVisitor.rangesFromExpr(tree)
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

    let addrReferencesFromFormula(formula: string, path: string, wb: string, ws: string, ignore_parse_errors: bool) : seq<AST.Address> =
        match (parseFormula formula path wb ws),ignore_parse_errors with
        | Some(ast),_ ->
            // temporarily bail if a refers to object in a different workbook
            // TODO: we should open the other workbook and continue the analysis
            let refs = CellVisitor.addrsFromExpr(ast) |> Seq.ofList
            Seq.filter (fun (a: AST.Address) ->a.A1Path() = path) refs
        | None,false -> raise (ParseException formula)
        | None,true -> Seq.empty    // just ignore parse exceptions for now

    let rec formulaNamesFromExpr(ast: AST.Expression): string list =
        match ast with
        | AST.ReferenceExpr(r) -> formulaNamesFromRef(r)
        | AST.BinOpExpr(op, e1, e2) -> formulaNamesFromExpr(e1) @ formulaNamesFromExpr(e2)
        | AST.UnaryOpExpr(op, e) -> formulaNamesFromExpr(e)
        | AST.ParensExpr(e) -> formulaNamesFromExpr(e)

    and formulaNamesFromRef(ref: AST.Reference): string list =
        match ref with
        | :? AST.ReferenceFunction as r -> [r.FunctionName]
        | _ -> []

//    let formulaNamesFromFormula(formula: string)(path: string)(wb: string)(ws: string)(ignore_parse_errors: bool) =
//        let abspath = Some(System.IO.Path.GetDirectoryName(wb))
//        match (parseFormula formula abspath wb ws),ignore_parse_errors with
//        | Some(ast),_ -> formulaNamesFromExpr(ast) |> Seq.ofList
//        | None,false -> raise (ParseException formula)
//        | None,true -> Seq.empty    // just ignore parse exceptions for now

