﻿module Parcel
    open FParsec

    (*
     * PRIVATE IMPLEMENTATIONS
     *)

    // wrapper for success/failure
    let private test p str =
        match runParserOnString p (AST.Env("", "", "")) "" str with
        | Success(result, _, _)     -> printfn "Success: %A" result
        | Failure(errorMsg, _, _)   -> printfn "Failure: %s" errorMsg

    let private getAddress(formula: string)(path: string)(wbname: string)(wsname: string): AST.Address =
        match runParserOnString (Grammar.AddrR1C1 .>> eof) (AST.Env(path, wbname, wsname)) "" formula with
        | Success(addr, _, _)       -> addr
        | Failure(errorMsg, _, _)   -> failwith errorMsg

    let private getRange(formula: string)(path: string)(wbname: string)(wsname: string): AST.Range option =
        match runParserOnString (Grammar.RangeR1C1 .>> eof) (AST.Env(path, wbname, wsname)) "" formula with
        | Success(range, _, _)      -> Some(range)
        | Failure(errorMsg, _, _)   -> None

    let private getReference(formula: string)(path: string)(wbname: string)(wsname: string): AST.Reference option =
        match runParserOnString (Grammar.Reference Grammar.RangeAny .>> eof) (AST.Env(path, wbname, wsname)) "" formula with
        | Success(reference, _, _)  -> Some(reference)
        | Failure(errorMsg, _, _)   -> None

    (*
     * PUBLIC API
     *)
    let isNumeric(str: string): bool =
        match run (pfloat .>> eof) str with
        | Success(number, _, _)     -> true
        | Failure(errorMsg, _, _)   -> false

    let parseFormula(formula: string)(path: string)(wbname: string)(wsname: string): AST.Expression option =
        try
            match runParserOnString Grammar.Formula (AST.Env(path, wbname, wsname)) "" formula with
            | Success(formula, _, _) -> Some(formula)
            | Failure(errorMsg, _, _) -> None
        with
        | :? AST.IndirectAddressingNotSupportedException as e -> raise e    // rethrow
        | :? AST.ParseException as e -> raise e                             // rethrow
        | e ->                                                              // wrap other exception
            let rec f = (fun (e:exn) ->
                            e.Message +
                            "\n\n" +
                            if (e.InnerException <> null) then f e.InnerException else ""
                        )
            raise (AST.ParseException (f e))
            

    let parseFormulaAtAddress(fAddr: AST.Address)(formula: string): AST.Expression =
        match parseFormula formula fAddr.Path fAddr.WorkbookName fAddr.WorksheetName with
        | Some ast -> ast
        | None -> failwith ("Parse error on formula: " + formula)

    // The parser REPL calls this; note that the
    // Formula parser looks for EOF
    let consoleParser(s: string) = test Grammar.Formula s

    // Call this for simple address parsing
    let simpleReferenceParser(s: string, e: AST.Env) : AST.Reference =
        match runParserOnString (Grammar.Reference Grammar.RangeAny) e "" s with
        | Success(result, _, _) -> result
        | Failure(errorMsg, _, _) -> failwith ("String \"" + s + "\" does not appear to be a Reference:\n" + errorMsg)

    let rangeReferencesFromExpr(tree: AST.Expression) : AST.Range[] =
        RangeVisitor.rangesFromExpr tree |> Seq.toArray

    let rangeReferencesFromFormula(formula: string, path: string, workbook: string, worksheet: string, ignore_parse_errors: bool) : AST.Range[] =
        try
            match (parseFormula formula path workbook worksheet),ignore_parse_errors with
            | Some(tree),_ -> rangeReferencesFromExpr tree
            | None,false -> raise (AST.ParseException(formula))
            | None,true -> [||]    // just ignore parse exceptions for now
        with
        // right now, we recognize indirect addresses but do not correctly dereference them,
        // which requires a program interpreter and input spreadsheet :(
        | :? AST.IndirectAddressingNotSupportedException -> [||]

    let addrReferencesFromExpr(tree: AST.Expression) : AST.Address[] =
        CellVisitor.addrsFromExpr tree |> Seq.toArray

    let addrReferencesFromFormula(formula: string, path: string, wb: string, ws: string, ignore_parse_errors: bool) : AST.Address[] =
        match (parseFormula formula path wb ws),ignore_parse_errors with
        | Some(tree),_ -> addrReferencesFromExpr tree
        | None,false -> raise (AST.ParseException formula)
        | None,true -> [||]    // just ignore parse exceptions for now

    let constantsFromExpr(e: AST.Expression) : AST.ReferenceConstant[] =
        CellVisitor.constantsFromExpr e |> Seq.toArray

    let constantsFromFormula(formula: string, path: string, wb: string, ws: string, ignore_parse_errors: bool) : AST.ReferenceConstant[] =
        match (parseFormula formula path wb ws),ignore_parse_errors with
        | Some(ast),_ -> constantsFromExpr ast
        | None,false -> raise (AST.ParseException formula)
        | None,true -> [||]    // just ignore parse exceptions for now

    let rec operatorNamesFromExpr(ast: AST.Expression): string list =
        match ast with
        | AST.ReferenceExpr(r) -> operatorNamesFromRef(r)
        | AST.BinOpExpr(op, e1, e2) -> op :: operatorNamesFromExpr(e1) @ operatorNamesFromExpr(e2)
        | AST.UnaryOpExpr(op, e) -> op.ToString() :: operatorNamesFromExpr(e)
        | AST.ParensExpr(e) -> operatorNamesFromExpr(e)

    and operatorNamesFromRef(ref: AST.Reference): string list =
        match ref with
        | :? AST.ReferenceFunction as r -> [r.FunctionName]
        | _ -> []


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

    let formulaNamesFromFormula(formula: string, path: string, wb: string, ws: string, ignore_parse_errors: bool) : seq<string> =
        let abspath = System.IO.Path.GetDirectoryName(wb)
        match (parseFormula formula abspath wb ws),ignore_parse_errors with
        | Some(ast),_ -> formulaNamesFromExpr(ast) |> Seq.ofList
        | None,false -> raise (AST.ParseException formula)
        | None,true -> Seq.empty    // just ignore parse exceptions for now

