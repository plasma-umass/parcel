module ExcelParserUtility
    open FParsec
    open SpreadsheetAST

    // using C#-style exceptions so that the can be handled by C# code
    type ParseException(formula: string, reason: string) =
        inherit System.Exception(formula)
        new(formula: string) = ParseException(formula, "")

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