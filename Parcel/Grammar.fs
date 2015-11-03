module Grammar
    open FParsec
    open AST
    open System.Text.RegularExpressions

    (*
     * FPARSEC CHEAT SHEET
     * -------------------
     *
     * <|>      The infix choice combinator <|> applies the parser on the right side if the parser on the left side fails.
     * |>>      The infix pipeline combinator |>> applies the function on the right side to the result of the parser on the left side.
     * >>.      p1 >>. p2 parses p1 and p2 in sequence and returns the result of p2. 
     * .>>      p1 .>> p2 also parses p1 and p2 in sequence, but it returns the result of p1 instead of p2.
     * .>>.     An infix synonym for tuple2
     * tuple2   tuple2 p1 p2 returns a tuple consisting of the result of p1 (first) and the result of p2 (second)
     * pipe2    pipe2 p1 p2 f sequentially applies the two parsers p1 and p2 and then returns the result of the function
     *          application f x1 x2, where x1 and x2 are the results returned by p1 and p2.
     * sepBy    sepBy p1 p2 takes an "element" parser (p1) and a "separator" parser (p2) as the arguments and returns
     *          a parser for a list of elements separated by separators. 
     * createParserForwardedToRef
     *          let p, pRef = createParserForwardedToRef() creates a parser p that forwards all calls to the parser in
     *          the reference cell pRef. Initially, pRef holds a reference to a dummy parser that raises an exception on
     *          any invocation.
     * do ... :=
     *          do pRef :=, instead of let p =, lets us overwrite the implementation of parser p using the reference to p, pRef.
     *)

    // The Defaults object is threaded through all combinators so that
    // parsers always have the path, workbook name, and worksheet name
    // so that they can be inserted into the AST nodes as needed.
    type Defaults(path: string, wbname: string, wsname: string) = class end
    // a simple typedef
    type P<'t> = Parser<'t, Defaults>

    
    // custom character classes
    let isWSChar(c: char) : bool =
        isDigit(c) || isLetter(c) || c = '-' || c = ' '

    // Special breakpoint-friendly parser
    let BP (p: Parser<_,_>)(stream: CharStream<'b>) =
        printfn "At index: %d, string remaining: %s" (stream.Index) (stream.PeekString 1000)
        p stream // set a breakpoint here

    let (<!>) (p: Parser<_,_>) label : Parser<_,_> =
        fun stream ->
            printfn "%A: Entering %s" stream.Position label
            let reply = p stream
            printfn "%A: Leaving %s (%A)" stream.Position label reply.Status
            reply

    // Grammar forward references
    let (ArgumentList: P<Expression list>, ArgumentListImpl) = createParserForwardedToRef()
    let (ExpressionSimple: P<Expression>, ExpressionSimpleImpl) = createParserForwardedToRef()
    let (ExpressionDecl: P<Expression>, ExpressionDeclImpl) = createParserForwardedToRef()

    // Addresses
    // We treat relative and absolute addresses the same-- they behave
    // exactly the same way unless you copy and paste them.
    let AddrR = pstring "R" >>. pint32
    let AddrC = pstring "C" >>. pint32
    let AddrIndirect = between (pstring "INDIRECT(") (pstring ")") (manySatisfy ((<>) ')')) |>> (fun exprstr -> IndirectAddress(exprstr) :> Address)
    let AddrR1C1 = (attempt AddrIndirect) <|> pipe2 AddrR AddrC (fun r c -> Address.NewFromR1C1(r,c,None,None,None))
    let AddrA = many1Satisfy isAsciiUpper
    let AddrAAbs = (pstring "$" <|> pstring "") >>. AddrA
    let Addr1 = pint32
    let Addr1Abs = (pstring "$" <|> pstring "") >>. Addr1
    let AddrA1 = (attempt AddrIndirect) <|> (pipe2 AddrAAbs Addr1Abs (fun col row -> Address.NewFromA1(row,col,None,None,None))) <!> "AddrA1"
    let AnyAddr = ((attempt AddrIndirect) <|> (attempt AddrR1C1) <|> AddrA1) <!> "AnyAddr"

    // Ranges
    let MoreAddrR1C1 = pstring ":" >>. AddrR1C1
    let RangeR1C1 = pipe2 AddrR1C1 MoreAddrR1C1 (fun r1 r2 -> Range(r1, r2))
    let MoreAddrA1 = pstring ":" >>. AddrA1
    let RangeA1 = pipe2 AddrA1 MoreAddrA1 (fun r1 r2 -> Range(r1, r2))
    let RangeAny = ((attempt RangeR1C1) <|> RangeA1) <!> "RangeAny"

    // Worksheet Names

    let WorksheetNameQuoted =
        let NormalChar = satisfy ((<>) '\'')
        let EscapedChar = pstring "''" |>> (fun s -> ''')
        between (pstring "'") (pstring "'")
                (many1Chars (NormalChar <|> EscapedChar))
    let WorksheetNameUnquoted = (many1Satisfy (fun c -> isWSChar(c))) <!> "WorksheetNameUnquoted"
    let WorksheetName = (WorksheetNameQuoted <|> WorksheetNameUnquoted) <!> "WorksheetName"

    // Workbook Names (this may be too restrictive)
    let Path = many1Satisfy ((<>) '[') <!> "Path"
    let WorkbookName = between (pstring "[") (pstring "]") (many1Satisfy (fun c -> c <> '[' && c <> ']')) <!> "WorkbookName"
    let Workbook = ((Path |>> Some) <|> ((pstring "") >>% None)) .>>. WorkbookName <!> "Workbook"

    // References
    // References consist of the following parts:
    //   An optional workbook name prefix
    //   An optional worksheet name prefix
    //   A single-cell ("Address") or multi-cell address ("Range")
    let RRWQuoted = (between (pstring "'") (pstring "'") (Workbook .>>. WorksheetNameUnquoted)) <!> "RRWQuoted"
    let RangeReferenceWorkbook = (pipe2 (RRWQuoted .>> pstring "!") RangeAny (fun ((wbpath, wbname), wsname) rng -> ReferenceRange(wbpath, Some(wbname), Some(wsname), rng) :> Reference)) <!> "RangeReferenceWorkbook"
    let RangeReferenceWorksheet = pipe2 (WorksheetName .>> pstring "!") RangeAny (fun wsname rng -> ReferenceRange(None, None, Some(wsname), rng) :> Reference)
    let RangeReferenceNoWorksheet = RangeAny |>> (fun rng -> ReferenceRange(None, None, None, rng) :> Reference)
    let RangeReference = (attempt RangeReferenceWorkbook) <|> (attempt RangeReferenceWorksheet) <|> RangeReferenceNoWorksheet

    let ARWQuoted = (between (pstring "'") (pstring "'") (Workbook .>>. WorksheetNameUnquoted)) <!> "ARWQuoted"
    let AddressReferenceWorkbook = (pipe2 (ARWQuoted .>> pstring "!") AnyAddr (fun ((wbpath, wbname), wsname) addr ->  ReferenceAddress(wbpath, Some(wbname), Some(wsname), addr) :> Reference)) <!> "AddressReferenceWorkbook"
    let AddressReferenceWorksheet = pipe2 (WorksheetName .>> pstring "!") AnyAddr (fun wsname addr -> ReferenceAddress(None, None, Some(wsname), addr) :> Reference)
    let AddressReferenceNoWorksheet = AnyAddr |>> (fun addr -> ReferenceAddress(None, None, None, addr) :> Reference)
    let AddressReference = (attempt AddressReferenceWorkbook) <|> (attempt AddressReferenceWorksheet) <|> AddressReferenceNoWorksheet

    let NamedReferenceFirstChar = satisfy (fun c -> c = '_' || isLetter(c))
    let NamedReferenceLastChars = manySatisfy (fun c -> c = '_' || isLetter(c) || isDigit(c))
    let NamedReference = pipe2 NamedReferenceFirstChar NamedReferenceLastChars (fun c s -> ReferenceNamed(None, c.ToString() + s) :> Reference)

    let StringReference = between (pstring "\"") (pstring "\"") (manySatisfy ((<>) '"')) |>> (fun s -> ReferenceString(None, s) :> Reference)

    let ConstantReference = (attempt (pfloat .>> pstring "%" |>> (fun r -> ReferenceConstant(None, r / 100.0) :> Reference)))
                            <|> (pfloat |>> (fun r -> ReferenceConstant(None, r) :> Reference))

    let Reference = (attempt RangeReference) <|> (attempt AddressReference) <|> (attempt ConstantReference) <|> (attempt StringReference) <|> NamedReference

    // Functions
    let FunctionName = (pstring "INDIRECT" >>. pzero) <|> many1Satisfy (fun c -> isLetter(c))
    let Function = pipe2 (FunctionName .>> pstring "(") (ArgumentList .>> pstring ")") (fun fname arglist -> ReferenceFunction(None, fname, arglist) :> Reference)
    do ArgumentListImpl := sepBy ExpressionDecl (spaces >>. pstring "," .>> spaces)

    // Binary arithmetic operators
    let BinOpChar = spaces >>. satisfy (fun c -> c = '+' || c = '-' || c = '/' || c = '*' || c = '<' || c = '>' || c = '=' || c = '^' || c = '&') .>> spaces
    let BinOp2Char = spaces >>. ((attempt (regex "<=")) <|> (attempt (regex ">=")) <|> regex "<>") .>> spaces
    let BinOpLong: P<string*Expression> = pipe2 BinOp2Char ExpressionDecl (fun op rhs -> (op, rhs))
    let BinOpShort: P<string*Expression> = pipe2 BinOpChar ExpressionDecl (fun op rhs -> (op.ToString(), rhs))
    let BinOp: P<string*Expression> = (attempt BinOpLong) <|> BinOpShort

    // Unary operators
    let UnaryOpChar = spaces >>. satisfy (fun c -> c = '+' || c = '-') .>> spaces

    // Expressions
    let ParensExpr: P<Expression> = (between (pstring "(") (pstring ")") ExpressionDecl) |>> ParensExpr
    let ExpressionAtom: P<Expression> = ((attempt Function) <|> Reference) |>> ReferenceExpr
    do ExpressionSimpleImpl := ExpressionAtom <|> ParensExpr
    let UnaryOpExpr: P<Expression> = pipe2 UnaryOpChar ExpressionDecl (fun op rhs -> UnaryOpExpr(op, rhs))
    let BinOpExpr: P<Expression> = pipe2 ExpressionSimple BinOp (fun lhs (op, rhs) -> BinOpExpr(op, lhs, rhs))
    do ExpressionDeclImpl := (attempt UnaryOpExpr) <|> (attempt BinOpExpr) <|> (attempt ExpressionSimple)

    // Formulas
    let Formula = pstring "=" .>> spaces >>. ExpressionDecl .>> eof