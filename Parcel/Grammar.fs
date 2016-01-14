module Grammar
    open FParsec
    open AST
    open System.Text.RegularExpressions

    (*
     * FPARSEC CHEAT SHEET
     * -------------------
     *
     * <|>      The infix "choice combinator" <|> applies the parser on the right side if the parser on the left side fails.
     * |>>      The infix "pipeline combinator" |>> applies the function on the right side to the result of the parser on the left side.
     * >>.      p1 >>. p2 parses p1 and p2 in sequence and returns the result of p2. 
     * .>>      p1 .>> p2 also parses p1 and p2 in sequence, but it returns the result of p1 instead of p2.
     * .>>.     An infix synonym for tuple2
     * >>=      The "bind combinator" takes a parser and a function producing a parser as arguments. p >>= f first applies
     *          the parser p to the input, then applies the function f to the result returned by p and finally applies
     *          the parser returned by f to the input.
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

    // a simple typedef
    type P<'t> = Parser<'t, Env>   
    
    // custom character classes
    let isWSChar(c: char) : bool =
        isDigit(c) || isLetter(c) || c = '-' || c = ' '

    // Special breakpoint-friendly parser
//    let BP (p: Parser<_,_>)(stream: CharStream<'b>) =
//        printfn "At index: %d, string remaining: %s" (stream.Index) (stream.PeekString 1000)
//        p stream // set a breakpoint here
//
    let (<!>) (p: Parser<_,_>) label : Parser<_,_> =
        fun stream ->
//            printfn "At index: %d, string remaining: %s" (stream.Index) (stream.PeekString 1000)
//            printfn "%A: Entering %s" stream.Position label
            let s = String.replicate (int (stream.Position.Column)) " "
            printfn "%d%sTrying %s(\"%s\")" (stream.Index) s label (stream.PeekString 1000)
            let reply = p stream
            printfn "%d%s%s(\"%s\") (%A)" (stream.Index) s label (stream.PeekString 1000) reply.Status
//            printfn "%A: Leaving %s (%A)" stream.Position label reply.Status
            reply

    // Grammar forward references
    let (ArgumentList: P<Expression list>, ArgumentListImpl) = createParserForwardedToRef()
    let (ExpressionSimple: P<Expression>, ExpressionSimpleImpl) = createParserForwardedToRef()
    let (ExpressionDecl: P<Expression>, ExpressionDeclImpl) = createParserForwardedToRef()
    let (RangeA1: P<Range>, RangeA1Impl) = createParserForwardedToRef()
    let (RangeR1C1: P<Range>, RangeR1C1Impl) = createParserForwardedToRef()

    // Addresses
    // We treat relative and absolute addresses the same-- they behave
    // exactly the same way unless you copy and paste them.
    let AddrR = pstring "R" >>. pint32
    let AddrC = pstring "C" >>. pint32
    let AddrIndirect = getUserState >>=
                        fun us ->
                            between
                                (pstring "INDIRECT(")
                                (pstring ")")
                                (manySatisfy ((<>) ')'))
                            |>> (fun expr -> IndirectAddress(expr, us) :> Address)
    let AddrR1C1 = getUserState >>=
                    fun (us: Env) ->
                        (attempt AddrIndirect)
                        <|> (pipe2
                                AddrR
                                AddrC
                                (fun row col ->
                                    Address.fromR1C1(row, col, us.WorksheetName, us.WorkbookName, us.Path)))
                    <!> "AddrR1C1"
    let AddrA = many1Satisfy isAsciiUpper
    let AddrAAbs = (pstring "$" <|> pstring "") >>. AddrA
    let Addr1 = pint32
    let Addr1Abs = (pstring "$" <|> pstring "") >>. Addr1
    let AddrA1 = getUserState >>=
                    fun (us: Env) ->
                        (attempt AddrIndirect)
                        <|> (pipe2
                            AddrAAbs
                            Addr1Abs
                            (fun col row ->
                                Address.fromA1(row, col, us.WorksheetName, us.WorkbookName, us.Path)))
                    <!> "AddrA1"
    let AnyAddr = ((attempt AddrIndirect)
                  <|> (attempt AddrR1C1)
                  <|> AddrA1)
                  <!> "AnyAddr"

    // Ranges
    let RA1TO = (* RTO *)
                (attempt (pstring ":" >>. AddrA1) <!> "RTO (first)")
                <|> ((pstring ":" >>. AddrA1 .>> pstring ",") <!> "RTO (second)")
    let RA1TC = (* RTC *)
                (attempt (pstring "," >>. AddrA1) <!> "RTC (first)")
                <|> ((pstring "," >>. AddrA1 .>> pstring ",") <!> "RTC (second)")
    let RA1_1 = (* AddrA1 RTC *) pipe2 AddrA1 RA1TC (fun a1 a2 -> Range(a1,a2)) <!> "AddrA1 RTC"
    let RA1_2 = (* AddrA1 RTO *) pipe2 AddrA1 RA1TO (fun a1 a2 -> Range(a1,a2)) <!> "AddrA1 RTO"
    let RA1_3 = (* AddrA1 RTO RTC *) pipe3 AddrA1 RA1TO RA1TC (fun a1 a2 a3 -> Range([(a1,a2);(a3,a3)])) <!> "AddrA1 RTO RTC"
    let RA1_4 = (* AddrA1 "," R *) pipe2 (AddrA1 .>> (pstring ",")) RangeA1 (fun a1 r1 -> Range(r1.Ranges() @ [(a1, a1)])) <!> """AddrA1 "," R"""
    let RA1_5 = (* AddrA1 RTO R *) pipe3 AddrA1 RA1TO RangeA1 (fun a1 a2 r2 -> Range ((a1,a2) :: r2.Ranges())) <!> "AddrA1 RTO R"
    do RangeA1Impl :=   (attempt RA1_4)
                        <|> (attempt RA1_5)
                        <|> (attempt RA1_3)
                        <|> (attempt RA1_1)
                        <|> (attempt RA1_2)
                        
    do RangeR1C1Impl := RangeA1

    let R = RangeA1 <!> "Range"
//    let R = (attempt RangeR1C1) <|> RangeA1

    // Worksheet Names
    let WorksheetNameQuoted =
        let NormalChar = satisfy ((<>) '\'')
        let EscapedChar = pstring "''" |>> (fun s -> ''')
        between (pstring "'") (pstring "'")
                (many1Chars (NormalChar <|> EscapedChar))
    let WorksheetNameUnquoted = (many1Satisfy (fun c -> isWSChar(c)))
    let WorksheetName = (WorksheetNameQuoted <|> WorksheetNameUnquoted)
    // Workbook Names (this may be too restrictive)
    let Path = many1Satisfy ((<>) '[')
    let WorkbookName = between
                        (pstring "[")
                        (pstring "]")
                        (many1Satisfy (fun c -> c <> '[' && c <> ']'))
    let Workbook = (
                        (Path |>> Some)
                        <|> ((pstring "") >>% None)
                   )
                   .>>. WorkbookName

    // References
    // References consist of the following parts:
    //   A workbook name prefix
    //   A worksheet name prefix
    //   A single-cell address ("Address") or multi-cell address ("Range")
    let RRWQuoted = (between (pstring "'") (pstring "'") (Workbook .>>. WorksheetNameUnquoted))
    let RangeReferenceWorkbook = getUserState >>=
                                    fun us ->
                                        (pipe2
                                            (RRWQuoted .>> pstring "!")
                                            R
                                            (fun ((wbpath, wbname), wsname) rng ->
                                                match wbpath with
                                                | Some(pth) -> ReferenceRange(Env(pth, wbname, wsname), rng) :> Reference
                                                | None      -> ReferenceRange(Env(us.Path, wbname, wsname), rng) :> Reference
                                            )
                                        )
                                        <!> "RangeReferenceWorkbook"
    let RangeReferenceWorksheet = getUserState >>=
                                    fun us ->
                                        pipe2
                                            (WorksheetName .>> pstring "!")
                                            R
                                            (fun wsname rng -> ReferenceRange(Env(us.Path, us.WorkbookName, wsname), rng) :> Reference)
                                    <!> "RangeReferenceWorksheet"
    let RangeReferenceNoWorksheet = getUserState >>=
                                        fun us ->
                                            R
                                            |>> (fun rng -> ReferenceRange(us, rng) :> Reference)
                                    <!> "RangeReferenceNOWorksheet"
    let RangeReference = (attempt RangeReferenceWorkbook)
                         <|> (attempt RangeReferenceWorksheet)
                         <|> RangeReferenceNoWorksheet
                         <!> "RangeReference"

    let ARWQuoted = (between
                        (pstring "'")
                        (pstring "'")
                        (Workbook .>>. WorksheetNameUnquoted))
                    <!> "ARWQuoted"
    let AddressReferenceWorkbook = getUserState >>=
                                    fun us ->
                                        (pipe2
                                            (ARWQuoted .>> pstring "!")
                                            AnyAddr
                                            (fun ((wbpath, wbname), wsname) addr ->
                                                match wbpath with
                                                | Some(pth) -> ReferenceAddress(Env(pth, wbname, wsname), addr) :> Reference
                                                | None      -> ReferenceAddress(Env(us.Path, wbname, wsname), addr) :> Reference
                                            )
                                        )
                                    <!> "AddressReferenceWorkbook"
    let AddressReferenceWorksheet = getUserState >>=
                                        fun us ->
                                            pipe2
                                                (WorksheetName .>> pstring "!")
                                                AnyAddr
                                                (fun wsname addr -> ReferenceAddress(Env(us.Path, us.WorkbookName, wsname), addr) :> Reference)
                                    <!> "AddressReferenceWorksheet"
    let AddressReferenceNoWorksheet = getUserState >>=
                                        fun us ->
                                            AnyAddr
                                            |>> (fun addr -> ReferenceAddress(us, addr) :> Reference)
                                      <!> "AddressReferenceNOWorksheet"
    let AddressReference = (attempt AddressReferenceWorkbook)
                           <|> (attempt AddressReferenceWorksheet)
                           <|> AddressReferenceNoWorksheet
                           <!> "AddressReference"

    let NamedReferenceFirstChar = satisfy (fun c -> c = '_' || isLetter(c))
    let NamedReferenceLastChars = manySatisfy (fun c -> c = '_' || isLetter(c) || isDigit(c))
    let NamedReference = getUserState >>=
                            fun us ->
                                pipe2
                                    NamedReferenceFirstChar
                                    NamedReferenceLastChars
                                    (fun c s -> ReferenceNamed(us, c.ToString() + s) :> Reference)
                         <!> "NamedReference"
    let StringReference = getUserState >>=
                            fun us ->
                                between
                                    (pstring "\"")
                                    (pstring "\"")
                                    (manySatisfy ((<>) '"'))
                                |>> (fun s -> ReferenceString(us, s) :> Reference)
                          <!> "StringReference"
    let ConstantReference = getUserState >>= 
                                fun us ->
                                    (attempt
                                        (pfloat .>> pstring "%"
                                        |>> fun r -> ReferenceConstant(us, r / 100.0) :> Reference))
                                    <|> (pfloat |>> (fun r -> ReferenceConstant(us, r) :> Reference))
                            <!> "ConstantReference"
    let Reference = (attempt RangeReference) <|> (attempt AddressReference) <|> (attempt ConstantReference) <|> (attempt StringReference) <|> NamedReference
                    <!> "Reference"
    // Functions
    let FunctionName = (pstring "INDIRECT" >>. pzero) <|> many1Satisfy (fun c -> isLetter(c))
                       <!> "FunctionName"
    let Function = getUserState >>=
                    fun us ->
                        pipe2
                            (FunctionName .>> pstring "(")
                            (ArgumentList .>> pstring ")")
                            (fun fname arglist -> ReferenceFunction(us, fname, arglist) :> Reference)
                   <!> "Function"
    do ArgumentListImpl := sepBy ExpressionDecl (spaces >>. pstring "," .>> spaces) <!> "ArgumentList"

    // Binary arithmetic operators
    let BinOpChar = spaces >>. satisfy (fun c -> c = '+' || c = '-' || c = '/' || c = '*' || c = '<' || c = '>' || c = '=' || c = '^' || c = '&') .>> spaces
    let BinOp2Char = spaces >>. ((attempt (regex "<=")) <|> (attempt (regex ">=")) <|> regex "<>") .>> spaces
    let BinOpLong: P<string*Expression> = pipe2 BinOp2Char ExpressionDecl (fun op rhs -> (op, rhs))
    let BinOpShort: P<string*Expression> = pipe2 BinOpChar ExpressionDecl (fun op rhs -> (op.ToString(), rhs))
    let BinOp: P<string*Expression> = ((attempt BinOpLong) <|> BinOpShort) <!> "BinaryOp"

    // Unary operators
    let UnaryOpChar = (spaces >>. satisfy (fun c -> c = '+' || c = '-') .>> spaces) <!> "UnaryOp"

    // Expressions
    let ParensExpr: P<Expression> = ((between (pstring "(") (pstring ")") ExpressionDecl) |>> ParensExpr) <!> "ParensExpr"
    let ExpressionAtom: P<Expression> = (((attempt Function) <|> Reference) |>> ReferenceExpr) <!> "ExpressionAtom"
    do ExpressionSimpleImpl := (ExpressionAtom <|> ParensExpr) <!> "ExpressionSimple"
    let UnaryOpExpr: P<Expression> = pipe2 UnaryOpChar ExpressionDecl (fun op rhs -> UnaryOpExpr(op, rhs)) <!> "UnaryOpExpr"
    let BinOpExpr: P<Expression> = pipe2 ExpressionSimple BinOp (fun lhs (op, rhs) -> BinOpExpr(op, lhs, rhs)) <!> "BinaryOpExpr"
    do ExpressionDeclImpl := ((attempt UnaryOpExpr) <|> (attempt BinOpExpr) <|> (attempt ExpressionSimple)) <!> "Expression"

    // Formulas
    let Formula = (pstring "=" .>> spaces >>. ExpressionDecl .>> eof) <!> "Formula"