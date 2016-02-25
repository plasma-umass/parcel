module Grammar
    open FParsec
    open AST
    open System.Text.RegularExpressions

    // This is a version of the forwarding parser that
    // also takes a generic argument.
    let createParserForwardedToRefWithArguments() =
        let dummyParser = fun x y -> failwith "a parser created with createParserForwardedToRefWithArguments was not initialized"
        let r = ref dummyParser
        (fun arg stream -> !r arg stream), r : ('t -> Parser<'a, 'u>) * ('t -> Parser<'a, 'u>) ref

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
    let (ArgumentList: P<Range> -> P<Expression list>, ArgumentListImpl) = createParserForwardedToRefWithArguments()
    let (ExpressionDecl: P<Range> -> P<Expression>, ExpressionDeclImpl) = createParserForwardedToRefWithArguments()
    let (RangeA1Union: P<Range>, RangeA1UnionImpl) = createParserForwardedToRef()
    let (RangeA1NoUnion: P<Range>, RangeA1NoUnionImpl) = createParserForwardedToRef()
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
    let AddrAAbs = (attempt (pstring "$") <|> pstring "") >>. AddrA
    let Addr1 = pint32
    let Addr1Abs = (attempt (pstring "$") <|> pstring "") >>. Addr1
    let AddrA1 = getUserState >>=
                    fun (us: Env) ->
                        (attempt AddrIndirect)
                        <|> (pipe2
                            AddrAAbs
                            Addr1Abs
                            (fun col row ->
                                Address.fromA1(row, col, us.WorksheetName, us.WorkbookName, us.Path)))
                    
    let AnyAddr = ((attempt AddrIndirect)
                  <|> (attempt AddrR1C1)
                  <|> AddrA1)
                  <!> "AnyAddr"

    // Ranges
    let RA1TO_1 = pstring ":" >>. AddrA1 .>> pstring "," <!> "RTO1"
    let RA1TO_2 = pstring ":" >>. AddrA1 <!> "RTO2"
    let RTO = (attempt RA1TO_1) <|> RA1TO_2

    let RA1TC_1 = pstring "," >>. AddrA1 .>> pstring "," <!> "RTC1"
    let RA1TC_2 = pstring "," >>. AddrA1 <!> "RTC2"
    let RTC = (attempt RA1TC_1) <|> RA1TC_2

    let RA1_1 = (* AddrA1 RTC (1)*) pipe2 AddrA1 RTC (fun a1 a2 -> Range([(a2,a2); (a1,a1)])) <!> "AddrA1 RTC ==(1)=="
    let RA1_2 = (* AddrA1 RTO (2)*) pipe2 AddrA1 RTO (fun a1 a2 -> Range(a1,a2)) <!> "AddrA1 RTO ==(2)=="

    let rat_fun = fun a1 a2 a3 -> Range([(a1,a2);(a3,a3)])
    let RA1_3a = pipe3 AddrA1 RA1TO_1 RA1TC_1 rat_fun <!> "AddrA1 RTO1 RTC1 ===(3a)==="
    let RA1_3b = pipe3 AddrA1 RA1TO_1 RA1TC_2 rat_fun <!> "AddrA1 RTO1 RTC2 ===(3b)==="
    let RA1_3c = pipe3 AddrA1 RA1TO_2 RA1TC_1 rat_fun <!> "AddrA1 RTO2 RTC1 ===(3c)==="
    let RA1_3d = pipe3 AddrA1 RA1TO_2 RA1TC_2 rat_fun <!> "AddrA1 RTO2 RTC2 ===(3d)==="
    let RA1_3 = (* AddrA1 RTO RTC (3)*) 
                (attempt RA1_3a) <|> (attempt RA1_3b) <|> (attempt RA1_3c) <|> RA1_3d <!> "AddrA1 RTO RTC ==(3)=="

    let RA1_4_UNION = (* AddrA1 "," R (4)*) pipe2 (AddrA1 .>> (pstring ",")) RangeA1Union (fun a1 r1 -> Range(r1.Ranges() @ [(a1, a1)])) <!> """AddrA1 "," R ==(4)=="""
    let RA1_5_UNION = (* AddrA1 RTO R (5)*) pipe3 AddrA1 RTO RangeA1Union (fun a1 a2 r2 -> Range ((a1,a2) :: r2.Ranges())) <!> "AddrA1 RTO R ==(5)=="
    do RangeA1UnionImpl :=   (attempt RA1_5_UNION)
                        <|> (attempt RA1_4_UNION)
                        <|> (attempt RA1_3)
                        <|> (attempt RA1_1)
                        <|> RA1_2
    do RangeA1NoUnionImpl := pipe2 AddrA1 RA1TO_2 (fun a1 a2 -> Range(a1, a2))
                        
    do RangeR1C1Impl := RangeA1Union

    let RangeWithUnion = (attempt RangeA1Union) <!> "Union Range"
    let RangeNoUnion = (attempt RangeA1NoUnion) <!> "No-Union Range"

    let RangeAny = ((attempt RangeWithUnion) <|> RangeWithUnion) <!> "RangeAny"

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
    let RangeReferenceWorkbook R = getUserState >>=
                                    fun (us: Env) ->
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
    let RangeReferenceWorksheet R = getUserState >>=
                                    fun (us: Env) ->
                                        pipe2
                                            (WorksheetName .>> pstring "!")
                                            R
                                            (fun wsname rng -> ReferenceRange(Env(us.Path, us.WorkbookName, wsname), rng) :> Reference)
                                    <!> "RangeReferenceWorksheet"
    let RangeReferenceNoWorksheet R = getUserState >>=
                                        fun (us: Env) ->
                                            R
                                            |>> (fun rng -> ReferenceRange(us, rng) :> Reference)
                                    <!> "RangeReferenceNOWorksheet"
    let RangeReference R = (attempt (RangeReferenceWorkbook R))
                         <|> (attempt (RangeReferenceWorksheet R))
                         <|> RangeReferenceNoWorksheet R
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
    let Reference R = (attempt (RangeReference R))
                        <|> (attempt AddressReference)
                        <|> (attempt ConstantReference)
                        <|> (attempt StringReference)
                        <|> NamedReference
                    <!> "Reference"
    // Functions
//    let FunctionName = (pstring "INDIRECT" >>. pzero) <|> many1Satisfy (fun c -> isLetter(c))
//                       <!> "FunctionName"

    let ArityNFunctionNameMaker n xs = xs |> List.map (fun name -> attempt (pstring name)) |> choice <!> ("Arity" + n.ToString() + "FunctionName")
    let Arity1FunctionName: P<string> = ArityNFunctionNameMaker 1 ["ABS"]
    let Arity2FunctionName: P<string> = ArityNFunctionNameMaker 2 ["SUMX2MY2"]
    let Arity3FunctionName: P<string> = ArityNFunctionNameMaker 3 [""]
    let Arity4FunctionName: P<string> = ArityNFunctionNameMaker 4 [""]
    let Arity5FunctionName: P<string> = ArityNFunctionNameMaker 5 [""]
    let Arity6FunctionName: P<string> = ArityNFunctionNameMaker 6 ["ACCRINT"]
    let Arity7FunctionName: P<string> = ArityNFunctionNameMaker 7 ["ACCRINT"]
    let Arity8FunctionName: P<string> = ArityNFunctionNameMaker 8 ["ACCRINT"]

    let arityNameArr: P<string>[] = 
        [|
            Arity1FunctionName;
            Arity2FunctionName;
            Arity3FunctionName;
            Arity4FunctionName;
            Arity5FunctionName;
            Arity6FunctionName;
            Arity7FunctionName;
            Arity8FunctionName;
        |]
    let FunctionNamesForArity i: P<string> = arityNameArr.[i-1]

    let VarArgsFunctionName: P<string> = ["SUM"] |> List.map (fun name -> pstring name) |> choice <!> "VarArgsFunctionName"

    let ArgumentsN R n = (pipe2
                            (parray (n-1) ((ExpressionDecl R) .>> pstring ",") )
                            (ExpressionDecl R)
                            (fun exprArr expr -> Array.toList exprArr @ [expr] ) <!> ("Arguments" + n.ToString())
                          )

    let ArityNFunction n R =
                // here, we ignore whatever Range context we are given
                // and use RangeNoUnion instead
                getUserState >>=
                fun us ->
                    pipe2
                        (FunctionNamesForArity(n) .>> pstring "(")
                        ((ArgumentsN RangeNoUnion n) .>> pstring ")")
                        (fun fname arglist -> ReferenceFunction(us, fname, arglist, Some(n)) :> Reference)
                <!> ("Arity" + n.ToString() + "Function")

    let VarArgsFunction R =
                  // here, we ignore whatever Range context we are given
                  // and use RangeAnyUnion instead (i.e., try both)
                  getUserState >>=
                    fun us ->
                        pipe2
                            (VarArgsFunctionName .>> pstring "(")
                            ((ArgumentList RangeAny) .>> pstring ")")
                            (fun fname arglist -> ReferenceFunction(us, fname, arglist, None) :> Reference)
                   <!> "VarArgsFunction"

    let Function R: P<Reference> = (arityNameArr |> Array.mapi (fun i _ -> (attempt (ArityNFunction (i + 1) R))) |> choice
                                    <|> VarArgsFunction R
                                    ) <!> "Function"
    
    do ArgumentListImpl := fun (R: P<Range>) -> sepBy (ExpressionDecl R) (spaces >>. pstring "," .>> spaces) <!> "ArgumentList"

    // Binary arithmetic operators
    let BinOpChar = spaces >>. satisfy (fun c -> c = '+' || c = '-' || c = '/' || c = '*' || c = '<' || c = '>' || c = '=' || c = '^' || c = '&') .>> spaces
    let BinOp2Char = spaces >>. ((attempt (regex "<=")) <|> (attempt (regex ">=")) <|> regex "<>") .>> spaces
    let BinOpLong R: P<string*Expression> = pipe2 BinOp2Char (ExpressionDecl R) (fun op rhs -> (op, rhs))
    let BinOpShort R: P<string*Expression> = pipe2 BinOpChar (ExpressionDecl R) (fun op rhs -> (op.ToString(), rhs))
    let BinOp R: P<string*Expression> = ((attempt (BinOpLong R)) <|> (BinOpShort R)) <!> "BinaryOp"

    // Unary operators
    let UnaryOpChar = (spaces >>. satisfy (fun c -> c = '+' || c = '-') .>> spaces) <!> "UnaryOp"

    // Expressions
    let ParensExpr(R: P<Range>): P<Expression> = ((between (pstring "(") (pstring ")") (ExpressionDecl R)) |>> ParensExpr) <!> "ParensExpr"
    let ExpressionAtom(R: P<Range>): P<Expression> = (((attempt (Function R)) <|> (Reference R)) |>> ReferenceExpr) <!> "ExpressionAtom"
    let ExpressionSimple(R: P<Range>): P<Expression> = ((attempt (ExpressionAtom R)) <|> (ParensExpr R)) <!> "ExpressionSimple"
    let UnaryOpExpr(R: P<Range>): P<Expression> = pipe2 UnaryOpChar (ExpressionDecl R) (fun op rhs -> UnaryOpExpr(op, rhs)) <!> "UnaryOpExpr"
    let BinOpExpr(R: P<Range>): P<Expression> = pipe2 (ExpressionSimple R) (BinOp R) (fun lhs (op, rhs) -> BinOpExpr(op, lhs, rhs)) <!> "BinaryOpExpr"
    
    // This is very confusing without some context:
    // BASICALLY, parsing Excel argument lists is ambiguous without knowing
    // the arity of the function that you are parsing.  Thus the Expression parser
    // takes a Range parser that either parses a Range with or without union semantics.
    // The appropriate parser is chosen by the Function parser.
    // Note that the top-level Formula parser tries to parse both (RangeAny).
    do ExpressionDeclImpl :=
        fun (R: P<Range>) ->
            (
                ((attempt (UnaryOpExpr R))
                <|> (attempt (BinOpExpr R))
                <|> (ExpressionSimple R))
            )
            <!> "Expression"

    // Formulas
    let Formula = (pstring "=" .>> spaces >>. (ExpressionDecl RangeAny) .>> eof) <!> "Formula"