module Grammar
    open FParsec
    open AST
    open System.Text.RegularExpressions

    let DEBUG_MODE = false
    let NO_OPTIMIZATIONS = false

    //#region OPTIMIZED_IMPLEMENTATIONS
    // this is an optimized "choice" parser that disables
    // error message logging when not in debug mode
    let (<||>) (p1: Parser<'a,'u>) (p2: Parser<'a,'u>) : Parser<'a,'u> =
        if NO_OPTIMIZATIONS then
            fun stream ->
                let mutable stateTag = stream.StateTag
                let mutable reply = p1 stream
                if reply.Status = Error && stateTag = stream.StateTag then
                    let error = reply.Error
                    reply <- p2 stream
                    if stateTag = stream.StateTag then
                        reply.Error <- mergeErrors reply.Error error
                reply
        else
            fun stream ->
                let mutable stateTag = stream.StateTag
                let mutable reply = p1 stream
                if reply.Status = Error && stateTag = stream.StateTag then
                    reply <- p2 stream
                reply

    // this is an optimized "attempt" parser that disables
    // error message logging when not in debug mode
    let attempt_opt (p: Parser<'a,'u>) : Parser<'a,'u> =
        if NO_OPTIMIZATIONS then
            fun stream ->
                // state is only declared mutable so it can be passed by ref, it won't be mutated
                let mutable state = CharStreamState(stream) // = stream.State (manually inlined)
                let mutable reply = p stream
                if reply.Status <> Ok then
                    if state.Tag <> stream.StateTag then
                        reply.Error  <- nestedError stream reply.Error
                        reply.Status <- Error // turns FatalErrors into Errors
                        stream.BacktrackTo(&state) // passed by ref as a (slight) optimization
                    elif reply.Status = FatalError then
                        reply.Status <- Error
                reply
        else
            fun stream ->
                // state is only declared mutable so it can be passed by ref, it won't be mutated
                let mutable state = CharStreamState(stream) // = stream.State (manually inlined)
                let mutable reply = p stream
                if reply.Status <> Ok then
                    if state.Tag <> stream.StateTag then
                        reply.Status <- Error // turns FatalErrors into Errors
                        stream.BacktrackTo(&state) // passed by ref as a (slight) optimization
                    elif reply.Status = FatalError then
                        reply.Status <- Error
                reply

    // this is an optimized "choice" parser that disables
    // error message logging when not in debug mode
    let choice_opt (ps: seq<Parser<'a,'u>>)  =
        if NO_OPTIMIZATIONS then
            match ps with
            | :? (Parser<'a,'u>[]) as ps ->
                if ps.Length = 0 then pzero
                else
                    fun stream ->
                        let stateTag = stream.StateTag
                        let mutable error = NoErrorMessages
                        let mutable reply = ps.[0] stream
                        let mutable i = 1
                        while reply.Status = Error && stateTag = stream.StateTag && i < ps.Length do
                            error <- mergeErrors error reply.Error
                            reply <- ps.[i] stream
                            i <- i + 1
                        if stateTag = stream.StateTag then
                            error <- mergeErrors error reply.Error
                            reply.Error <- error
                        reply
            | :? (Parser<'a,'u> list) as ps ->
                match ps with
                | [] -> pzero
                | hd::tl ->
                    fun stream ->
                        let stateTag = stream.StateTag
                        let mutable error = NoErrorMessages
                        let mutable hd, tl = hd, tl
                        let mutable reply = hd stream
                        while reply.Status = Error && stateTag = stream.StateTag
                              && (match tl with
                                  | h::t -> hd <- h; tl <- t; true
                                  | _ -> false)
                           do
                            error <- mergeErrors error reply.Error
                            reply <- hd stream
                        if stateTag = stream.StateTag then
                            error <- mergeErrors error reply.Error
                            reply.Error <- error
                        reply
            | _ -> fun stream ->
                       use iter = ps.GetEnumerator()
                       if iter.MoveNext() then
                           let stateTag = stream.StateTag
                           let mutable error = NoErrorMessages
                           let mutable reply = iter.Current stream
                           while reply.Status = Error && stateTag = stream.StateTag && iter.MoveNext() do
                               error <- mergeErrors error reply.Error
                               reply <- iter.Current stream
                           if stateTag = stream.StateTag then
                               error <- mergeErrors error reply.Error
                               reply.Error <- error
                           reply
                       else
                           Reply()
        else
            match ps with
            | :? (Parser<'a,'u>[]) as ps ->
                if ps.Length = 0 then pzero
                else
                    fun stream ->
                        let stateTag = stream.StateTag
                        let mutable reply = ps.[0] stream
                        let mutable i = 1
                        while reply.Status = Error && stateTag = stream.StateTag && i < ps.Length do
                            reply <- ps.[i] stream
                            i <- i + 1
                        reply
            | :? (Parser<'a,'u> list) as ps ->
                match ps with
                | [] -> pzero
                | hd::tl ->
                    fun stream ->
                        let mutable hd, tl = hd, tl
                        let mutable reply = hd stream
                        while reply.Status = Error && stream.StateTag = stream.StateTag
                              && (match tl with
                                  | h::t -> hd <- h; tl <- t; true
                                  | _ -> false)
                           do
                            reply <- hd stream
                        reply
            | _ -> fun stream ->
                       use iter = ps.GetEnumerator()
                       if iter.MoveNext() then
                           let mutable reply = iter.Current stream
                           while reply.Status = Error && stream.StateTag = stream.StateTag && iter.MoveNext() do
                               reply <- iter.Current stream
                           reply
                       else
                           Reply()

    //#endregion OPTIMIZED_IMPLEMENTATIONS

    let (<!>) (p: Parser<_,_>) label : Parser<_,_> =
        if DEBUG_MODE then
            fun stream ->
                let s = String.replicate (int (stream.Position.Column)) " "
                printfn "%d%sTrying %s(\"%s\")" (stream.Index) s label (stream.PeekString 1000)
//                printfn "At index: %d, string remaining: %s" (stream.Index) (stream.PeekString 1000)
                let reply = p stream    // set a breakpoint here when debugging
                printfn "%d%s%s(\"%s\") (%A)" (stream.Index) s label (stream.PeekString 1000) reply.Status
                reply
        else
            p

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
     * <||>      The infix "choice combinator" <||> applies the parser on the right side if the parser on the left side fails.
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

    // Grammar forward references
    let (ArgumentList: P<Range> -> P<Expression list>, ArgumentListImpl) = createParserForwardedToRefWithArguments()
    let (ExpressionDecl: P<Range> -> P<Expression>, ExpressionDeclImpl) = createParserForwardedToRefWithArguments()
    let (RangeA1Union: P<Range>, RangeA1UnionImpl) = createParserForwardedToRef()
    let (RangeA1NoUnion: P<Range>, RangeA1NoUnionImpl) = createParserForwardedToRef()
    let (RangeR1C1: P<Range>, RangeR1C1Impl) = createParserForwardedToRef()

    // space-insensitive commas
    let Comma = spaces >>. pstring "," .>> spaces

    // Addresses
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
                        (attempt_opt AddrIndirect)
                        <||> (pipe2
                                AddrR
                                AddrC
                                (fun row col ->
                                    // TODO: R1C1 absolute/relative
                                    Address.fromR1C1withMode(row, col, AddressMode.Relative, AddressMode.Relative, us.WorksheetName, us.WorkbookName, us.Path)))
                    <!> "AddrR1C1"
    let AbsOrNot : P<AddressMode> = (attempt_opt ((pstring "$") >>% Absolute) <||> (pstring "" >>% Relative)) 
    let AddrA = many1Satisfy isAsciiUpper
    let AddrAAbs = pipe2 AbsOrNot AddrA (fun mode col -> (mode, col))
    let Addr1 = pint32
    let Addr1Abs = pipe2 AbsOrNot Addr1 (fun mode row -> (mode, row))
    let AddrA1 = getUserState >>=
                    fun (us: Env) ->
                        (attempt_opt AddrIndirect)
                        <||> (pipe2
                            AddrAAbs
                            Addr1Abs
                            (fun col row ->
                                let (colMode,colA) = col
                                let (rowMode,row1) = row
                                Address.fromA1withMode(row1, colA, rowMode, colMode, us.WorksheetName, us.WorkbookName, us.Path)
                                ))
                    
    let AnyAddr = ((attempt_opt AddrIndirect)
                  <||> (attempt_opt AddrR1C1)
                  <||> AddrA1)
                  <!> "AnyAddr"

    // Ranges
    let RA1TO_1 = (pstring ":" >>. AddrA1 .>> Comma) <!> "RA1TO_1"
    let RA1TO_2 = (pstring ":" >>. AddrA1) <!> "RA1TO_2"
    let RTO = ((attempt_opt RA1TO_1) <||> RA1TO_2) <!> "RTO"

    let RA1TC_1 = (Comma >>. AddrA1 .>> Comma) <!> "RA1TC_1"
    let RA1TC_2 = (Comma >>. AddrA1) <!> "RA1TC_2"
    let RTC = ((attempt_opt RA1TC_1) <||> RA1TC_2) <!> "RTC"

    let RA1_1 = (* AddrA1 RTC (1)*) (pipe2 AddrA1 RTC (fun a1 a2 -> Range([(a2,a2); (a1,a1)]))) <!> "RA1_1"
    let RA1_2 = (* AddrA1 RTO (2)*) (pipe2 AddrA1 RTO (fun a1 a2 -> Range(a1,a2))) <!> "RA1_2"

    let rat_fun = fun a1 a2 a3 -> Range([(a1,a2);(a3,a3)])
    let RA1_3a = (pipe3 AddrA1 RA1TO_1 RA1TC_1 rat_fun) <!> "RA1_3a"
    let RA1_3b = (pipe3 AddrA1 RA1TO_1 RA1TC_2 rat_fun) <!> "RA1_3b"
    let RA1_3c = (pipe3 AddrA1 RA1TO_2 RA1TC_1 rat_fun) <!> "RA1_3c"
    let RA1_3d = (pipe3 AddrA1 RA1TO_2 RA1TC_2 rat_fun) <!> "RA1_3d"
    let RA1_3 = (* AddrA1 RTO RTC (3)*) 
                ((attempt_opt RA1_3a) <||> (attempt_opt RA1_3b) <||> (attempt_opt RA1_3c) <||> RA1_3d) <!> "RA1_3"

    let RA1_4_UNION = (* AddrA1 "," R (4)*) (pipe2 (AddrA1 .>> Comma) RangeA1Union (fun a1 r1 -> Range(r1.Ranges() @ [(a1, a1)]))) <!> "RA1_4_UNION"
    let RA1_5_UNION = (* AddrA1 RTO R (5)*) (pipe3 AddrA1 RTO RangeA1Union (fun a1 a2 r2 -> Range ((a1,a2) :: r2.Ranges()))) <!> "RA1_5_UNION"
    do RangeA1UnionImpl :=   (attempt_opt RA1_5_UNION)
                        <||> (attempt_opt RA1_4_UNION)
                        <||> (attempt_opt RA1_3)
                        <||> (attempt_opt RA1_1)
                        <||> RA1_2
    do RangeA1NoUnionImpl := pipe2 AddrA1 RA1TO_2 (fun a1 a2 -> Range(a1, a2))
                        
    do RangeR1C1Impl := RangeA1Union

    let RangeWithUnion = (attempt_opt RangeA1Union) <!> "Union Range"
    let RangeNoUnion = (attempt_opt RangeA1NoUnion) <!> "No-Union Range"

    let RangeAny = ((attempt_opt RangeWithUnion) <||> RangeWithUnion) <!> "RangeAny"

    // Worksheet Names
    let WorksheetNameQuoted =
        let NormalChar = satisfy ((<>) '\'')
        let EscapedChar = pstring "''" |>> (fun s -> ''')
        between (pstring "'") (pstring "'")
                (many1Chars (NormalChar <||> EscapedChar))
    let WorksheetNameUnquoted = (many1Satisfy (fun c -> isWSChar(c)))
    let WorksheetName = (WorksheetNameQuoted <||> WorksheetNameUnquoted)
    // Workbook Names (this may be too restrictive)
    let Path = many1Satisfy ((<>) '[')
    let WorkbookName = between
                        (pstring "[")
                        (pstring "]")
                        (many1Satisfy (fun c -> c <> '[' && c <> ']'))
    let Workbook = (
                        (Path |>> Some)
                        <||> ((pstring "") >>% None)
                   )
                   .>>. WorkbookName

    // References
    // References consist of the following parts:
    //   A workbook name prefix
    //   A worksheet name prefix
    //   A single-cell address ("Address") or multi-cell address ("Range")
    let RRWQuoted = (between (pstring "'") (pstring "'") (Workbook .>>. WorksheetNameUnquoted)) <!> "RRWQuoted"
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
    let RangeReference R = (attempt_opt (RangeReferenceWorkbook R))
                         <||> (attempt_opt (RangeReferenceWorksheet R))
                         <||> RangeReferenceNoWorksheet R
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
    let AddressReference = (attempt_opt AddressReferenceWorkbook)
                           <||> (attempt_opt AddressReferenceWorksheet)
                           <||> AddressReferenceNoWorksheet
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
                                    (attempt_opt
                                        (pfloat .>> pstring "%"
                                        |>> fun r -> ReferenceConstant(us, r / 100.0) :> Reference))
                                    <||> (pfloat |>> (fun r -> ReferenceConstant(us, r) :> Reference))
                            <!> "ConstantReference"
    let BooleanReference = getUserState >>=
                             fun us ->
                                 Array.map (fun b -> attempt_opt (pstring b)) [| "TRUE"; "FALSE" |]
                                 |> choice_opt
                                 |>> fun r -> ReferenceBoolean(us, System.Boolean.Parse(r)) :> Reference
                           <!> "BooleanReference"

    // This helper function sorts strings from
    // longest length to shortest length; length
    // ties are sorted lexicographically; ensures
    // that a substring is never matched, consumed,
    // and then returned to the calling parser which
    // then fails.
    let lmf(sa: string[]) : string[] =
        sa |>
        Array.sortWith
            (fun e1 e2 ->
                if e1.Length < e2.Length then
                    1
                elif e1.Length = e2.Length then
                    e1.CompareTo e2
                else
                    -1
            )

    // Functions
    let ArityNFunctionNameMaker n xs = xs |> Array.map (fun name -> attempt_opt (pstring name)) |> choice_opt <!> ("Arity" + n.ToString() + "FunctionName")
    let ArityAtLeastNFunctionNameMaker nplus xs = xs |> Array.map (fun name -> attempt_opt (pstring name)) |> choice_opt <!> ("Arity" + nplus.ToString() + "+FunctionName")

    let Arity0Names: string[] = [|"COLUMN"; "NA"; "NOW"; "PI"; "RAND"; "ROW"; "SHEET"; "SHEETS"; "TODAY"|]
    let Arity0FunctionName: P<string> = ArityNFunctionNameMaker 0 (lmf Arity0Names)

    let Arity1Names: string[] = [|"ABS"; "ACOS"; "ACOSH"; "ACOT"; "ACOTH"; "ARABIC"; "AREAS"; "ASC"; "ASIN"; "ASINH"; "ATAN"; "ATANH";
        "BAHTTEXT"; "BIN2DEC"; "BIN2HEX"; "BIN2OCT"; "CEILING.PRECISE"; "CELL"; "CHAR"; "CLEAN"; "CODE"; "COLUMN"; "COLUMNS"; "COS"; "COSH"; "COT"; "COTH"; "COUNTBLANK";
        "CSC"; "CSCH"; "CUBESETCOUNT"; "DAY"; "DBCS"; "DEC2BIN"; "DEC2HEX"; "DEC2OCT"; "DEGREES"; "DELTA"; "DOLLAR"; "ENCODEURL"; "ERF"; "ERF.PRECISE"; "ERFC"; "ERFC.PRECISE";
        "ERROR.TYPE"; "EVEN"; "EXP"; "FACT"; "FACTDOUBLE"; "FISHER"; "FISHERINV"; "FIXED"; "FLOOR.PRECISE"; "FORMULATEXT"; "GAMMA"; "GAMMALN"; "GAMMALN.PRECISE"; "GAUSS";
        "GESTEP"; "GROWTH"; "HEX2BIN"; "HEX2DEC"; "HEX2OCT"; "HOUR"; "HYPERLINK"; "IMABS"; "IMAGINARY"; "IMARGUMENT"; "IMCONJUGATE"; "IMCOS"; "IMCOSH"; "IMCOT"; "IMCSC";
        "IMCSCH"; "IMEXP"; "IMLN"; "IMLOG10"; "IMLOG2"; "IMREAL"; "IMSEC"; "IMSECH"; "IMSIN"; "IMSINH"; "IMSQRT"; "IMTAN"; "INDEX"; "INFO"; "INT";  "IRR";
        "ISBLANK"; "ISERR"; "ISERROR"; "ISEVEN"; "ISFORMULA"; "ISLOGICAL"; "ISNA"; "ISNONTEXT"; "ISNUMBER"; "ISODD"; "ISREF"; "ISTEXT"; "ISO.CEILING"; "ISOWEEKNUM";  "JIS";
        "LEFT"; "LEFTB"; "LEN"; "LENB"; "LINEST"; "LN"; "LOG"; "LOG10"; "LOGEST"; "LOWER"; "MDETERM"; 
        "MEDIAN"; "MINUTE"; "MDETERM"; "MDETERM"; "MINVERSE"; "MONTH"; "MUNIT";
        "N"; "NORMSDIST"; "NORM.S.INV"; "NORMSINV"; "NOT"; "NUMBERVALUE"; 
        "OCT2BIN"; "OCT2DEC"; "OCT2HEX"; "ODD"; "OR"; 
        "PHI"; "PHONETIC"; "PRODUCT"; "PROPER";
        "RADIANS"; "RIGHT"; "RIGHTB"; "ROMAN"; "ROW"; "ROWS";
        "SEC"; "SECH"; "SECOND"; "SHEET"; "SHEETS"; "SIGN"; "SIN"; "SINH"; "SKEW"; "SKEW.P"; "SQL.REQUEST"; "SQRT"; "SQRTPI"; "STDEV"; "STDEV.P"; "STDEV.S"; "STDEVA"; "STDEVP"; "STDEVPA"; "SUMPRODUCT"; "SUMSQ"; 
        "T"; "TAN"; "TANH"; "TIMEVALUE"; "TRANSPOSE"; "TREND"; "TRIM"; "TRUNC"; "TYPE";
        "UNICHAR"; "UNICODE"; "UPPER";
        "VALUE"; "VAR"; "VAR.P"; "VAR.S"; "VARA"; "VARP"; "VARPA"; 
        "WEBSERVICE"; "WEEKDAY"; "WEEKNUM"; 
        "XOR";  
        "YEAR"|]
    let Arity1FunctionName: P<string> = ArityNFunctionNameMaker 1 (lmf Arity1Names)

    let Arity2Names: string[] = [|"ADDRESS"; "ATAN2"; "AVERAGEIF"; "BASE"; "BESSELI"; "BESSELJ"; "BESSELK"; "BESSELY";
        "BIN2HEX"; "BIN2OCT"; "BITAND"; "BITLSHIFT"; "BITOR"; "BITRSHIFT"; "BITXOR"; "CEILING"; "CEILING.PRECISE"; "CELL"; "CHIDIST"; "CHIINV"; "CHITEST";
        "CHISQ.DIST.RT"; "CHISQ.TEST"; "COMBIN"; "COMBINA"; "COMPLEX"; "COUNTIF"; "COVAR"; "COVARIANCE.P"; "COVARIANCE.S"; "CUBEMEMBER"; "CUBERANKEDMEMBER";
        "CUBESET"; "DATEVALUE"; "DAYS"; "DAYS360"; "DEC2BIN"; "DEC2HEX"; "DEC2OCT"; "DECIMAL"; "DELTA"; "DOLLAR"; "DOLLARDE"; "DOLLARFR"; "EDATE"; "EFFECT";
        "EOMONTH"; "ERF"; "EXACT"; "F.DIST"; "FILTERXML"; "FIND"; "FINDB"; "FIXED"; "FLOOR"; "FLOOR.PRECISE"; "FORECAST.ETS.SEASONALITY"; "FREQUENCY"; "F.TEST";
        "FTEST"; "FVSCHEDULE"; "GESTEP"; "GROWTH"; "HEX2BIN"; "HEX2OCT"; "HYPERLINK"; "IF"; "IFERROR"; "IFNA"; "IMDIV"; "IMPOWER"; "IMSUB"; "INDEX";
        "INTERCEPT"; "IRR"; "ISO.CEILING"; "LARGE"; "LEFT"; "LEFTB"; "LINEST"; "LOG"; "LOGEST"; "LOOKUP";
        "MATCH";
        "MEDIAN"; "MMULT"; "MOD"; "MROUND";    
        "NETWORKDAYS.INTL"; "NETWORKDAYS"; "NOMINAL"; "NORM.S.DIST"; "NPV"; "NUMBERVALUE"; 
        "OCT2BIN"; "OCT2DEC"; "OCT2HEX"; "ODDFPRICE"; "OR";
        "PEARSON"; "PERCENTILE.EXC"; "PERCENTILE.INC"; "PERCENTILE"; "PERCENTRANK.EXC"; "PERCENTRANK.INC"; "PERCENTRANK"; "PERMUT"; "PERMUTATIONA"; "POWER"; "PROB"; "PRODUCT"; 
        "QUARTILE"; "QUARTILE.EXC"; "QUARTILE.INC"; "QUOTIENT";
        "RANDBETWEEN"; "RANK.AVG"; "RANK.EQ"; "RANK"; "REGISTER.ID"; "REPT"; "RIGHT"; "RIGHTB"; "ROMAN"; "ROUND"; "ROUNDDOWN"; "ROUNDUP"; "RSQ";
        "SEARCH"; "SEARCHB"; "SKEW"; "SKEW.P"; "SLOPE"; "SMALL"; "SQL.REQUEST"; "STDEV"; "STDEV.P"; "STDEV.S"; "STDEVA"; "STDEVP"; "STDEVPA"; "STEYX"; "SUBTOTAL"; "SUMIF"; "SUMPRODUCT"; "SUMSQ"; "SUMX2MY2"; "SUMX2PY2"; "SUMXMY2"; 
        "TBILLYIELD"; "T.DIST.RT"; "TEXT"; "T.INV"; "TINV"; "T.INV.2T"; "TREND"; "TRIMMEAN"; "TRUNC";
        "VAR"; "VAR.P"; "VAR.S"; "VARA"; "VARP"; "VARPA";
        "WEEKDAY"; "WEEKNUM"; "WORKDAY"; "WORKDAY.INTL";
        "XIRR"; "XOR";
        "YEARFRAC"; 
        "ZTEST"; "Z.TEST"|]
    let Arity2FunctionName: P<string> = ArityNFunctionNameMaker 2 (lmf Arity2Names)

    let Arity3Names: string[] = [|"ADDRESS"; "AVERAGEIF"; "BASE"; "BETADIST"; "BETAINV"; "BETA.INV"; "BINOM.DIST.RANGE"; "BINOM.INV";
        "CEILING.MATH"; "CHISQ.DIST"; "COMPLEX"; "CONFIDENCE"; "CONFIDENCE.NORM"; "CONFIDENCE.T"; "CONVERT"; "CORREL"; "COUPDAYBS"; "COUPDAYS"; "COUPDAYSNC"; "COUPNCD";
        "COUPNUM"; "COUPPCD"; "CRITBINOM"; "CUBEKPIMEMBER"; "CUBEMEMBER"; "CUBEMEMBERPROPERTY"; "CUBERANKEDMEMBER"; "CUBESET";
        "DATE"; "DATEDIF"; "DAVERAGE"; "DAYS360"; "DCOUNT"; "DCOUNTA"; "DGET"; "DMAX"; "DMIN"; "DPRODUCT"; "DSTDEV"; "DSTDEVP"; "DSUM"; "DVAR"; "DVARP"; "EXPON.DIST";
        "EXPONDIST"; "FDIST"; "F.DIST.RT"; "FIND"; "FINDB";"F.INV"; "F.INV.RT"; "FINV"; "FIXED"; "FLOOR.MATH"; "FORECAST"; "FORECAST.ETS"; "FORECAST.ETS.SEASONALITY";
        "FORECAST.LINEAR"; "FORECAST.ETS.CONFINT"; "FORECAST.ETS.STAT"; "FV"; "GAMMA.INV"; "GAMMAINV"; "GROWTH"; "HLOOKUP";
        "IF"; "INDEX"; "LINEST"; "LOGEST"; "LOGINV"; "LOGNORMDIST"; "LOGNORM.INV"; "LOOKUP";
        "MATCH";
        "MEDIAN"; "MIRR"; "MID"; "MIDB"; "MULTINOMIAL";
        "NEGBINOMDIST"; "NETWORKDAYS"; "NETWORKDAYS.INTL"; "NORM.INV"; "NORMINV"; "NPER"; "NPV"; "NUMBERVALUE"; 
        "OFFSET"; "OR";
        "PDURATION"; "PERCENTRANK.EXC"; "PERCENTRANK.INC"; "PERCENTRANK"; "PMT"; "POISSON.DIST"; "POISSON"; "PROB"; "PRODUCT"; "PV"; 
        "RANK.AVG"; "RANK.EQ"; "RANK"; "RATE"; "REGISTER.ID"; "RRI"; "RTD"; 
        "SEARCH"; "SEARCHB"; "SKEW"; "SKEW.P"; "SLN"; "SQL.REQUEST"; "STANDARDIZE"; "STDEV"; "STDEV.P"; "STDEV.S"; "STDEVA"; "STDEVP"; "STDEVPA"; "SUBSTITUTE"; "SUBTOTAL"; "SUMIF"; "SUMIFS"; "SUMPRODUCT"; "SUMSQ"; "SWITCH";
        "TBILLEQ"; "TBILLPRICE"; "T.DIST"; "TDIST"; "T.DIST.2T"; "TEXTJOIN"; "TIME"; "TREND";
        "VAR"; "VAR.P"; "VAR.S"; "VARA"; "VARP"; "VARPA"; "VLOOKUP"; 
        "WORKDAY"; "WORKDAY.INTL";
        "XIRR"; "XNPV"; "XOR";
        "YEARFRAC"; 
        "ZTEST"; "Z.TEST"|]
    let Arity3FunctionName: P<string> = ArityNFunctionNameMaker 3 (lmf Arity3Names)

    let Arity4Names: string[] = [|"ADDRESS"; "ACCRINTM"; "BETADIST"; "BETA.DIST"; "BETAINV"; "BETA.INV"; "BINOMDIST"; "BINOM.DIST";
        "BINOM.DIST.RANGE"; "COUPDAYBS"; "COUPDAYS"; "COUPDAYSNC"; "COUPNCD"; "COUPNUM"; "COUPPCD"; "CUBESET"; "DB"; "DDB"; "DISC"; "F.DIST"; "FORECAST.ETS";
        "FORECAST.ETS.SEASONALITY"; "FORECAST.ETS.CONFINT"; "FORECAST.ETS.STAT"; "FV"; "GAMMA.DIST"; "GAMMADIST"; "GROWTH"; "HLOOKUP"; "HYPGEOMDIST"; "INTRATE";
        "LINEST"; "LOGEST"; "LOGNORM.DIST"; 
        "MEDIAN"; "IPMT"; "ISPMT";
        "NEGBINOM.DIST"; "NETWORKDAYS.INTL"; "NORM.DIST"; "NORMDIST"; "NPER"; "NPV"; 
        "OFFSET"; "OR";
        "PMT"; "PPMT"; "PRICEDISC"; "PRICEMAT"; "PROB"; "PRODUCT"; "PV"; 
        "RATE"; "RECEIVED"; "REPLACE"; "REPLACEB"; "RTD"; 
        "SERIESSUM"; "SKEW"; "SKEW.P"; "SQL.REQUEST"; "STDEV"; "STDEV.P"; "STDEV.S"; "STDEVA"; "STDEVP"; "STDEVPA"; "SUBSTITUTE"; "SUBTOTAL"; "SUMPRODUCT"; "SUMSQ"; "SWITCH"; "SYD";
        "TEXTJOIN"; "TREND";"T.TEST"; "TTEST";
        "VAR"; "VAR.P"; "VAR.S"; "VARA"; "VARP"; "VARPA"; "VLOOKUP"; 
        "WEIBULL"; "WEIBULL.DIST"; "WORKDAY.INTL";
        "XOR";
        "YIELDDISC"|]
    let Arity4FunctionName: P<string> = ArityNFunctionNameMaker 4 (lmf Arity4Names)

    let Arity5Names = [|"ADDRESS"; "ACCRINTM"; "BETADIST"; "BETA.DIST"; "BETAINV"; "BETA.INV"; "CUBESET"; "DB"; "DDB";
        "DISC"; "DURATION"; "EUROCONVERT"; "FORECAST.ETS"; "FORECAST.ETS.CONFINT"; "FORECAST.ETS.STAT"; "FV"; "HYPGEOM.DIST"; "INTRATE"; "IPMT"; 
        "MEDIAN"; "MDURATION"; "MDURATION";
        "NPER"; "NPV";
        "OFFSET"; "OR";
        "PMT"; "PPMT"; "PRICEDISC"; "PRICEMAT"; "PRODUCT"; "PV"; 
        "RATE"; "RECEIVED"; "RTD"; 
        "SKEW"; "SKEW.P"; "SQL.REQUEST"; "STDEV"; "STDEV.P"; "STDEV.S"; "STDEVA"; "STDEVP"; "STDEVPA"; "SUBTOTAL"; "SUMIFS"; "SUMPRODUCT"; "SUMSQ"; "SWITCH";
        "TEXTJOIN"; 
        "VAR"; "VAR.P"; "VARA"; "VARP";"VARPA"; "VDB"; 
        "XOR";
        "YIELDDISC"; "YIELDMAT"|]
    let Arity5FunctionName: P<string> = ArityNFunctionNameMaker 5 (lmf Arity5Names)

    let Arity6Names: string[] = [|"ACCRINT"; "AMORDEGRC"; "AMORLINC"; "BETA.DIST";
        "CUMIPMT"; "CUMPRINC"; "DURATION"; "FORECAST.ETS"; "FORECAST.ETS.CONFINT"; "FORECAST.ETS.STAT"; "IPMT"; 
        "MEDIAN"; "MDURATION"; "MDURATION"; 
        "NPV"; 
        "OR";
        "PPMT"; "PRICE"; "PRODUCT"; 
        "RATE"; "RTD"; 
        "SKEW"; "SKEW.P"; "STDEV"; "STDEV.P"; "STDEV.S"; "STDEVA"; "STDEVP"; "STDEVPA"; "SUBTOTAL"; "SUMPRODUCT"; "SUMSQ"; "SWITCH";
        "TEXTJOIN"; 
        "VAR"; "VAR.P"; "VAR.S"; "VARA"; "VARP"; "VARPA"; "VDB";
        "XOR";
        "YIELD"; "YIELDMAT"|]
    let Arity6FunctionName: P<string> = ArityNFunctionNameMaker 6 (lmf Arity6Names)

    let Arity7Names: string[] = [|"ACCRINT"; "AMORDEGRC"; "AMORLINC"; "FORECAST.ETS.CONFINT";
        "MEDIAN"; "MULTINOMIAL";
        "NPV"; 
        "ODDLPRICE"; "ODDLYIELD"; "OR";
        "PRICE"; "PRODUCT"; 
        "RTD"; 
        "STDEV"; "STDEV.P"; "STDEV.S"; "STDEVA"; "STDEVP"; "STDEVPA"; "SUBTOTAL"; "SUMIFS"; "SUMPRODUCT"; "SUMSQ"; "SWITCH";
        "TEXTJOIN"; 
        "VAR"; "VAR.P"; "VAR.S"; "VARA"; "VARP"; "VARPA"; "VDB";
        "XOR";
        "YIELD"|]
    let Arity7FunctionName: P<string> = ArityNFunctionNameMaker 7 (lmf Arity7Names)

    let Arity8Names: string[] = [|"ACCRINT";
        "MEDIAN"; "MULTINOMIAL";
        "NPV"; 
        "ODDFPRICE"; "ODDFYIELD"; "ODDLPRICE"; "ODDLYIELD"; "OR";
        "PRODUCT"; 
        "RTD"; 
        "STDEV"; "STDEV.P"; "STDEV.S"; "STDEVA"; "STDEVP"; "STDEVPA"; "SUBTOTAL"; "SUMPRODUCT"; "SUMSQ"; "SWITCH";
        "TEXTJOIN"; 
        "VAR"; "VAR.P"; "VAR.S"; "VARA"; "VARP"; "VARPA";
        "XOR"|]
    let Arity8FunctionName: P<string> = ArityNFunctionNameMaker 8 (lmf Arity8Names)

    let Arity9Names: string[] = [|"ACCRINT";
        "MEDIAN"; "MULTINOMIAL";
        "NPV"; 
        "ODDFPRICE"; "ODDFYIELD"; "OR";
        "PRODUCT"; 
        "RTD";
        "STDEV"; "STDEV.P"; "STDEV.S"; "STDEVA"; "STDEVP"; "STDEVPA"; "SUBTOTAL"; "SUMPRODUCT"; "SUMSQ"; "SWITCH";
        "VAR"; "VAR.P"; "VAR.S"; "VARA"; "VARP"; "VARPA";
        "XOR"|]
    let Arity9FunctionName: P<string> = ArityNFunctionNameMaker 9 (lmf Arity9Names)

    let ArityAtLeast1Names: string[] = [|"AND"; "AVEDEV"; "AVERAGEA"; "AVERAGE";
        "CALL"; "CONCATENATE"; "CONCAT"; "COUNTA"; "COUNT"; "CUBEVALUE"; "DEVSQ"; "GCD"; "GEOMEAN"; "HARMEAN"; "IMPRODUCT";
        "IMSUM"; "KURT"; "LCM"; "MAXA"; "MAX"; "MINA"; "MIN"; "MODE.MULT"; "MODE.SNGL"; "MODE"; "MULTINOMIAL"; |]
    let ArityAtLeast1FunctionName: P<string> = ArityAtLeastNFunctionNameMaker 1 (lmf ArityAtLeast1Names)

    let ArityAtLeast2Names: string[] = [|"CHOOSE"; "COUNTIFS"; "GETPIVOTDATA"; "IFS"; |]
    let ArityAtLeast2FunctionName: P<string> = ArityAtLeastNFunctionNameMaker 2 (lmf ArityAtLeast2Names)

    let ArityAtLeast3Names: string[] = [|"AGGREGATE"; "AVERAGEIFS"; "MAXIFS"; "MINIFS"; |]
    let ArityAtLeast3FunctionName: P<string> = ArityAtLeastNFunctionNameMaker 3 (lmf ArityAtLeast3Names)

    let VarArgsNames: string[] = [|"SUM"|]
    let VarArgsFunctionName: P<string> =
        (lmf VarArgsNames)
        |> Array.map (fun name -> pstring name) |> choice_opt <!> "VarArgsFunctionName"

    let arityNNameArr: P<string>[] = 
        [|
            Arity0FunctionName;
            Arity1FunctionName;
            Arity2FunctionName;
            Arity3FunctionName;
            Arity4FunctionName;
            Arity5FunctionName;
            Arity6FunctionName;
            Arity7FunctionName;
            Arity8FunctionName;
        |]

    let ReservedWords: P<string>[] = 
        [|
            Arity0Names;
            Arity1Names;
            Arity2Names;
            Arity3Names;
            Arity4Names;
            Arity5Names;
            Arity6Names;
            Arity7Names;
            Arity8Names;
            Arity9Names;
            ArityAtLeast1Names;
            ArityAtLeast2Names;
            ArityAtLeast3Names;
            VarArgsNames;
        |]
        |> Array.concat
        |> lmf
        |> Array.map (fun ns -> pstring ns)

    let arityAtLeastNNameArr: P<string>[] =
        [|
            // NOTE: "ArityAtLeast0" is just VarArgs
            ArityAtLeast1FunctionName;
            ArityAtLeast2FunctionName;
            ArityAtLeast3FunctionName;
        |]

    let FunctionNamesForArity i: P<string> = arityNNameArr.[i]
    let FunctionNamesForAtLeastArity i: P<string> = arityAtLeastNNameArr.[i-1]

    // from http://stackoverflow.com/a/22252504/480764
    let rec pmap f xs =
        match xs with
        | x :: xs -> parse { let! y = f x
                             let! ys = pmap f xs
                             return y :: ys }
        | [] -> parse { return [] }

    let Argument R = fun (i: int) -> (((ExpressionDecl R) .>> Comma) <!> ("Argument #" + i.ToString()))

    let ArgumentsN R n =
        (pipe2
            (pmap (Argument R) [1..n-1])
            ((ExpressionDecl R) <!> ("Argument #" + n.ToString() + " (last arg)"))
            (fun exprs expr -> exprs @ [expr] ) <!> ("Arguments" + n.ToString())
        )

    let ArgumentsAtLeastN R n =
        // What I really want is sepByN; this parser basically
        // does the same thing by requiring a mandatory match of
        // n-1 arguments, then one more, and then any number
        // of optional arguments.
        (pipe3
            (parray (n-1) ((ExpressionDecl R) .>> Comma) )
            (sepBy1 (ExpressionDecl R) Comma)
            (ArgumentList R)
            (fun exprArrReqd exprAtLeastOne varArgs -> (Array.toList exprArrReqd) @ exprAtLeastOne @ varArgs) <!> ("Arguments" + n.ToString() + "+")
        )

    let ArityNFunction n R =
                // here, we ignore whatever Range context we are given
                // and use RangeNoUnion instead
                if n = 0 then
                    getUserState >>=
                    fun us ->
                        (FunctionNamesForArity n .>> (pstring "()" <!> "Arity" + n.ToString() + "FunctionOPENCLOSEBRACE"))
                        |>>
                            (fun fname -> ReferenceFunction(us, fname, [], Fixed(0)) :> Reference)
                else
                    getUserState >>=
                    fun us ->
                        pipe2
                            (FunctionNamesForArity n .>> (pstring "(" <!> "Arity" + n.ToString() + "FunctionOPENINGBRACE"))
                            ((ArgumentsN RangeNoUnion n) .>> (pstring ")" <!> "Arity" + n.ToString() + "FunctionCLOSINGBRACE" ))
                            (fun fname arglist -> ReferenceFunction(us, fname, arglist, Fixed(n)) :> Reference)
                <!> ("Arity" + n.ToString() + "Function")

    let ArityAtLeastNFunction n R =
                // here, we ignore whatever Range context we are given
                // and use RangeNoUnion instead
                getUserState >>=
                fun us ->
                    pipe2
                        (FunctionNamesForAtLeastArity n .>> pstring "(")
                        ((ArgumentsAtLeastN RangeNoUnion n) .>> pstring ")")
                        (fun fname arglist -> ReferenceFunction(us, fname, arglist, LowBound(n)) :> Reference)
                <!> ("Arity" + n.ToString() + "PlusFunction")

    let VarArgsFunction R =
                  // here, we ignore whatever Range context we are given
                  // and use RangeAnyUnion instead (i.e., try both)
                  getUserState >>=
                    fun us ->
                        pipe2
                            (VarArgsFunctionName .>> pstring "(")
                            ((ArgumentList RangeAny) .>> pstring ")")
                            (fun fname arglist -> ReferenceFunction(us, fname, arglist, VarArgs) :> Reference)
                   <!> "VarArgsFunction"

    let Function R: P<Reference> =
        (       
            (   // fixed arity
                arityNNameArr |>
                    Array.mapi (fun i _ -> (attempt_opt (ArityNFunction i R))) |> choice_opt)
            <||> (// low-bounded arity
                arityAtLeastNNameArr |>
                    // note: no "at-least-0" parser; that's just a varargs
                    Array.mapi (fun i _ -> (attempt_opt (ArityAtLeastNFunction (i + 1) R))) |> choice_opt)
            <||> (// varargs
                VarArgsFunction R)
        ) <!> "Function"
    
    do ArgumentListImpl := fun (R: P<Range>) -> sepBy ((ExpressionDecl R) <!> "VarArgs Argument") Comma <!> "ArgumentList"

    // Unary operators
    let UnaryOpChar = (spaces >>. satisfy (fun c -> c = '+' || c = '-') .>> spaces) <!> "UnaryOp"

    // reserved words
    let ReservedWordsReference: P<Reference> =
        (Array.map (fun (rwp: P<string>) -> (attempt_opt rwp)) ReservedWords)
        |> choice_opt
        |>> (fun s ->
                raise (AST.ParseException ("'" + s + "' is a reserved word."))
            )

    // references
    let Reference R = (attempt_opt (RangeReference R))
                        <||> (attempt_opt AddressReference)
                        <||> (attempt_opt BooleanReference)
                        <||> (attempt_opt ConstantReference)
                        <||> (attempt_opt ReservedWordsReference)
                        <||> (attempt_opt StringReference)
                        <||> NamedReference
                      <!> "Reference"

    // Expressions
    let ParensExpr(R: P<Range>): P<Expression> = ((between (pstring "(") (pstring ")") (ExpressionDecl R)) |>> ParensExpr) <!> "ParensExpr"
    let ExpressionAtom(R: P<Range>): P<Expression> = (((attempt_opt (Function R)) <||> (Reference R)) |>> ReferenceExpr) <!> "ExpressionAtom"
    let ExpressionSimple(R: P<Range>): P<Expression> = ((attempt_opt (ExpressionAtom R)) <||> (ParensExpr R)) <!> "ExpressionSimple"
    let UnaryOpExpr(R: P<Range>): P<Expression> = pipe2 UnaryOpChar (ExpressionDecl R) (fun op rhs -> UnaryOpExpr(op, rhs)) <!> "UnaryOpExpr"
   
    // Binary arithmetic operators
    let opp = new OperatorPrecedenceParser<Expression,unit,Env>()
    // ignore ranges for now
    let BinOpExpr(R: P<Range>): P<Expression> = opp.ExpressionParser
    let term = (ExpressionSimple RangeAny) .>> spaces
    opp.TermParser <- term
    opp.AddOperator(InfixOperator("<", spaces, 1, Associativity.Left, (fun lhs rhs -> AST.BinOpExpr("<", lhs, rhs))))
    opp.AddOperator(InfixOperator(">", spaces, 1, Associativity.Left, (fun lhs rhs -> AST.BinOpExpr(">", lhs, rhs))))
    opp.AddOperator(InfixOperator("=", spaces, 1, Associativity.None, (fun lhs rhs -> AST.BinOpExpr("=", lhs, rhs))))
    opp.AddOperator(InfixOperator("<=", spaces, 1, Associativity.None, (fun lhs rhs -> AST.BinOpExpr("<=", lhs, rhs))))
    opp.AddOperator(InfixOperator(">=", spaces, 1, Associativity.None, (fun lhs rhs -> AST.BinOpExpr(">=", lhs, rhs))))
    opp.AddOperator(InfixOperator("<>", spaces, 1, Associativity.None, (fun lhs rhs -> AST.BinOpExpr("<>", lhs, rhs))))
    opp.AddOperator(InfixOperator("&", spaces, 2, Associativity.Left, (fun lhs rhs -> AST.BinOpExpr("&", lhs, rhs))))
    opp.AddOperator(InfixOperator("+", spaces, 3, Associativity.Left, (fun lhs rhs -> AST.BinOpExpr("+", lhs, rhs))))
    opp.AddOperator(InfixOperator("-", spaces, 3, Associativity.Left, (fun lhs rhs -> AST.BinOpExpr("-", lhs, rhs))))
    opp.AddOperator(InfixOperator("*", spaces, 4, Associativity.Left, (fun lhs rhs -> AST.BinOpExpr("*", lhs, rhs))))
    opp.AddOperator(InfixOperator("/", spaces, 4, Associativity.Left, (fun lhs rhs -> AST.BinOpExpr("/", lhs, rhs))))
    opp.AddOperator(InfixOperator("^", spaces, 5, Associativity.Left, (fun lhs rhs -> AST.BinOpExpr("^", lhs, rhs))))

    // Parsing Excel argument lists is ambiguous without knowing
    // the arity of the function that you are parsing.  Thus parsers in this grammar
    // take a Range parser that parses ranges with the appropriate semantics.
    // The appropriate parser is overridden by the Function parser.
    // By default, the grammar tries to parse both (see top-level Formula parser).
    do ExpressionDeclImpl :=
        fun (R: P<Range>) ->
            (
                ((attempt_opt (UnaryOpExpr R))
                <||> (attempt_opt (BinOpExpr R))
                <||> (ExpressionSimple R))
            )
            <!> "Expression"

    // Formulas
    let Formula = (pstring "=" .>> spaces >>. (ExpressionDecl RangeAny) .>> eof) <!> "Formula"
