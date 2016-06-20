module Grammar
    open FParsec
    open AST
    open System.Text.RegularExpressions

    let DEBUG_MODE = false

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
        if DEBUG_MODE then
            fun stream ->
                let s = String.replicate (int (stream.Position.Column)) " "
                printfn "%d%sTrying %s(\"%s\")" (stream.Index) s label (stream.PeekString 1000)
                let reply = p stream
                printfn "%d%s%s(\"%s\") (%A)" (stream.Index) s label (stream.PeekString 1000) reply.Status
                reply
        else
            p

    // Grammar forward references
    let (ArgumentList: P<Range> -> P<Expression list>, ArgumentListImpl) = createParserForwardedToRefWithArguments()
    let (ExpressionDecl: P<Range> -> P<Expression>, ExpressionDeclImpl) = createParserForwardedToRefWithArguments()
    let (RangeA1Union: P<Range>, RangeA1UnionImpl) = createParserForwardedToRef()
    let (RangeA1NoUnion: P<Range>, RangeA1NoUnionImpl) = createParserForwardedToRef()
    let (RangeR1C1: P<Range>, RangeR1C1Impl) = createParserForwardedToRef()

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
                        (attempt AddrIndirect)
                        <|> (pipe2
                                AddrR
                                AddrC
                                (fun row col ->
                                    // TODO: R1C1 absolute/relative
                                    Address.fromR1C1withMode(row, col, AddressMode.Relative, AddressMode.Relative, us.WorksheetName, us.WorkbookName, us.Path)))
                    <!> "AddrR1C1"
    let AbsOrNot : P<AddressMode> = (attempt ((pstring "$") >>% Absolute) <|> (pstring "" >>% Relative)) 
    let AddrA = many1Satisfy isAsciiUpper
    let AddrAAbs = pipe2 AbsOrNot AddrA (fun mode col -> (mode, col))
    let Addr1 = pint32
    let Addr1Abs = pipe2 AbsOrNot Addr1 (fun mode row -> (mode, row))
    let AddrA1 = getUserState >>=
                    fun (us: Env) ->
                        (attempt AddrIndirect)
                        <|> (pipe2
                            AddrAAbs
                            Addr1Abs
                            (fun col row ->
                                let (colMode,colA) = col
                                let (rowMode,row1) = row
                                Address.fromA1withMode(row1, colA, rowMode, colMode, us.WorksheetName, us.WorkbookName, us.Path)
                                ))
                    
    let AnyAddr = ((attempt AddrIndirect)
                  <|> (attempt AddrR1C1)
                  <|> AddrA1)
                  <!> "AnyAddr"

    // Ranges
    let RA1TO_1 = pstring ":" >>. AddrA1 .>> pstring ","
    let RA1TO_2 = pstring ":" >>. AddrA1
    let RTO = (attempt RA1TO_1) <|> RA1TO_2

    let RA1TC_1 = pstring "," >>. AddrA1 .>> pstring ","
    let RA1TC_2 = pstring "," >>. AddrA1
    let RTC = (attempt RA1TC_1) <|> RA1TC_2

    let RA1_1 = (* AddrA1 RTC (1)*) pipe2 AddrA1 RTC (fun a1 a2 -> Range([(a2,a2); (a1,a1)]))
    let RA1_2 = (* AddrA1 RTO (2)*) pipe2 AddrA1 RTO (fun a1 a2 -> Range(a1,a2))

    let rat_fun = fun a1 a2 a3 -> Range([(a1,a2);(a3,a3)])
    let RA1_3a = pipe3 AddrA1 RA1TO_1 RA1TC_1 rat_fun
    let RA1_3b = pipe3 AddrA1 RA1TO_1 RA1TC_2 rat_fun
    let RA1_3c = pipe3 AddrA1 RA1TO_2 RA1TC_1 rat_fun
    let RA1_3d = pipe3 AddrA1 RA1TO_2 RA1TC_2 rat_fun
    let RA1_3 = (* AddrA1 RTO RTC (3)*) 
                (attempt RA1_3a) <|> (attempt RA1_3b) <|> (attempt RA1_3c) <|> RA1_3d

    let RA1_4_UNION = (* AddrA1 "," R (4)*) pipe2 (AddrA1 .>> (pstring ",")) RangeA1Union (fun a1 r1 -> Range(r1.Ranges() @ [(a1, a1)]))
    let RA1_5_UNION = (* AddrA1 RTO R (5)*) pipe3 AddrA1 RTO RangeA1Union (fun a1 a2 r2 -> Range ((a1,a2) :: r2.Ranges()))
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
//                                        <!> "RangeReferenceWorkbook"
    let RangeReferenceWorksheet R = getUserState >>=
                                    fun (us: Env) ->
                                        pipe2
                                            (WorksheetName .>> pstring "!")
                                            R
                                            (fun wsname rng -> ReferenceRange(Env(us.Path, us.WorkbookName, wsname), rng) :> Reference)
//                                    <!> "RangeReferenceWorksheet"
    let RangeReferenceNoWorksheet R = getUserState >>=
                                        fun (us: Env) ->
                                            R
                                            |>> (fun rng -> ReferenceRange(us, rng) :> Reference)
//                                    <!> "RangeReferenceNOWorksheet"
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
//                                    <!> "AddressReferenceWorkbook"
    let AddressReferenceWorksheet = getUserState >>=
                                        fun us ->
                                            pipe2
                                                (WorksheetName .>> pstring "!")
                                                AnyAddr
                                                (fun wsname addr -> ReferenceAddress(Env(us.Path, us.WorkbookName, wsname), addr) :> Reference)
//                                    <!> "AddressReferenceWorksheet"
    let AddressReferenceNoWorksheet = getUserState >>=
                                        fun us ->
                                            AnyAddr
                                            |>> (fun addr -> ReferenceAddress(us, addr) :> Reference)
//                                      <!> "AddressReferenceNOWorksheet"
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
    let ArityNFunctionNameMaker n xs = xs |> List.map (fun name -> attempt (pstring name)) |> choice <!> ("Arity" + n.ToString() + "FunctionName")
    let ArityAtLeastNFunctionNameMaker nplus xs = xs |> List.map (fun name -> attempt (pstring name)) |> choice <!> ("Arity" + nplus.ToString() + "+FunctionName")

    let Arity0NoParensFunctionName: P<string> = ArityNFunctionNameMaker 0 ["FALSE"; "TRUE";]

    let Arity0FunctionName: P<string> = ArityNFunctionNameMaker 0 ["COLUMN"; "FALSE"; "NA"; "NOW"; "PI"; "RAND"; "ROW"; "SHEET"; "SHEETS"; "TODAY"; "TRUE"]

    let Arity1FunctionName: P<string> = ArityNFunctionNameMaker 1 ["ABS"; "ACOS"; "ACOSH"; "ACOT"; "ACOTH"; "ARABIC"; "AREAS"; "ASC"; "ASIN"; "ASINH"; "ATAN"; "ATANH";
        "BAHTTEXT"; "BIN2DEC"; "BIN2HEX"; "BIN2OCT"; "CEILING.PRECISE"; "CELL"; "CHAR"; "CLEAN"; "CODE"; "COLUMN"; "COLUMNS"; "COS"; "COSH"; "COT"; "COTH"; "COUNTBLANK";
        "CSC"; "CSCH"; "CUBESETCOUNT"; "DAY"; "DBCS"; "DEC2BIN"; "DEC2HEX"; "DEC2OCT"; "DEGREES"; "DELTA"; "DOLLAR"; "ENCODEURL"; "ERF"; "ERF.PRECISE"; "ERFC"; "ERFC.PRECISE";
        "ERROR.TYPE"; "EVEN"; "EXP"; "FACT"; "FACTDOUBLE"; "FISHER"; "FISHERINV"; "FIXED"; "FLOOR.PRECISE"; "FORMULATEXT"; "GAMMA"; "GAMMALN"; "GAMMALN.PRECISE"; "GAUSS";
        "GESTEP"; "GROWTH"; "HEX2BIN"; "HEX2DEC"; "HEX2OCT"; "HOUR"; "HYPERLINK"; "IMABS"; "IMAGINARY"; "IMARGUMENT"; "IMCONJUGATE"; "IMCOS"; "IMCOSH"; "IMCOT"; "IMCSC";
        "IMCSCH"; "IMEXP"; "IMLN"; "IMLOG10"; "IMLOG2"; "IMREAL"; "IMSEC"; "IMSECH"; "IMSIN"; "IMSINH"; "IMSQRT"; "IMTAN"; "INDEX"; "INDIRECT"; "INFO"; "INT";  "IRR";
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
        "YEAR"]

    let Arity2FunctionName: P<string> = ArityNFunctionNameMaker 2 ["ADDRESS"; "ATAN2"; "AVERAGEIF"; "BASE"; "BESSELI"; "BESSELJ"; "BESSELK"; "BESSELY";
        "BIN2HEX"; "BIN2OCT"; "BITAND"; "BITLSHIFT"; "BITOR"; "BITRSHIFT"; "BITXOR"; "CEILING"; "CEILING.PRECISE"; "CELL"; "CHIDIST"; "CHIINV"; "CHITEST";
        "CHISQ.DIST.RT"; "CHISQ.TEST"; "COMBIN"; "COMBINA"; "COMPLEX"; "COUNTIF"; "COVAR"; "COVARIANCE.P"; "COVARIANCE.S"; "CUBEMEMBER"; "CUBERANKEDMEMBER";
        "CUBESET"; "DATEVALUE"; "DAYS"; "DAYS360"; "DEC2BIN"; "DEC2HEX"; "DEC2OCT"; "DECIMAL"; "DELTA"; "DOLLAR"; "DOLLARDE"; "DOLLARFR"; "EDATE"; "EFFECT";
        "EOMONTH"; "ERF"; "EXACT"; "F.DIST"; "FILTERXML"; "FIND"; "FINDB"; "FIXED"; "FLOOR"; "FLOOR.PRECISE"; "FORECAST.ETS.SEASONALITY"; "FREQUENCY"; "F.TEST";
        "FTEST"; "FVSCHEDULE"; "GESTEP"; "GROWTH"; "HEX2BIN"; "HEX2OCT"; "HYPERLINK"; "IF"; "IFERROR"; "IFNA"; "IMDIV"; "IMPOWER"; "IMSUB"; "INDEX"; "INDIRECT";
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
        "ZTEST"; "Z.TEST"]

    let Arity3FunctionName: P<string> = ArityNFunctionNameMaker 3 ["ADDRESS"; "AVERAGEIF"; "BASE"; "BETADIST"; "BETAINV"; "BETA.INV"; "BINOM.DIST.RANGE"; "BINOM.INV";
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
        "ZTEST"; "Z.TEST"]

    let Arity4FunctionName: P<string> = ArityNFunctionNameMaker 4 ["ADDRESS"; "ACCRINTM"; "BETADIST"; "BETA.DIST"; "BETAINV"; "BETA.INV"; "BINOMDIST"; "BINOM.DIST";
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
        "YIELDDISC"]

    let Arity5FunctionName: P<string> = ArityNFunctionNameMaker 5 ["ADDRESS"; "ACCRINTM"; "BETADIST"; "BETA.DIST"; "BETAINV"; "BETA.INV"; "CUBESET"; "DB"; "DDB";
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
        "YIELDDISC"; "YIELDMAT"]

    let Arity6FunctionName: P<string> = ArityNFunctionNameMaker 6 ["ACCRINT"; "AMORDEGRC"; "AMORLINC"; "BETA.DIST";
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
        "YIELD"; "YIELDMAT"]

    let Arity7FunctionName: P<string> = ArityNFunctionNameMaker 7 ["ACCRINT"; "AMORDEGRC"; "AMORLINC"; "FORECAST.ETS.CONFINT";
        "MEDIAN"; "MULTINOMIAL";
        "NPV"; 
        "ODDLPRICE"; "ODDLYIELD"; "OR";
        "PRICE"; "PRODUCT"; 
        "RTD"; 
        "STDEV"; "STDEV.P"; "STDEV.S"; "STDEVA"; "STDEVP"; "STDEVPA"; "SUBTOTAL"; "SUMIFS"; "SUMPRODUCT"; "SUMSQ"; "SWITCH";
        "TEXTJOIN"; 
        "VAR"; "VAR.P"; "VAR.S"; "VARA"; "VARP"; "VARPA"; "VDB";
        "XOR";
        "YIELD"]

    let Arity8FunctionName: P<string> = ArityNFunctionNameMaker 8 ["ACCRINT";
        "MEDIAN"; "MULTINOMIAL";
        "NPV"; 
        "ODDFPRICE"; "ODDFYIELD"; "ODDLPRICE"; "ODDLYIELD"; "OR";
        "PRODUCT"; 
        "RTD"; 
        "STDEV"; "STDEV.P"; "STDEV.S"; "STDEVA"; "STDEVP"; "STDEVPA"; "SUBTOTAL"; "SUMPRODUCT"; "SUMSQ"; "SWITCH";
        "TEXTJOIN"; 
        "VAR"; "VAR.P"; "VAR.S"; "VARA"; "VARP"; "VARPA";
        "XOR"]

    let Arity9FunctionName: P<string> = ArityNFunctionNameMaker 9 ["ACCRINT";
        "MEDIAN"; "MULTINOMIAL";
        "NPV"; 
        "ODDFPRICE"; "ODDFYIELD"; "OR";
        "PRODUCT"; 
        "RTD";
        "STDEV"; "STDEV.P"; "STDEV.S"; "STDEVA"; "STDEVP"; "STDEVPA"; "SUBTOTAL"; "SUMPRODUCT"; "SUMSQ"; "SWITCH";
        "VAR"; "VAR.P"; "VAR.S"; "VARA"; "VARP"; "VARPA";
        "XOR"]

    let ArityAtLeast1FunctionName: P<string> = ArityAtLeastNFunctionNameMaker 1 ["AND"; "AVEDEV"; "AVERAGE"; "AVERAGEA";
        "CALL"; "CONCAT"; "CONCATENATE"; "COUNT"; "COUNTA"; "CUBEVALUE"; "DEVSQ"; "GCD"; "GEOMEAN"; "HARMEAN"; "IMPRODUCT";
        "IMSUM"; "KURT"; "LCM"; "MAX"; "MAXA"; "MIN"; "MINA"; "MODE"; "MODE.MULT"; "MODE.SNGL"; "MULTINOMIAL"; ]
    let ArityAtLeast2FunctionName: P<string> = ArityAtLeastNFunctionNameMaker 2 ["CHOOSE"; "COUNTIFS"; "GETPIVOTDATA"; "IFS"; ]
    let ArityAtLeast3FunctionName: P<string> = ArityAtLeastNFunctionNameMaker 3 ["AGGREGATE"; "AVERAGEIFS"; "MAXIFS"; "MINIFS"; ]

    let VarArgsFunctionName: P<string> =
        ["SUM"]
        |> List.map (fun name -> pstring name) |> choice <!> "VarArgsFunctionName"

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

    let Argument R = fun (i: int) -> (((ExpressionDecl R) .>> (spaces >>. pstring "," .>> spaces)) <!> ("Argument #" + i.ToString()))

    let ArgumentsN R n =
        (pipe2
            (pmap (Argument R) [1..n-1])
            ((ExpressionDecl R) <!> ("Argument #" + n.ToString() + " (last arg)"))
            (fun exprs expr -> exprs @ [expr] ) <!> ("Arguments" + n.ToString())
        )

    let ArgumentsAtLeastN R n =
        (pipe3
            (parray (n-1) ((ExpressionDecl R) .>> pstring ",") )
            (ExpressionDecl R)
            (ArgumentList R)
            (fun exprArr expr varArgs -> Array.toList exprArr @ [expr] @ varArgs) <!> ("Arguments" + n.ToString() + "+")
        )

    let Arity0NoParensFunction R =
        // Range context is irrelevant here since we
        // have no arguments
        getUserState >>=
        fun us ->
            Arity0FunctionName |>> (fun fname -> ReferenceFunction(us, fname, [], Fixed(0)) :> Reference)
        <!> ("Arity0NoParensFunction")

    let ArityNFunction n R =
                // here, we ignore whatever Range context we are given
                // and use RangeNoUnion instead
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
        (       // arity 0 and no parens
            (attempt (Arity0NoParensFunction R))
            <|> (// fixed arity
                arityNNameArr |>
                    Array.mapi (fun i _ -> (attempt (ArityNFunction i R))) |> choice)
            <|> (// low-bounded arity
                arityAtLeastNNameArr |>
                    // note: no "at-least-0" parser; that's just a varargs
                    Array.mapi (fun i _ -> (attempt (ArityAtLeastNFunction (i + 1) R))) |> choice)
                // varargs
            <|> VarArgsFunction R
        ) <!> "Function"
    
    do ArgumentListImpl := fun (R: P<Range>) -> sepBy ((ExpressionDecl R) <!> "VarArgs Argument") (spaces >>. pstring "," .>> spaces) <!> "ArgumentList"

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
