﻿module AST
    open System

    type ParseException(formula: string, reason: string) =
        inherit System.Exception(formula)
        new(formula: string) = ParseException(formula, "")

    type IndirectAddressingNotSupportedException(expression: string) =
        inherit Exception(expression)

    // The Env object is threaded through grammar combinators (via UserState)
    // so that parsers always have the path, workbook name, and worksheet name
    // from the calling workbook environment in order to construct fully-
    // qualified AST nodes.
    [<Serializable>]
    type Env(path: string, wbname: string, wsname: string) =
        member self.Path = if path.EndsWith @"\" then path else path + @"\"
        member self.WorkbookName = wbname
        member self.WorksheetName = wsname
        override self.GetHashCode() : int =
            path.GetHashCode() |||
            wbname.GetHashCode() |||
            wsname.GetHashCode()
        override self.Equals(obj: obj) : bool =
            let d2 = obj :?> Env
            path = d2.Path &&
            wbname = d2.WorkbookName &&
            wsname = d2.WorksheetName

    [<Serializable>]
    type AddressMode =
    | Absolute
    | Relative

    [<Serializable>]
    type Address(R: int, C: int, Rmode: AddressMode, Cmode: AddressMode, env: Env) =
        let mutable _cachedHC : int option = None

        new(R: int, C: int, env: Env) = Address(R, C, Relative, Relative, env)

        interface IComparable with
            member self.CompareTo(obj) =
                let addr = obj :?> Address
                match env.WorksheetName.CompareTo(addr.WorksheetName) with
                | -1 -> -1
                | 1  -> 1
                | _  ->
                    // this guarantees a consistent total order, but if used in user-facing
                    // contexts, it might not be very intuitive to end-users.
                    let c = Hash.cantorPair (System.Convert.ToUInt32 R) (System.Convert.ToUInt32 R) System.UInt32.MaxValue
                    let c' = Hash.cantorPair (System.Convert.ToUInt32 addr.Row) (System.Convert.ToUInt32 addr.Col) System.UInt32.MaxValue
                    c.CompareTo c'

        static member fromR1C1withMode(R: int, C: int, rmode: AddressMode, cmode: AddressMode, wsname: string, wbname: string, path: string) : Address =
            Address(R, C, rmode, cmode, Env(path, wbname, wsname))
        static member fromA1withMode(row: int, col: string, rmode: AddressMode, cmode: AddressMode, wsname: string, wbname: string, path: string) : Address =
            Address(row, Address.CharColToInt(col), rmode, cmode, Env(path, wbname, wsname))
        member self.copyWithNewEnv(envnew: Env) =
            Address(R, C, Rmode, Cmode, envnew)
        member self.A1Local() : string =
            let cmode =
                match self.ColMode with
                | AddressMode.Absolute -> "$"
                | AddressMode.Relative -> ""
            let rmode =
                match self.RowMode with
                | AddressMode.Absolute -> "$"
                | AddressMode.Relative -> ""
            cmode + Address.IntToColChars(self.X) + rmode + self.Y.ToString()
        member self.A1Path() : string = env.Path
        member self.A1Worksheet() : string = env.WorksheetName
        member self.A1Workbook() : string = env.WorkbookName
        member self.A1FullyQualified() : string =
            "'" + self.A1Path() + "[" + self.A1Workbook() + "]" + self.A1Worksheet() + "'!" + self.A1Local()
        member self.R1C1 =
            let wsstr = env.WorksheetName + "!"
            let wbstr = "[" + env.WorkbookName + "]"
            let pstr = env.Path
            pstr + wbstr + wsstr + "R" + R.ToString() + "C" + C.ToString()
        member self.XMode = Cmode
        member self.YMode = Rmode
        member self.X: int = C
        member self.Y: int = R
        member self.RowMode = Rmode
        member self.ColMode = Cmode
        member self.Row = R
        member self.Col = C
        member self.Path = env.Path
        member self.WorksheetName = env.WorksheetName
        member self.WorkbookName = env.WorkbookName
        static member addressHash(R: int)(C: int)(sheetname: string) : int =
            // 2^m = w = expected maximum number of worksheets
            // the following is Cantor's pairing function mod w;
            // for signed int, that gives us 2^(32 - m) cells per
            // sheet without collisions.
            assert(R >= 0)
            assert(C >= 0)
            let m = 4u
            let w = 2u <<< int32 (m - 1u) // there are rarely more than 16 worksheets
            let r = (2u <<< int32 (32u - m - 1u)) // the total number of distinct integers allowed from pi
            let k_1 = System.Convert.ToUInt32(R)
            let k_2 = System.Convert.ToUInt32(C)
            // get uint from pairing function
            let pi = Hash.cantorPair k_1 k_2 r
            // shift pi depending on the worksheet hashcode
            // (we don't have access to the real index here; getting the
            //  sheetname's hashcode mod w is an approximation)
            let sheet_hc = uint32 (sheetname.GetHashCode()) % w
            let hashcode = pi + sheet_hc * r
            int32 hashcode // this casts; it does not convert
        // necessary because Address is used as a Dictionary key
        override self.GetHashCode() : int =
            match _cachedHC with
            | Some(hc) -> hc
            | None ->
                let hc = Address.addressHash R C self.WorksheetName
                _cachedHC <- Some(hc)
                hc
        override self.Equals(obj: obj) : bool =
            let addr = obj :?> Address
            self.SameAs addr
        member self.SameAs(addr: Address) : bool =
            // odd construction is for breakpoint-friendliness
            let a = self.X = addr.X
            let b = self.Y = addr.Y
            let c = self.WorksheetName = addr.WorksheetName
            let d = self.WorkbookName = addr.WorkbookName
            a && b && c && d
        override self.ToString() =
            "(" + self.X.ToString() + "," + self.Y.ToString() + ")"
        member self.ToFormula = self.A1FullyQualified()
        static member CharColToInt(col: string) : int =
            let rec ccti(idx: int) : int =
                let ltr = (int col.[idx]) - 64
                let num = (int (Math.Pow(26.0, float (col.Length - idx - 1)))) * ltr
                if idx = 0 then
                    num
                else
                    num + ccti(idx - 1)
            ccti(col.Length - 1)
        static member FromA1String(addr: string, wsname: string, wbname: string, path: string) : Address =
            let reg = System.Text.RegularExpressions.Regex("(?<cmode>\$?)(?<column>[A-Z]+)(?<rmode>\$?)(?<row>[0-9]+)")
            let m = reg.Match(addr)
            let r = System.Convert.ToInt32(m.Groups.["row"].Value)
            let c = Address.CharColToInt(m.Groups.["column"].Value)
            let rmode = if String.IsNullOrEmpty(m.Groups.["rmode"].Value) then AddressMode.Relative else AddressMode.Absolute
            let cmode = if String.IsNullOrEmpty(m.Groups.["cmode"].Value) then AddressMode.Relative else AddressMode.Absolute
            Address.fromR1C1withMode(r, c, rmode, cmode, wsname, wbname, path)
        static member FromA1StringForceMode(addr: string, rmode: AddressMode, cmode: AddressMode, wsname: string, wbname: string, path: string) : Address =
            let reg = System.Text.RegularExpressions.Regex("\$?(?<column>[A-Z]+)\$?(?<row>[0-9]+)")
            let m = reg.Match(addr)
            let r = System.Convert.ToInt32(m.Groups.["row"].Value)
            let c = Address.CharColToInt(m.Groups.["column"].Value)
            Address.fromR1C1withMode(r, c, rmode, cmode, wsname, wbname, path)
        static member FromString(addr: string, wsname: string, wbname: string, path: string) : Address =
            let reg = System.Text.RegularExpressions.Regex("R(?<row>[0-9]+)C(?<column>[0-9]+)")
            let m = reg.Match(addr)
            let r = System.Convert.ToInt32(m.Groups.["row"].Value)
            let c = System.Convert.ToInt32(m.Groups.["column"].Value)
            Address.fromR1C1withMode(r, c, AddressMode.Relative, AddressMode.Relative, wsname, wbname, path)
        static member IntToColChars(dividend: int) : string =
            let mutable quot = dividend / 26
            let rem = dividend % 26
            if rem = 0 then
                quot <- quot - 1
            let ltr = if rem = 0 then
                        'Z'
                      else
                        char (64 + rem)
            if quot = 0 then
                ltr.ToString()
            else
                Address.IntToColChars(quot) + ltr.ToString()

    and IndirectAddress(expr: string, env: Env) =
        inherit Address(0,0,env)
        do
            // indirect references are essentially lambdas for
            // constructing addresses
            raise(IndirectAddressingNotSupportedException(expr))

    // note that regions may overlap; be careful not to double-count cells!
    // Default constructor -- each address-address pair corresponds to top-left
    // and bottom right of a region
    and Range(regions: (Address * Address) list) =
        let _regions = regions
        new(regions: (Address * Address)[]) = Range(List.ofArray regions)
        new(range1: Range, range2: Range) = Range(range1.Ranges() @ range2.Ranges())
     
        new(addr1: Address, addr2: Address) = Range([(addr1,addr2)])
        override self.ToString() =
            let sregs = List.map (fun (tl: Address, br: Address) ->
                tl.ToString() + ":" + br.ToString()) _regions
            "List(" + String.Join(",", sregs) + ")"
        member self.copyWithNewEnv(envnew: Env) =
            Range(List.map (fun (tl: Address, br: Address) ->
                    tl.copyWithNewEnv(envnew), br.copyWithNewEnv(envnew)) _regions
                 )
        member self.A1Local() : string =
            let sregs = List.map (fun (tl: Address, br: Address) ->
                            tl.A1Local() + ":" + br.A1Local()
                        ) _regions
            String.Join(",", sregs)
        member self.GetWorksheetName() : string =
            let names = List.fold (fun wss (tl: Address, br: Address) ->
                                    tl.WorksheetName :: br.WorksheetName :: wss
                                  ) [] _regions |>
                        List.toSeq |>
                        Seq.distinct
            assert(Seq.length names = 1)
            Seq.head names
        member self.GetWorkbookName() : string =
            // I am fairly certain that cross-workbook / cross-worksheet
            // ranges are not possible; just return the first
            // path but check that this is true
            let names = List.fold (fun wss (tl: Address, br: Address) ->
                                    tl.WorkbookName :: br.WorkbookName :: wss
                                  ) [] _regions |>
                        List.toSeq |>
                        Seq.distinct
            assert(Seq.length names = 1)
            Seq.head names
        member self.GetPathName() : string =
            // I am fairly certain that cross-workbook / cross-worksheet
            // ranges are not possible; just return the first
            // path but check that this is true
            let paths = List.fold (fun wss (tl: Address, br: Address) ->
                                    tl.Path :: br.Path :: wss
                                  ) [] _regions |>
                        List.toSeq |>
                        Seq.distinct
            assert(Seq.length paths = 1)
            Seq.head paths
        member self.Ranges() : (Address*Address) list = _regions
        // regions may overlap; thus we only return distinct cell addresses
        member self.Addresses() : Address[] =
            // for every contigious region
            let xs =
                List.map (fun (tl: Address, br: Address) ->
                    // for every column in that region
                    Array.map (fun c ->
                        // and every row in that region
                        Array.map (fun r ->
                            // get the address of the cell contained
                            Address.fromR1C1withMode(r, c, tl.RowMode, tl.ColMode, tl.WorksheetName, tl.WorkbookName, tl.Path)
                        ) [|tl.Y..br.Y|] |>
                        Array.toList
                    ) [|tl.X..br.X|] |>
                    Array.toList |>
                    List.concat
                ) _regions |>
                List.concat
            
            // only enumerate overlapping cells once
            let dist = xs |> List.distinct
            let xsa = dist |> Array.ofList

            // sort by leftmost-topmost
            let sorted =
                xsa |>
                Array.sortWith (fun a b ->
                    if a.Row < b.Row then
                        -1
                    else if a.Row = b.Row && a.Col < b.Col then
                        -1
                    else if a.Row = b.Row && a.Col = b.Col then
                        0
                    else
                        1
                   )
            sorted

        static member addrsInRegion(lt: int*int)(rb: int*int) : (int*int)[] =
            let lt_x = fst lt
            let lt_y = snd lt
            let rb_x = fst rb
            let rb_y = snd rb

            // for every contigious region
            // for every column in that region
            Array.map (fun c ->
                // and every row in that region
                Array.map (fun r -> (c, r)
                ) [|lt_y..rb_y|]
            ) [|lt_x..rb_x|] |>
            Array.concat
            
        override self.GetHashCode() : int =
            Hash.jenkinsOneAtATimeHash _regions
        override self.Equals(obj: obj) : bool =
            let r = obj :?> Range
            let r_set = Set.ofArray(r.Addresses())
            let self_set = Set.ofArray(self.Addresses())
            self_set = r_set
        member self.Equals2(obj: obj) : bool =
            let r = obj :?> Range
            failwith "no"
        member self.ToFormula : string =
            let subrngs = List.map (fun (a1: Address,a2: Address) ->
                              a1.A1FullyQualified() + ":" + a2.A1Local()
                          ) _regions
            String.Join(",", subrngs)

    type ReferenceType =
    | ReferenceAddress  = 0
    | ReferenceRange    = 1
    | ReferenceFunction = 2
    | ReferenceConstant = 3
    | ReferenceString   = 4
    | ReferenceNamed    = 5
    | ReferenceBoolean  = 6

    type Arity =
    | Fixed of int
    | LowBound of int
    | VarArgs

    [<AbstractClass>]
    type Reference(env: Env) =
        abstract member Type: ReferenceType
        abstract member ToFormula: string
        member self.Path = env.Path
        member self.WorkbookName = env.WorkbookName
        member self.WorksheetName = env.WorksheetName
        member self.Environment = env

    and ReferenceRange(env: Env, rng: Range) =
        inherit Reference(env)
        let _rng = rng.copyWithNewEnv(env)

        override self.Type = ReferenceType.ReferenceRange
        override self.ToString() =
            "ReferenceRange(" + env.Path + ",[" + env.WorkbookName + "]," + env.WorksheetName + "," + _rng.ToString() + ")"
        member self.Range = _rng
        override self.Equals(obj: obj) : bool =
            let rr = obj :?> ReferenceRange
            self.Path = rr.Path &&
            self.WorkbookName = rr.WorkbookName &&
            self.WorksheetName = rr.WorksheetName &&
            self.Range = rr.Range
        override self.GetHashCode() : int =
            env.GetHashCode() ||| rng.GetHashCode()
        override self.ToFormula : string = rng.ToFormula

    and ReferenceAddress(env: Env, addr: Address) =
        inherit Reference(env)
        let _addr = addr.copyWithNewEnv(env)
        override self.Type = ReferenceType.ReferenceAddress
        override self.ToString() =
            "ReferenceAddress(" + env.Path + ",[" + env.WorkbookName + "]," + env.WorksheetName + "," + _addr.ToString() + ")"
        member self.Address = _addr
        override self.Equals(obj: obj) : bool =
            let ra = obj :?> ReferenceAddress
            self.Path = ra.Path &&
            self.WorkbookName = ra.WorkbookName &&
            self.WorksheetName = ra.WorksheetName &&
            self.Address = ra.Address
        override self.GetHashCode() : int =
            env.GetHashCode() ||| addr.GetHashCode()
        override self.ToFormula = addr.ToFormula

    and ReferenceFunction(env: Env, fnname: string, arglist: Expression list, arity: Arity) =
        inherit Reference(env)
        override self.Type = ReferenceType.ReferenceFunction
        member self.ArgumentList = arglist
        member self.FunctionName = fnname.ToUpper()
        member self.Arity = arity
        override self.ToString() =
            match arity with
            | Fixed a -> self.FunctionName + "[function" + a.ToString() + "](" + String.Join(",", (List.map (fun arg -> arg.ToString()) arglist)) + ")"
            | LowBound aplus -> self.FunctionName + "[function" + aplus.ToString() + "+](" + String.Join(",", (List.map (fun arg -> arg.ToString()) arglist)) + ")"
            | VarArgs -> self.FunctionName + "[functionVarArgs](" + String.Join(",", (List.map (fun arg -> arg.ToString()) arglist)) + ")"
        override self.Equals(obj: obj) : bool =
            let rf = obj :?> ReferenceFunction
            let arglists = List.zip arglist rf.ArgumentList
            self.Path = rf.Path &&
            self.WorkbookName = rf.WorkbookName &&
            self.WorksheetName = rf.WorksheetName &&
            self.FunctionName = rf.FunctionName &&
            // and recursively compare arglists
            List.fold (fun acc (ethis, ethat) -> acc && ethis = ethat) true arglists
        override self.GetHashCode() : int =
            env.GetHashCode() ||| fnname.GetHashCode() ||| arglist.GetHashCode()
        override self.ToFormula = self.FunctionName +
                                  "(" + String.Join(",", (List.map (fun (arg: Expression) -> arg.ToFormula) arglist)) + ")"

    and ReferenceConstant(env: Env, value: double) =
        inherit Reference(env)
        override self.Type = ReferenceType.ReferenceConstant
        member self.Value = value
        override self.ToString() = "Constant(" + value.ToString() + ")"
        override self.Equals(obj: obj) : bool =
            let rc = obj :?> ReferenceConstant
            self.Path = rc.Path &&
            self.WorkbookName = rc.WorkbookName &&
            self.WorksheetName = rc.WorksheetName &&
            self.Value = rc.Value
        override self.GetHashCode() : int =
            env.GetHashCode() ||| value.GetHashCode()
        override self.ToFormula = value.ToString()

    and ReferenceBoolean(env: Env, value: bool) =
        inherit Reference(env)
        override self.Type = ReferenceType.ReferenceBoolean
        member self.Value = value
        override self.ToString() = "Boolean(" + value.ToString() + ")"
        override self.Equals(obj: obj) : bool =
            let rc = obj :?> ReferenceBoolean
            self.Path = rc.Path &&
            self.WorkbookName = rc.WorkbookName &&
            self.WorksheetName = rc.WorksheetName &&
            self.Value = rc.Value
        override self.GetHashCode() : int =
            env.GetHashCode() ||| value.GetHashCode()
        override self.ToFormula = value.ToString()

    and ReferenceString(env: Env, value: string) =
        inherit Reference(env)
        override self.Type = ReferenceType.ReferenceString
        member self.Value = value
        override self.ToString() = "String(" + value + ")"
        override self.Equals(obj: obj) : bool =
            let rs = obj :?> ReferenceString
            self.Path = rs.Path &&
            self.WorkbookName = rs.WorkbookName &&
            self.WorksheetName = rs.WorksheetName &&
            self.Value = rs.Value
        override self.GetHashCode() : int =
            env.GetHashCode() ||| value.GetHashCode()
        override self.ToFormula = value

    and ReferenceNamed(env: Env, varname: string) =
        inherit Reference(env)
        override self.Type = ReferenceType.ReferenceNamed
        member self.Name = varname
        override self.ToString() = "ReferenceName(" + varname + ")"
        override self.Equals(obj: obj) : bool =
            let rn = obj :?> ReferenceNamed
            self.Path = rn.Path &&
            self.WorkbookName = rn.WorkbookName &&
            self.WorksheetName = rn.WorksheetName &&
            self.Name = rn.Name
        override self.GetHashCode() : int =
            env.GetHashCode() ||| varname.GetHashCode()
        override self.ToFormula = varname

    and ReferenceUnion(env: Env, refs: Expression list) =
        inherit Reference(env)
        override self.Type = ReferenceType.ReferenceNamed
        member self.Name = "union"
        member self.References = refs
        override self.ToString() = "ReferenceUnion(" + String.Join(",", refs) + ")"
        override self.Equals(obj: obj) : bool =
            match obj with
            | :? ReferenceUnion as ru ->
                List.zip refs ru.References
                |> List.fold (fun acc (r1,r2) ->
                      acc && r1 = r2
                   ) true
            | _ -> false
        override self.GetHashCode() : int =
            refs |> List.fold (fun acc r -> acc ||| r.GetHashCode()) (env.GetHashCode())
        override self.ToFormula = "(" + String.Join(",", refs) + ")"

    // TODO: implement .Equals!
    and Expression =
    | ReferenceExpr of Reference
    | BinOpExpr of string * Expression * Expression
    | UnaryOpExpr of char * Expression
    | ParensExpr of Expression
        override self.ToString() =
            match self with
            | ReferenceExpr(r) -> "ReferenceExpr(" + r.ToString() + ")"
            | BinOpExpr(op,e1,e2) -> "BinOpExpr(" + op.ToString() + "," + e1.ToString() + "," + e2.ToString() + ")"
            | UnaryOpExpr(op, e) -> "UnaryOpExpr(" + op.ToString() + "," + e.ToString() + ")"
            | ParensExpr(e) -> "ParensExpr(" + e.ToString() + ")"
        member self.ToFormula =
            match self with
            | ReferenceExpr(r) -> r.ToFormula
            | BinOpExpr(op,e1,e2) -> e1.ToFormula + op + e2.ToFormula
            | UnaryOpExpr(op, e) -> op.ToString() + e.ToFormula
            | ParensExpr(e) -> "(" + e.ToFormula + ")"
        member self.WellFormedFormula = "=" + self.ToFormula