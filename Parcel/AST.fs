module AST
    open System
    open System.Diagnostics
    open System.Collections.Generic

    type IndirectAddressingNotSupportedException(expression: string) =
        inherit Exception(expression)

    // The Env object is threaded through grammar combinators (via UserState)
    // so that parsers always have the path, workbook name, and worksheet name
    // from the calling workbook environment in order to construct fully-
    // qualified AST nodes.
    [<Serializable>]
    type Env(path: string, wbname: string, wsname: string) =
        member self.Path = path
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
    type Address(R: int, C: int, env: Env) =
        static member fromR1C1(R: int, C: int, wsname: string, wbname: string, path: string) : Address =
            Address(R, C, Env(path, wbname, wsname))
        static member fromA1(row: int, col: string, wsname: string, wbname: string, path: string) : Address =
            Address(row, Address.CharColToInt(col), Env(path, wbname, wsname))
        member self.copyWithNewEnv(envnew: Env) =
            Address(R, C, envnew)
        member self.A1Local() : string = Address.IntToColChars(self.X) + self.Y.ToString()
        member self.A1Path() : string = env.Path
        member self.A1Worksheet() : string = env.WorksheetName
        member self.A1Workbook() : string = env.WorkbookName
        member self.A1FullyQualified() : string =
            "[" + self.A1Workbook() + "]" + self.A1Worksheet() + "!" + self.A1Local()
        member self.R1C1 =
            let wsstr = env.WorksheetName + "!"
            let wbstr = "[" + env.WorkbookName + "]"
            let pstr = env.Path
            pstr + wbstr + wsstr + "R" + R.ToString() + "C" + C.ToString()
        member self.X: int = C
        member self.Y: int = R
        member self.Row = R
        member self.Col = C
        member self.Path = env.Path
        member self.WorksheetName = env.WorksheetName
        member self.WorkbookName = env.WorkbookName
        // Address is used as a Dictionary key, and reference equality
        // does not suffice, therefore GetHashCode and Equals are provided
        override self.GetHashCode() : int = String.Intern(self.A1FullyQualified()).GetHashCode()
        override self.Equals(obj: obj) : bool =
            let addr = obj :?> Address
            self.SameAs addr
        member self.SameAs(addr: Address) : bool =
            // odd construction is for breakpoint-friendliness
            let a = self.X = addr.X
            let b = self.Y = addr.Y
            let c = self.WorksheetName = addr.WorksheetName
            let d = self.WorkbookName = addr.WorkbookName
            let e = self.Path = self.Path
            a && b && c && d && e
        member self.InsideRange(rng: Range) : bool =
            not (self.X < rng.getXLeft() ||
                 self.Y < rng.getYTop() ||
                 self.X > rng.getXRight() ||
                 self.Y > rng.getYBottom())
        member self.InsideAddr(addr: Address) : bool =
            self.X = addr.X && self.Y = addr.Y
        override self.ToString() =
            "(" + self.Y.ToString() + "," + self.X.ToString() + ")"
        static member CharColToInt(col: string) : int =
            let rec ccti(idx: int) : int =
                let ltr = (int col.[idx]) - 64
                let num = (int (Math.Pow(26.0, float (col.Length - idx - 1)))) * ltr
                if idx = 0 then
                    num
                else
                    num + ccti(idx - 1)
            ccti(col.Length - 1)
        static member FromString(addr: string, wsname: string, wbname: string, path: string) : Address =
            let reg = System.Text.RegularExpressions.Regex("R(?<row>[0-9]+)C(?<column>[0-9]+)")
            let m = reg.Match(addr)
            let r = System.Convert.ToInt32(m.Groups.["row"].Value)
            let c = System.Convert.ToInt32(m.Groups.["column"].Value)
            Address.fromR1C1(r, c, wsname, wbname, path)
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
        
    and Range(topleft: Address, bottomright: Address) =
        let _tl = topleft
        let _br = bottomright
        override self.ToString() =
            let tlstr = topleft.ToString()
            let brstr = bottomright.ToString()
            tlstr + "," + brstr
        member self.copyWithNewEnv(envnew: Env) =
            Range(_tl.copyWithNewEnv(envnew), _br.copyWithNewEnv(envnew))
        member self.TopLeft = _tl
        member self.BottomRight = _br
        member self.A1Local() : string =
            _tl.A1Local() + ":" + _br.A1Local()
        member self.getUniqueID() : string =
            topleft.ToString() + "," + bottomright.ToString()
        member self.getXLeft() : int = _tl.X
        member self.getXRight() : int = _br.X
        member self.getYTop() : int = _tl.Y
        member self.getYBottom() : int = _br.Y
        member self.InsideRange(rng: Range) : bool =
            not (self.getXLeft() < rng.getXLeft() ||
                 self.getYTop() < rng.getYTop() ||
                 self.getXRight() > rng.getXRight() ||
                 self.getYBottom() > rng.getYBottom())
        // Yup, weird case.  This is because we actually
        // distinguish between addresses and ranges, unlike Excel.
        member self.InsideAddr(addr: Address) : bool =
            not (self.getXLeft() < addr.X ||
                 self.getYTop() < addr.Y ||
                 self.getXRight() > addr.X ||
                 self.getYBottom() > addr.Y)
        member self.GetWorksheetNames() : seq<string> =
            [_tl.WorksheetName; _br.WorksheetName] |> List.toSeq |> Seq.distinct
        member self.GetWorkbookNames() : seq<string> =
            [_tl.WorkbookName; _br.WorkbookName] |> List.toSeq |> Seq.distinct
        member self.GetPathNames() : seq<string> =
            [_tl.Path; _br.Path] |> List.toSeq
        member self.Addresses() : Address[] =
            Array.map (fun c ->
                Array.map (fun r ->
                    Address.fromR1C1(r, c, _tl.WorksheetName, _tl.WorkbookName, _tl.Path)
                ) [|self.getYTop()..self.getYBottom()|]
            ) [|self.getXLeft()..self.getXRight()|] |>
            Array.concat
        override self.Equals(obj: obj) : bool =
            let r = obj :?> Range
            topleft = r.TopLeft &&
            bottomright = r.BottomRight

    type ReferenceType =
    | ReferenceAddress  = 0
    | ReferenceRange    = 1
    | ReferenceFunction = 2
    | ReferenceConstant = 3
    | ReferenceString   = 4
    | ReferenceNamed    = 5

    [<AbstractClass>]
    type Reference(env: Env) =
        abstract member InsideRef: Reference -> bool
        abstract member Type: ReferenceType
        member self.Path = env.Path
        member self.WorkbookName = env.WorkbookName
        member self.WorksheetName = env.WorksheetName
        default self.InsideRef(ref: Reference) = false

    and ReferenceRange(env: Env, rng: Range) =
        inherit Reference(env)
        let _rng = rng.copyWithNewEnv(env)

        override self.Type = ReferenceType.ReferenceRange
        override self.ToString() =
            "ReferenceRange(" + env.Path + ",[" + env.WorkbookName + "]," + env.WorksheetName + "," + _rng.ToString() + ")"
        override self.InsideRef(ref: Reference) : bool =
            match ref with
            | :? ReferenceAddress as ar -> _rng.InsideAddr(ar.Address)
            | :? ReferenceRange as rr -> _rng.InsideRange(rr.Range)
            | _ -> failwith "Unknown Reference subclass."
        member self.Range = _rng
        override self.Equals(obj: obj) : bool =
            let rr = obj :?> ReferenceRange
            self.Path = rr.Path &&
            self.WorkbookName = rr.WorkbookName &&
            self.WorksheetName = rr.WorksheetName &&
            self.Range = rr.Range

    and ReferenceAddress(env: Env, addr: Address) =
        inherit Reference(env)
        let _addr = addr.copyWithNewEnv(env)
        override self.Type = ReferenceType.ReferenceAddress
        override self.ToString() =
            "ReferenceAddress(" + env.Path + ",[" + env.WorkbookName + "]," + env.WorksheetName + "," + _addr.ToString() + ")"
        member self.Address = _addr
        override self.InsideRef(ref: Reference) =
            match ref with
            | :? ReferenceAddress as ar -> _addr.InsideAddr(ar.Address)
            | :? ReferenceRange as rr -> _addr.InsideRange(rr.Range)
            | _ -> failwith "Invalid Reference subclass."
        override self.Equals(obj: obj) : bool =
            let ra = obj :?> ReferenceAddress
            self.Path = ra.Path &&
            self.WorkbookName = ra.WorkbookName &&
            self.WorksheetName = ra.WorksheetName &&
            self.Address = ra.Address

    and ReferenceFunction(env: Env, fnname: string, arglist: Expression list) =
        inherit Reference(env)
        override self.Type = ReferenceType.ReferenceFunction
        member self.ArgumentList = arglist
        member self.FunctionName = fnname.ToUpper()
        override self.ToString() =
            self.FunctionName + "[function](" + String.Join(",", (List.map (fun arg -> arg.ToString()) arglist)) + ")"
        override self.Equals(obj: obj) : bool =
            let rf = obj :?> ReferenceFunction
            self.Path = rf.Path &&
            self.WorkbookName = rf.WorkbookName &&
            self.WorksheetName = rf.WorksheetName &&
            self.FunctionName = rf.FunctionName
            // TODO: should also check ArgumentList here!

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