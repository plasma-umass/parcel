module internal RangeVisitor
    let rec rangesFromRangeRef(ref: AST.ReferenceRange) : AST.Range list = [ref.Range]

    and rangesFromFunctionRef(ref: AST.ReferenceFunction) : AST.Range list =
        if UglyHacks.isIgnored(ref.FunctionName) then
            []
        else
            List.map (fun arg -> rangesFromExpr(arg)) ref.ArgumentList |> List.concat
        
    and rangesFromExpr(expr: AST.Expression) : AST.Range list =
        match expr with
        | AST.ReferenceExpr(r) -> rangesFromRef(r)
        | AST.BinOpExpr(op, e1, e2) -> rangesFromExpr(e1) @ rangesFromExpr(e2)
        | AST.UnaryOpExpr(op, e) -> rangesFromExpr(e)
        | AST.ParensExpr(e) -> rangesFromExpr(e)

    and rangesFromRef(ref: AST.Reference) : AST.Range list =
        match ref with
        | :? AST.ReferenceRange as r -> rangesFromRangeRef(r)
        | :? AST.ReferenceAddress -> []
        | :? AST.ReferenceNamed -> []   // TODO: symbol table lookup
        | :? AST.ReferenceFunction as r -> rangesFromFunctionRef(r)
        | :? AST.ReferenceConstant -> []
        | :? AST.ReferenceString -> []
        | _ -> failwith "Unknown reference type."
