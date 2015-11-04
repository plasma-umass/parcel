module internal CellVisitor
    let rec addrsFromExpr(expr: AST.Expression) : AST.Address list =
        match expr with
        | AST.ReferenceExpr(r) -> addrsFromRef(r)
        | AST.BinOpExpr(op, e1, e2) -> addrsFromExpr(e1) @ addrsFromExpr(e2)
        | AST.UnaryOpExpr(op, e) -> addrsFromExpr(e)
        | AST.ParensExpr(e) -> addrsFromExpr(e)

    and addrsFromRef(ref: AST.Reference) : AST.Address list =
        match ref with
        | :? AST.ReferenceRange -> []
        | :? AST.ReferenceAddress as r -> addrsFromAddrRef(r)
        | :? AST.ReferenceNamed -> []   // TODO: symbol table lookup
        | :? AST.ReferenceFunction as r -> addrsFromFunctionRef(r)
        | :? AST.ReferenceConstant -> []
        | :? AST.ReferenceString -> []
        | _ -> failwith "Unknown reference type."

    and addrsFromAddrRef(ref: AST.ReferenceAddress) : AST.Address list = [ref.Address]

    and addrsFromFunctionRef(ref: AST.ReferenceFunction) : AST.Address list =
        List.map (fun arg -> addrsFromExpr(arg)) ref.ArgumentList |> List.concat
