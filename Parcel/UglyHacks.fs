module internal UglyHacks
    let isIgnored(fnname: string) : bool =
        match fnname with
        | "INDEX" -> true
        | "HLOOKUP" -> true
        | "VLOOKUP" -> true
        | "LOOKUP" -> true
        | "OFFSET" -> true
        | _ -> false