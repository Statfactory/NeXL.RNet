namespace NeXL.RNet
open NeXL.ManagedXll
open NeXL.XlInterop
open System
open System.IO
open System.Runtime.InteropServices
open System.Data
open RProvider
open Microsoft.VisualBasic
open RDotNet
open RProvider.``base``
open RProvider.stats
open RProvider.utils
open RProvider.datasets
open RProvider.Helpers

[<XlQualifiedName(true)>]
module RNet =

    type internal SymbolicExpression with  
        member internal this.Description =
            match this with 
                | CharacterVector(_) -> "CharacterVector"
                | ComplexVector(_) -> "ComplexVector"
                | IntegerVector(_) -> "IntegerVector"
                | LogicalVector(_) -> "LogicalVector"
                | NumericVector(_) -> "NumericVector"
                | RawVector(_) -> "RawVector"
                | UntypedVector(_) -> "UntypedVector"
                | BuiltinFunction(_) -> "BuiltinFunction"
                | Closure(_) -> "Closure"
                | SpecialFunction(_) -> "SpecialFunction"
                | Function(_) -> "Function"
                | Environment(_) -> "Environment"
                | Expression(_) -> "Expression"
                | Language(_) -> "Language"
                | List(_) -> "List"
                | Pairlist(_) -> "Pairlist"
                | Null(_) -> "Null"
                | Symbol(_) -> "Symbol"
                | Factor(_) -> "Factor"
                | CharacterMatrix(_) -> "CharacterMatrix"
                | ComplexMatrix(_) -> "ComplexMatrix"
                | IntegerMatrix(_) -> "IntegerMatrix"
                | LogicalMatrix(_) -> "LogicalMatrix"
                | NumericMatrix(_) -> "NumericMatrix"
                | RawMatrix(_) -> "RawMatrix"
                | DataFrame(_) -> "DataFrame"
                | _ -> "Unknown type"

    let internal toColMajorArray (v : 'T[,]) =
        let nrows = v.GetLength(0)
        Array.init v.Length (fun i -> v.[i % nrows, i / nrows])

    let internal allBool (v : XlValue[]) =
        v |> Array.map (fun x -> match x with | XlBool(_) -> true | _ -> false) |> Array.reduce (&&)

    let internal allNumericOrMiss (v : XlValue[]) =
        v |> Array.map (fun x -> match x with | XlNumeric(_) | XlNil | XlMissing -> true | _ -> false) |> Array.reduce (&&)

    let internal allCharacterOrMiss (v : XlValue[]) =
        v |> Array.map (fun x -> match x with | XlString(_) | XlNil | XlMissing -> true | _ -> false) |> Array.reduce (&&)

    [<XlConverter>]
    let convToSymExpr (v : XlData) =
        match v with   
            | XlArray2D(arr) ->
                let isVector = arr.GetLength(0) = 1 || arr.GetLength(1) = 1
                let v = arr |> toColMajorArray
                if v |> allBool then
                    let v = v |> Array.map (fun x -> match x with | XlBool(b) -> b | _ -> raise (new InvalidOperationException()))
                    if isVector then
                        R.c(v)
                    else
                        R.matrix(data = v, nrow = arr.GetLength(0), ncol = arr.GetLength(1), byrow = false)
                elif v |> allCharacterOrMiss then
                    let v = v |> Array.map (fun x -> match x with | XlString(s) -> s | _ -> String.Empty)
                    if isVector then
                        R.c(v)
                    else
                        R.matrix(data = v, nrow = arr.GetLength(0), ncol = arr.GetLength(1), byrow = false)
                elif v |> allNumericOrMiss then
                    let v = v |> Array.map (fun x -> match x with | XlNumeric(x) -> x | _ -> Double.NaN)
                    if isVector then
                        R.c(v)
                    else
                        R.matrix(data = v, nrow = arr.GetLength(0), ncol = arr.GetLength(1), byrow = false)
                else
                    raise (new ArgumentException("Cannot convert to SymbolicExpr"))

    let internal frameToDataTable (frame : DataFrame) : DataTable =
        let dbTable = new DataTable()
        let rowNames = frame.RowNames
        let rowNamesInt = rowNames |> Array.map (Int32.TryParse >> fst) |> Array.reduce (&&)
        let rowNameCol =
            if rowNamesInt then
                new DataColumn(" ", typeof<int>)
            else
                new DataColumn(" ", typeof<string>)

        dbTable.Columns.Add(rowNameCol)

        let cols = frame.ColumnNames |> Array.map (fun colName -> 
                                                       match frame.AsList().[colName] with  
                                                           | IntegerVector(v) -> new DataColumn(colName, typeof<int>)
                                                           | NumericVector(v) -> new DataColumn(colName, typeof<float>)
                                                           | LogicalVector(v) -> new DataColumn(colName, typeof<bool>)
                                                           | CharacterVector(v) -> new DataColumn(colName, typeof<string>)
                                                           | _ -> raise (new InvalidOperationException())
                                                  ) 
        dbTable.Columns.AddRange(cols)
        let rowCount = frame.RowCount
        [0.. rowCount - 1] |> List.iter (fun i -> 
                                             let row = dbTable.NewRow()
                                             if rowNamesInt then
                                                 row.[0] <- rowNames.[i] |> Int32.TryParse |> snd
                                             else
                                                 row.[0] <- rowNames.[i]
                                             dbTable.Rows.Add(row)
                                        )

        frame.ColumnNames |> Array.iter (fun colName ->
                                            match frame.AsList().[colName] with  
                                                | IntegerVector(v) -> 
                                                    v.GetValue<int[]>() |> Array.iteri (fun i x ->
                                                                                            if x = Int32.MinValue then
                                                                                                dbTable.Rows.[i].[colName] <- DBNull.Value
                                                                                            else
                                                                                                dbTable.Rows.[i].[colName] <- x
                                                                                       )
                                                | NumericVector(v) -> 
                                                    v.GetValue<float[]>() |> Array.iteri (fun i x -> dbTable.Rows.[i].[colName] <- x)
                                                | LogicalVector(v) -> 
                                                    v.GetValue<bool[]>() |> Array.iteri (fun i x -> dbTable.Rows.[i].[colName] <- x)
                                                | CharacterVector(v) ->
                                                    v.GetValue<string[]>() |> Array.iteri (fun i x -> dbTable.Rows.[i].[colName] <- x)
                                                | _ -> raise (new InvalidOperationException())     
                                        )
        dbTable

    let internal matrixToXlTable (conv : 'T -> XlValue) (data : 'T[,]) (rowNames : string[]) (colNames : string[]) =
        if rowNames <> null && rowNames.Length = data.GetLength(0) then
            if colNames <> null && colNames.Length = data.GetLength(1) then
                let xlCols = 
                    [|0..colNames.Length|] |> Array.map (fun i -> if i = 0 then {Name = ""; IsDateTime = false} else {Name = colNames.[i - 1]; IsDateTime = false})
                let xlData = 
                    Array2D.init (rowNames.Length) (colNames.Length + 1) (fun i j -> 
                                                                                if j = 0 then
                                                                                    rowNames.[i] |> XlString
                                                                                else
                                                                                    data.[i, j - 1] |> conv
                                                                            )
                new XlTable(xlCols, xlData, "", "", false, false, true)
            else
                let xlData = 
                    Array2D.init (rowNames.Length) (data.GetLength(1) + 1) (fun i j -> 
                                                                                if j = 0 then
                                                                                    rowNames.[i] |> XlString
                                                                                else
                                                                                    data.[i, j - 1] |> conv
                                                                            )
                new XlTable([||], xlData, "", "", false, false, false)
        else
            if colNames <> null && colNames.Length = data.GetLength(1) then
                let xlCols = colNames |> Array.map (fun name -> {Name = name; IsDateTime = false})
                let xlData = data |> Array2D.map conv
                new XlTable(xlCols, xlData, "", "", false, false, true)
            else
                let xlData = data |> Array2D.map conv
                new XlTable([||], xlData, "", "", false, false, false)

    let internal vectorToXlTable (conv : 'T -> XlValue) (data : 'T[]) (names : string[]) =
        let xlData = 
            if names <> null && names.Length = data.Length then
                Array2D.init (data.Length) 2 (fun i j -> 
                                                if j = 0 then
                                                    names.[i] |> XlString
                                                else
                                                    data.[i] |> conv
                                            )
            else
                Array2D.init (data.Length) 1 (fun i j -> 
                                                    data.[i] |> conv
                                            )
        new XlTable([||], xlData, "", "", false, false, false)
        
    let getErrors newOnTop =
        UdfErrorHandler.OnError |> Event.scan (fun s e -> e :: s) []
                                |> Event.map (fun errs ->
                                                  let errs = if newOnTop then errs |> List.toArray else errs |> List.rev |> List.toArray
                                                  XlTable.Create(errs, "", "", false, false, true)
                                             )

    let eval (text : string) =
        R.eval(R.parse(text = text)).AsFunction()

    let invoke (f : Function, arg0 : SymbolicExpression option, arg1 : SymbolicExpression option, arg2 : SymbolicExpression option,
                arg3 : SymbolicExpression option, arg4 : SymbolicExpression option, arg5 : SymbolicExpression option,
                arg6 : SymbolicExpression option, arg7 : SymbolicExpression option
               ) =
        let args = [|arg0; arg1; arg2; arg3; arg4; arg5; arg6; arg7|] |> Array.choose id   
        f.Invoke(args)

    let invokeNamed (f: Function, name0 : string option, arg0 : SymbolicExpression option,
                     name1 : string option, arg1 : SymbolicExpression option,
                     name2 : string option, arg2 : SymbolicExpression option,
                     name3 : string option, arg3 : SymbolicExpression option,
                     name4 : string option, arg4 : SymbolicExpression option,
                     name5 : string option, arg5 : SymbolicExpression option,
                     name6 : string option, arg6 : SymbolicExpression option,
                     name7 : string option, arg7 : SymbolicExpression option
                    ) =
        let namedArgs = [|name0, arg0; name1, arg1; name2, arg2;  name3, arg3; name4, arg4; name5, arg5;  name6, arg6; name7, arg7|]
                            |> Array.map (fun (x, y) -> match x, y with | Some(x), Some(y) -> Some(x, y) | _ -> None)
                            |> Array.choose id
        f.InvokeNamed(namedArgs)


    let readTable(file : string, separator : string, headers : bool, rowNamesCol : int) =
        let prms = namedParams ["file", box file; "sep", box separator; "header", box headers; "row.names", box rowNamesCol]
        let frame = R.read_table(prms)
        match frame with 
            | DataFrame(df) -> df
            | _ -> raise (new InvalidOperationException())

    let lm(formula : string, dataFrame : DataFrame) =
        R.lm(formula = formula, data = dataFrame)

    let summary(symbolicExpr : SymbolicExpression) =
        R.summary(symbolicExpr)

    let listMember(symbolicExpr : SymbolicExpression, listMember : string) =
        if symbolicExpr.IsList() then
            symbolicExpr.AsList().[listMember]
        else
            raise (new ArgumentException("Symbolic expr is not a list"))


    let rec asXlTable(symbolicExpr : SymbolicExpression, listMember : string option) =
        let listMember = defaultArg listMember String.Empty
        match symbolicExpr with 
            | DataFrame(dataFrame) ->
                if listMember <> String.Empty then
                    asXlTable(dataFrame.AsList().[listMember], None)
                else
                    let dbTable = frameToDataTable dataFrame
                    new XlTable(dbTable, "", "", false, false, true)

            | List(lst) -> 
                if listMember <> String.Empty then
                    asXlTable(lst.[listMember], None)
                else
                    let names = lst.Names
                    let desc = names |> Array.map (fun n -> lst.[n].Description)
                    let xlCols : XlColumn[] = [|{Name = "ListMember"; IsDateTime = false}
                                                {Name = "Type"; IsDateTime = false}
                                            |]
                    let xlData = Array2D.init names.Length 2 (fun i j -> if j = 0 then names.[i] else desc.[i]) |> Array2D.map XlString
                    new XlTable(xlCols, xlData, "", "", false, false, true)

            | NumericMatrix(v) ->
                let rowNames = v.RowNames
                let colNames = v.ColumnNames
                let data = v.ToArray()
                matrixToXlTable XlNumeric data rowNames colNames

            | IntegerMatrix(v) ->
                let rowNames = v.RowNames
                let colNames = v.ColumnNames
                let data = v.ToArray()
                matrixToXlTable (fun i -> if i = Int32.MinValue then XlNumeric(Double.NaN) else XlNumeric(float(i))) data rowNames colNames

            | CharacterMatrix(v) ->
                let rowNames = v.RowNames
                let colNames = v.ColumnNames
                let data = v.ToArray()
                matrixToXlTable XlString data rowNames colNames

            | LogicalMatrix(v) ->
                let rowNames = v.RowNames
                let colNames = v.ColumnNames
                let data = v.ToArray()
                matrixToXlTable XlBool data rowNames colNames

            | NumericVector(v) -> 
                let names = v.Names
                let data = v.ToArray()
                vectorToXlTable XlNumeric data names

            | IntegerVector(v) -> 
                let names = v.Names
                let data = v.ToArray()
                vectorToXlTable (fun i -> if i = Int32.MinValue then XlNumeric(Double.NaN) else XlNumeric(float(i))) data names

            | CharacterVector(v) -> 
                let names = v.Names
                let data = v.ToArray()
                vectorToXlTable XlString data names

            | LogicalVector(v) -> 
                let names = v.Names
                let data = v.ToArray()
                vectorToXlTable XlBool data names

            | Null -> new XlTable("Null" |> XlString)

            | RawVector(_) -> new XlTable("Raw vector" |> XlString)

            | RawMatrix(_) -> new XlTable("Raw matrix" |> XlString)

            | _ -> new XlTable("No converter to XlTable found" |> XlString)



                






