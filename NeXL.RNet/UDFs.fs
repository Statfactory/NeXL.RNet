namespace NeXL.RNet
open NeXL.ManagedXll
open NeXL.XlInterop
open System
open System.IO
open System.Runtime.InteropServices
open System.Data
open FSharp.Data
open Newtonsoft.Json
open Newtonsoft.Json.Linq
open Deedle
open RProvider
open Microsoft.VisualBasic
open RDotNet
open RProvider.``base``
open RProvider.stats
open RProvider.utils
open Deedle.RPlugin
open RProvider.datasets
open RProvider.Helpers

[<XlQualifiedName(true)>]
module RNet =
    let private frameToDataTable (frame : DataFrame) : DataTable =
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
        
    let getErrors newOnTop =
        UdfErrorHandler.OnError |> Event.scan (fun s e -> e :: s) []
                                |> Event.map (fun errs ->
                                                  let errs = if newOnTop then errs |> List.toArray else errs |> List.rev |> List.toArray
                                                  XlTable.Create(errs, "", "", false, false, true)
                                             )

    let readCsv(file : string, separator : string, headers : bool, rowNames : bool) =
        let prms = namedParams ["file", box file; "sep", box separator; "header", box headers; "row.names", box 1]
        let frame = R.read_table(prms)
        match frame with 
            | DataFrame(df) -> df
            | _ -> raise (new InvalidOperationException())

    let asXlTable(dataFrame : DataFrame) =
        let dbTable = frameToDataTable dataFrame
        new XlTable(dbTable, "", "", false, false, true)


