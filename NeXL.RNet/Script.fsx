
#r "../packages/DocumentFormat.OpenXml/lib/DocumentFormat.OpenXml.dll"
#r "../packages/NeXL/lib/net45/NeXL.XlInterop.dll"
#r "../packages/NeXL/lib/net45/NeXL.ManagedXll.dll"
#r "../packages/FSharp.Data/lib/net40/FSharp.Data.dll"
#r "../packages/Newtonsoft.Json/lib/net45/Newtonsoft.Json.dll"
#r "../packages/Deedle/lib/net40/Deedle.dll"
#r "../packages/Deedle.RPlugin/lib/net40/Deedle.RProvider.Plugin.dll"

#I "../packages/RProvider/lib/net40"
#I "../packages/R.NET.Community/lib/net40"
#r "RProvider.dll"
#r "RProvider.Runtime.dll"
#r "RDotNet.dll"
#r "../packages/R.NET.Community.FSharp/lib/net40/RDotNet.FSharp.dll"

#load "Types.fs"

open NeXL.ManagedXll
open NeXL.XlInterop
open System
open System.IO
open System.Runtime.InteropServices
open System.Data
open FSharp.Data
open FSharp.Data.JsonExtensions
open FSharp.Data.HtmlExtensions
open Newtonsoft.Json
open Newtonsoft.Json.Linq
open Deedle
open NeXL.RNet
open RDotNet
open RProvider
open RProvider.``base``
open RProvider.stats
open RProvider.utils
open Deedle.RPlugin
open RProvider.datasets
open RProvider.Helpers

do fsi.AddPrinter(fun (printer:Deedle.Internal.IFsiFormattable) -> "\n" + (printer.Format()))
let csvUrl = "http://vincentarelbundock.github.io/Rdatasets/csv/datasets/airquality.csv"
//let frame = Frame.ReadCsv(Http.RequestStream(csvUrl).ResponseStream)

let prms = namedParams ["file", box csvUrl; "sep", box ","; "header", box true; "row.names", box 1]
let rframe = R.read_table(prms)

let xx = 
    match rframe with 
        | DataFrame(df) -> df
        | _ -> raise (new NotImplementedException())

let v = xx.AsList().[xx.ColumnNames.[0]]
let res =
    match v with   
        | IntegerVector(x) -> x.GetValue<int[]>()
        | _ -> raise (new NotImplementedException())









