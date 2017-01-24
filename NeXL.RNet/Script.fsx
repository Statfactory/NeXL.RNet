
#r "../packages/DocumentFormat.OpenXml/lib/DocumentFormat.OpenXml.dll"
#r "../packages/NeXL/lib/net45/NeXL.XlInterop.dll"
#r "../packages/NeXL/lib/net45/NeXL.ManagedXll.dll"
#r "../packages/DynamicInterop/lib/net40/DynamicInterop.dll"

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
open NeXL.RNet
open RDotNet
open RProvider
open RProvider.``base``
open RProvider.stats
open RProvider.utils
open RProvider.datasets
open RProvider.Helpers

let csvUrl = "http://vincentarelbundock.github.io/Rdatasets/csv/datasets/mtcars.csv"
//let frame = Frame.ReadCsv(Http.RequestStream(csvUrl).ResponseStream)

let prms = namedParams ["file", box csvUrl; "sep", box ","; "header", box true; "row.names", box 1]
let rframe = R.read_table(prms)

let xx = 
    match rframe with 
        | DataFrame(df) -> df
        | _ -> raise (new NotImplementedException())

let ff = R.eval(R.parse(text="lm")).AsFunction()


let mmm = ff.InvokeNamed(("formula", R.c("mpg ~ am+wt+hp+disp+cyl")), ("data", xx:>SymbolicExpression))

let se = xx.AsList().["mpg"]

match se with
    | NumericVector(v) -> v.Names
    | _ -> [||]


let model = R.lm(formula = "mpg ~ am+wt+hp+disp+cyl", data = xx)
let s = R.summary_lm(mmm)

let nms = R.names(s)

let res = nms.AsCharacter().GetValue<string[]>()


let coeff = s.AsList().["coefficients"]

coeff.Type
let mat = coeff.AsNumericMatrix().ToArray()

let f = model.AsList().["fitted.values"]

model.IsList()

let ss = coeff.ToString()







