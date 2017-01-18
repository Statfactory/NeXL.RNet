#r "../packages/DocumentFormat.OpenXml/lib/DocumentFormat.OpenXml.dll"
#r "../packages/NeXL/lib/net45/NeXL.XlInterop.dll"
#r "../packages/NeXL/lib/net45/NeXL.ManagedXll.dll"
#r "../packages/FSharp.Data/lib/net40/FSharp.Data.dll"
#r "../packages/Newtonsoft.Json/lib/net45/Newtonsoft.Json.dll"
#r "../packages/Deedle/lib/net40/Deedle.dll"

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
