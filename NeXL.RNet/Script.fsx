
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

open NeXL.ManagedXll
open NeXL.XlInterop
open System
open System.IO
open System.Runtime.InteropServices
open System.Data
open RDotNet
open RProvider
open RProvider.``base``
open RProvider.stats
open RProvider.utils
open RProvider.datasets
open RProvider.Helpers









