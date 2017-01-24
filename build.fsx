// include Fake libs
#r "./packages/DocumentFormat.OpenXml/lib/DocumentFormat.OpenXml.dll"
#r "./packages/NetOffice.Core.Net45/lib/net45/NetOffice.dll"
#r "./packages/NetOffice.Core.Net45/lib/net45/OfficeApi.dll"
#r "./packages/NetOffice.Core.Net45/lib/net45/VBIDEApi.dll"
#r "./packages/NetOffice.Excel.Net45/lib/net45/ExcelApi.dll"
#r "./packages/FAKE/tools/FakeLib.dll"
#r "./packages/NeXL/lib/net45/NeXL.XlInterop.dll"
#r "./packages/NeXL/lib/net45/NeXL.ManagedXll.dll"


open Fake
open NeXL.ManagedXll
open System
open System.IO
open System.Runtime.InteropServices
open DocumentFormat.OpenXml

let x = DocumentFormat.OpenXml.SpreadsheetDocumentType.AddIn

let nexlVer = typeof<XlCustomizationAttribute>.Assembly.GetName().Version

let sourceDir = __SOURCE_DIRECTORY__

let projName = "NeXL.RNet"

let projDir = Path.Combine(sourceDir, projName)

let xlsmPath = Path.Combine(projDir, "RNet.xlsm") // path to Excel workbook


// Directories
let buildDir  = Path.Combine(projDir, "bin/Release")

let path32BitXll = Path.Combine(projDir, sprintf "NeXL32bit_%d_%d_%d.xll" nexlVer.Major nexlVer.Minor nexlVer.Build)
let path64BitXll = Path.Combine(projDir, sprintf "NeXL64bit_%d_%d_%d.xll" nexlVer.Major nexlVer.Minor nexlVer.Build)



let nugetPackages = "Deedle 1.2.5 net40\r\nNetOffice.Core.Net45 1.7.3.0 net45\r\nNetOffice.Excel.Net45 1.7.3.0 net45"
let vcrtx64 = @"C:\Program Files (x86)\Microsoft Visual Studio 14.0\VC\redist\x64\Microsoft.VC140.CRT"
let vcrtx86 = @"C:\Program Files (x86)\Microsoft Visual Studio 14.0\VC\redist\x86\Microsoft.VC140.CRT"


// Filesets
let appReferences  =
    !! "/**/*.csproj"
    ++ "/**/*.fsproj"

// Targets
Target "Clean" (fun _ ->
    CleanDirs [buildDir]
)

Target "Build" (fun _ ->
    MSBuildRelease "" "Build" appReferences |> ignore
    WorkbookComPackager.embedCustomization(buildDir, projDir, "")
)


Target "Embed" (fun _ ->
        let custAssemblyPath = Path.Combine(projDir, sprintf "bin\\release\\%s.dll" projName)
        let asmBytes = File.ReadAllBytes(custAssemblyPath)

        let ref1AsmBytes = File.ReadAllBytes(Path.Combine(projDir, @"bin\release\RDotNet.dll"))
        let ref2AsmBytes = File.ReadAllBytes(Path.Combine(projDir, @"bin\release\RDotNet.FSharp.dll"))
        let ref3AsmBytes = File.ReadAllBytes(Path.Combine(projDir, @"bin\release\RDotNet.NativeLibrary.dll"))
        let ref4AsmBytes = File.ReadAllBytes(Path.Combine(projDir, @"bin\release\RProvider.dll"))
        let ref5AsmBytes = File.ReadAllBytes(Path.Combine(projDir, @"bin\release\RProvider.Runtime.dll"))
        let ref6AsmBytes = File.ReadAllBytes(Path.Combine(projDir, @"bin\release\DynamicInterop.dll"))


        WorkbookPackager.EmbedRuntime(xlsmPath, path32BitXll, path64BitXll, vcrtx86, vcrtx64) 
        WorkbookPackager.EmbedCustomizationAssemblies(xlsmPath, [|asmBytes; ref1AsmBytes; ref2AsmBytes; ref3AsmBytes; ref4AsmBytes; ref5AsmBytes; ref6AsmBytes|])
        WorkbookPackager.EmbedPackages(xlsmPath, nugetPackages)
        WorkbookPackager.EmbedVersion(xlsmPath)
        WorkbookPackager.EmbedProdID(xlsmPath, "10A3AF2E-482E-4FF3-97C8-2D9105C9C31A")
        WorkbookPackager.EmbedVBA(xlsmPath)
)

// Build order
"Clean"
  ==> "Build"
  ==> "Embed"

// start build
RunTargetOrDefault "Build"
