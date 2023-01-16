// --------------------------------------------------------------------------------------
// FAKE build script
// --------------------------------------------------------------------------------------
#r "paket: groupref build //"
#load "./.fake/build.fsx/intellisense.fsx"
#r "netstandard"

open Fake.Core
open Fake.IO
open Fake.IO.Globbing.Operators
// --------------------------------------------------------------------------------------
// Build variables
// --------------------------------------------------------------------------------------
let f projName =
    let pattern = sprintf @"**/%s.fsproj" projName
    let xs = !! pattern
    xs
    |> Seq.tryExactlyOne
    |> Option.defaultWith (fun () ->
        xs
        |> List.ofSeq
        |> failwithf "'%s' expected exactly one but:\n%A" pattern
    )

let mainProjName = "TransportationTransformer"
let mainProjPath = f mainProjName
let mainProjDir = Path.getDirectory mainProjPath

let deployDir = Path.getFullName "./deploy"

let release = ReleaseNotes.load "RELEASE_NOTES.md"
// --------------------------------------------------------------------------------------
// Helpers
// --------------------------------------------------------------------------------------
open Fake.DotNet

let dotnet cmd workingDir =
    let result = DotNet.exec (DotNet.Options.withWorkingDirectory workingDir) cmd ""
    if result.ExitCode <> 0 then failwithf "'dotnet %s' failed in %s" cmd workingDir
// --------------------------------------------------------------------------------------
// Targets
// --------------------------------------------------------------------------------------
Target.create "Clean" (fun _ -> Shell.cleanDir deployDir)

let commonBuildArgs = "-c Release"

Target.create "Build" (fun _ ->
    mainProjDir
    |> dotnet (sprintf "build %s" commonBuildArgs)
)

Target.create "Run" (fun _ ->
    mainProjDir
    |> dotnet (sprintf "run %s" commonBuildArgs)
)

Target.create "Deploy" (fun _ ->
    mainProjDir
    |> dotnet (sprintf "build %s -o \"%s\"" commonBuildArgs deployDir)
)

Target.create "Release" (fun _ ->
    mainProjDir
    |> dotnet (sprintf "build %s" commonBuildArgs)
)

Target.create "RunIlMerge" (fun _ ->
    let ILMERGE_VERSION = "3.0.41"
    let ILMerge =
        sprintf @"%%USERPROFILE%%\.nuget\packages\ilmerge\%s\tools\net452\ILMerge.exe" ILMERGE_VERSION
        |> System.Environment.ExpandEnvironmentVariables
    let APP_NAME = mainProjName

    let dstFilenameWithoutExt = sprintf "%s%A" APP_NAME release.SemVer
    let dstFilename = sprintf "%s.exe" dstFilenameWithoutExt

    let command =
        RawCommand(ILMerge, Arguments.OfArgs [
            "/wildcards"
            sprintf "/out:%s" dstFilename
            sprintf "src/bin/Release/net461/%s.exe" APP_NAME
            "src/bin/Release/net461/*.dll"
        ])
    let res = Proc.run (CreateProcess.fromCommand command)
    if res.ExitCode <> 0 then
        failwithf "%A exit with %d" command res.ExitCode

    System.IO.File.Move(dstFilename, Path.combine deployDir dstFilename)
    File.delete (sprintf "%s.pdb" dstFilenameWithoutExt)
)

// --------------------------------------------------------------------------------------
// Build order
// --------------------------------------------------------------------------------------
open Fake.Core.TargetOperators

"Build"

"Clean"
  ==> "Deploy"

"Clean"
  ==> "Release"
  ==> "RunIlMerge"

Target.runOrDefault "Deploy"
