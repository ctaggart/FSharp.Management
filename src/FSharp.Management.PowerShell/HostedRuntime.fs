module FSharp.Management.PowerShellProvider.HostedRuntime

open TypeInference

open System
open System.Management.Automation
open System.Management.Automation.Runspaces
open System.Threading
open System.Security.Principal

let printfn format = Printf.ksprintf System.Diagnostics.Debug.WriteLine format

type IPSRuntime =
    abstract member AllCommands  : unit -> PSCommandSignature[]
    abstract member Execute     : string * obj seq -> obj
    abstract member GetXmlDoc   : string -> string

type PowerShell with

    /// throws if there is an error after invoking
    member x.InvokeThrow() =
        let o = x.Invoke()
        if x.HadErrors then
            let sb = Text.StringBuilder()
            sb.AppendLine "failed PowerShell invoke" |> ignore
            for error in x.Streams.Error do
                sb.AppendLine (sprintf "%A" error) |> ignore
            failwith (sb.ToString())
        o

    /// gets the base objects
    member x.GetSeq() =
        x.InvokeThrow() |> Seq.map (fun x -> x.BaseObject)

    /// gets the base objects of a specific type
    /// Get-Command and Get-Member can be used to figure out the base types
    member x.GetSeq<'T>() =
        seq {
            for o in x.GetSeq() do
                match o with 
                | :? 'T as t -> yield t 
                | _ -> () }

type Runspaces.Runspace with

    member x.CreateCommand (cmdlet:string) =
        PowerShell.Create(Runspace=x).AddCommand(cmdlet)

    member x.ImportModule name =
        use c = x.CreateCommand("Import-Module").AddParameter("Name", name)
        c.InvokeThrow() |> ignore

/// PowerShell runtime built into the current process
type PSRuntimeHosted(snapIns:string[], modules:string[]) =
    let runSpace =
        try
            let initState = InitialSessionState.CreateDefault()
            initState.AuthorizationManager <- new Microsoft.PowerShell.PSAuthorizationManager("Microsoft.PowerShell")
            
            // Import SnapIns
            for snapIn in snapIns do
                if not <| String.IsNullOrEmpty(snapIn) then
                    let _, ex = initState.ImportPSSnapIn(snapIn)
                    if ex <> null then
                        failwithf "ImportPSSnapInExceptions: %s" ex.Message

            // Import modules
            let modules = modules |> Array.filter (String.IsNullOrWhiteSpace >> not)
            if not <| Array.isEmpty modules then
                initState.ImportPSModule(modules);

//            for p in initState.Providers do
//                printfn "%s %s" p.Name p.ImplementingType.FullName

            let rs = RunspaceFactory.CreateRunspace(initState)
//            let rs = RunspaceFactory.CreateRunspace()
            Thread.CurrentPrincipal <- GenericPrincipal(GenericIdentity("PowerShellTypeProvider"), null)
            rs.Open()
//            modules |> Array.iter rs.ImportModule
            rs.ImportModule @"C:\Windows\System32\WindowsPowerShell\v1.0\Modules\Hyper-V\1.1\Hyper-V.psd1"
            rs
        with
        | e -> failwithf "Could not create PowerShell Runspace: '%s'" e.Message

    do
        // get the loaded modules
        use ps = runSpace.CreateCommand("Get-Module")
        for mi in ps.GetSeq<PSModuleInfo>() do
            printfn "mi: %s" mi.Name
            for KeyValue(c,ci) in mi.ExportedCmdlets do
                if ci.Name = "Get-VHD" then
                    printfn "%s %A" c ci.OutputType

    let commandInfos =
        //Get-Command -CommandType @("cmdlet","function") -ListImported
        use ps = runSpace.CreateCommand("Get-Command")
                    .AddParameter("CommandType", "cmdlet")  // Get only cmdlets and functions (without aliases)
                    .AddParameter("ListImported")                         // Get only commands imported into current runtime
        ps.InvokeThrow()
        |> Seq.map (fun x ->
            match x.BaseObject with
            | :? CommandInfo as ci ->
//                printfn "%s %s %A" ci.ModuleName ci.Name ci.OutputType
                ci
            | w -> failwithf "Unsupported type of command: %A" w)
        |> Seq.toArray
    let commands =
        try
            commandInfos
            |> Seq.map (fun cmd ->
                match cmd with
                | :? CmdletInfo | :? FunctionInfo->
                    if cmd.ParameterSets.Count > 0 then
                        seq {
                            // Generate function for each command's parameter set
                            for pSet in cmd.ParameterSets do
                                let cmdSignature = getPSCommandSignature cmd pSet
                                yield cmdSignature.UniqueID, cmdSignature
                        }
                    else
                        failwithf "Command is not loaded: %A" cmd
                | _ -> failwithf "Unexpected type of command: %A" cmd
            )
            |> Seq.concat
            |> Map.ofSeq
        with
        | e -> failwithf "Could not load command: %s\n%s" e.Message e.StackTrace
    let allCommands = commands |> Map.toSeq |> Seq.map snd |> Seq.toArray

    do
        for cmd in allCommands do
            if cmd.Name = "Get-VHD" then
                printfn "cmd %s %s %A" cmd.Name cmd.UniqueID cmd.ResultObjectTypes

    let getXmlDoc(cmdName:string) =
        let result =
            let ps = runSpace.CreateCommand("Get-Help")
                        .AddParameter("Name", cmdName)
            ps.InvokeThrow()
            |> Seq.toArray

        let (?) (this : PSObject) (prop : string) : obj =
            let prop = this.Properties |> Seq.find (fun p -> p.Name = prop)
            prop.Value
        match result with
        | [|help|] ->
            let lines =
                let description = (help?description :?> obj [])
                if description = null then [||]
                else
                    description
                    |> CollectionConverter<PSObject>.Convert
                    |> List.toArray
                    |> Array.map (fun x->x?Text :?> string)
            sprintf "<summary><para>%s</para></summary>"
                (String.Join("</para><para>", lines |> Array.map (fun s->s.Replace("<","").Replace(">",""))))
        | _ -> String.Empty
    let xmlDocs = System.Collections.Generic.Dictionary<_,_>()

    interface IPSRuntime with
        member __.AllCommands() = allCommands
        member __.Execute(uniqueId, parameters:obj seq) =
            let cmd = commands.[uniqueId]

            // Create and execute PowerShell command
            use ps = runSpace.CreateCommand(cmd.Name)
            parameters |> Seq.iteri (fun i value->
                let key, _,ty = cmd.ParametersInfo.[i]
                match ty with
                | _ when ty = typeof<System.Management.Automation.SwitchParameter>->
                    if (unbox<bool> value) then
                        ps.AddParameter(key) |> ignore
                | _ when ty.IsValueType ->
                    if (value <> System.Activator.CreateInstance(ty))
                    then ps.AddParameter(key, value) |> ignore
                | _ ->
                    if (value <> null)
                    then ps.AddParameter(key, value) |> ignore
            )
            let result = ps.InvokeThrow()

            // Infer type of the result
            match getTypeOfObjects cmd.ResultObjectTypes result with
            | None -> box None  // Result of execution is empty collection
            | Some(tyOfObj) ->
                let collectionConverter =
                    typedefof<CollectionConverter<_>>.MakeGenericType(tyOfObj)
                let collectionObj =
                    if (tyOfObj = typeof<PSObject>) then box result
                    else result |> Seq.map (fun x->x.BaseObject) |> box
                let typedCollection =
                    collectionConverter.GetMethod("Convert").Invoke(null, [|collectionObj|])

                let choise =
                    if (cmd.ResultObjectTypes.Length = 1)
                    then typedCollection
                    else let ind = cmd.ResultObjectTypes |> Array.findIndex (fun x-> x = tyOfObj)
                         let funcName = sprintf "NewChoice%dOf%d" (ind+1) (cmd.ResultObjectTypes.Length)
                         cmd.ResultType.GetGenericArguments().[0] // GenericTypeArguments in .NET 4.5
                            .GetMethod(funcName).Invoke(null, [|typedCollection|])

                cmd.ResultType.GetMethod("Some").Invoke(null, [|choise|])
        member __.GetXmlDoc (cmdName:string) =
            if not <| xmlDocs.ContainsKey cmdName
                then xmlDocs.Add(cmdName, getXmlDoc cmdName)
            xmlDocs.[cmdName]
