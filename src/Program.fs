open ClosedXML.Excel

type ResultComputation () =
    member _.Bind(x, f) =
        match x with
        | Ok ok -> f ok
        | Error err -> Error err

    member _.Using(x, f) =
        match x with
        | Ok ok -> using ok f
        | Error err -> Error err

    member _.Return x = x

let resultComputation = ResultComputation()

type ColumnHeader =
    | Indexes = 1
    | Trailers = 2
    | StartDateTimes = 3
    | StartLocations = 4
    | EndDateTimes = 6
    | EndLocations = 7

type CustomTime =
    {
        Hours: uint32
        Minutes: uint32
        Seconds: uint32 option
    }

type CustomDate =
    {
        Year: uint32
        Month: uint32
        Day: uint32
    }

type CustomDateTime =
    {
        Date: CustomDate option
        Time: CustomTime option
    }

[<CompilationRepresentation(CompilationRepresentationFlags.ModuleSuffix)>]
[<RequireQualifiedAccess>]
module CustomDateTime =
    module Parser =
        open FParsec

        type 'a Parser = Parser<'a, unit>

        let pdate: _ Parser =
            pipe3
                puint32
                (skipChar '-' >>. puint32)
                (skipChar '-' >>. puint32)
                (fun year month day ->
                    {
                        Year = year
                        Month = month
                        Day = day
                    }
                )

        let ptime: _ Parser =
            pipe3
                puint32
                (skipChar ':' >>. puint32)
                (opt (skipChar ':' >>. puint32))
                (fun hours mins secs ->
                    {
                        Hours = hours
                        Minutes = mins
                        Seconds = secs
                    }
                )

        let p: _ Parser =
            pipe2
                (opt (attempt (pdate .>> spaces)))
                (opt ptime)
                (fun date time ->
                    {
                        Date = date
                        Time = time
                    }
                )

        let runResult p str =
            match run p str with
            | Success(res, _, _) -> Result.Ok res
            | Failure(errMsg, _, _) -> Result.Error errMsg

    let parse =
        Parser.runResult Parser.p

    let toDateTime (customDateTime: CustomDateTime) =
        let date = customDateTime.Date
        let time = customDateTime.Time

        System.DateTime(
            date |> Option.map (fun x -> int x.Year) |> Option.defaultValue 1,
            date |> Option.map (fun x -> int x.Month) |> Option.defaultValue 1,
            date |> Option.map (fun x -> int x.Day) |> Option.defaultValue 1,
            time |> Option.map (fun x -> int x.Hours) |> Option.defaultValue 0,
            time |> Option.map (fun x -> int x.Minutes) |> Option.defaultValue 0,
            time |> Option.bind (fun x -> x.Seconds |> Option.map int) |> Option.defaultValue 0
        )

type RowId = int []

type DateTimeResult =
    | Unknown
    | Value of System.DateTime
[<CompilationRepresentation(CompilationRepresentationFlags.ModuleSuffix)>]
[<RequireQualifiedAccess>]
module DateTimeResult =
    let toOption (dtr: DateTimeResult) =
        match dtr with
        | Unknown -> None
        | Value date -> Some date

type Row =
    {
        /// 1.1, 1.1.1, ...
        Index: RowId
        /// в двухзначных индексах стоит дата, которая распределяется по всем остальным
        Trailer: Choice<CustomDate, string>
        StartDateTime: DateTimeResult
        StartLocation: string
        EndDataTime: DateTimeResult
        EndLocation: string
    }

type Rows = Row list
[<CompilationRepresentation(CompilationRepresentationFlags.ModuleSuffix)>]
[<RequireQualifiedAccess>]
module Rows =
    let parseFromXLWorksheet (worksheet: IXLWorksheet) : Rows =
        let parseIndexes =
            let r = System.Text.RegularExpressions.Regex(@"\d+")
            fun input ->
                r.Matches input
                |> Seq.cast<System.Text.RegularExpressions.Match>
                |> Seq.map (fun x ->
                    int x.Value
                )
                |> Array.ofSeq

        let rows =
            worksheet.Rows()
            |> Seq.skip 1 // пропускает заголовки
            |> Seq.takeWhile (fun x ->
                not <| x.Cell(1).Value.IsBlank
            )
            |> List.ofSeq

        let getItems () =
            let rec f (acc, lastDate: CustomDate option) = function
                | (row: IXLRow)::rows ->
                    let trailer = row.Cell(int ColumnHeader.Trailers)

                    let dateOrTrailer =
                        let value = trailer.GetText()
                        match CustomDateTime.Parser.runResult CustomDateTime.Parser.pdate value with
                        | Ok x ->
                            Choice1Of2 x
                        | Error _ ->
                            Choice2Of2 value

                    let lastDate =
                        match dateOrTrailer with
                        | Choice1Of2 newDate -> Some newDate
                        | _ -> lastDate

                    let parseGetDateTime (cell: IXLCell) =
                        match cell.GetString() with
                        | "Неизвестно" ->
                            DateTimeResult.Unknown
                        | cellValue ->
                        let res =
                            cellValue
                            |> CustomDateTime.parse

                        match res with
                        | Ok customDateTime ->
                            if Option.isSome customDateTime.Date then
                                customDateTime
                            else
                                { customDateTime with
                                    Date = lastDate
                                }
                            |> CustomDateTime.toDateTime
                            |> DateTimeResult.Value

                        | Error x ->
                            failwithf "%s" x

                    let item =
                        {
                            Index =
                                let index = row.Cell(int ColumnHeader.Indexes)
                                parseIndexes <| index.Value.GetText()
                            Trailer = dateOrTrailer
                            StartDateTime =
                                let startDateTime = row.Cell(int ColumnHeader.StartDateTimes)
                                parseGetDateTime startDateTime
                            StartLocation =
                                let startLocation = row.Cell(int ColumnHeader.StartLocations)
                                startLocation.GetString()
                            EndDataTime =
                                let endDateTime = row.Cell(int ColumnHeader.EndDateTimes)
                                parseGetDateTime endDateTime
                            EndLocation =
                                let endLocations = row.Cell(int ColumnHeader.EndLocations)
                                endLocations.GetString()
                        }

                    f (item::acc, lastDate) rows
                | [] ->
                    List.rev acc
            f ([], None) rows

        getItems ()

type Stop =
    {
        DateTime: DateTimeResult
        Location: string
    }

type Trailer = string

type TransitionId = RowId

type Transition =
    {
        Id: TransitionId
        Trailer: Trailer
        Stops: Stop []
    }

type StopIndex = int

type StopCoords = TransitionId * StopIndex

type Transitions = Map<TransitionId, Transition>
[<CompilationRepresentation(CompilationRepresentationFlags.ModuleSuffix)>]
[<RequireQualifiedAccess>]
module Transitions =
    let ofRows (rows: Rows) : Transitions =
        let getStops (rows: Rows) =
            let getStartedStop (row: Row) =
                {
                    DateTime = row.StartDateTime
                    Location = row.StartLocation
                }
            let getEndedStop (row: Row) =
                {
                    DateTime = row.EndDataTime
                    Location = row.EndLocation
                }

            match rows with
            | [x] ->
                [|
                    getStartedStop x
                    getEndedStop x
                |]
            | [] ->
                failwithf "Function GetStops needs at least one element in the list"
            | row::rows ->
                [|
                    getStartedStop row
                    yield!
                        rows |> List.map getEndedStop
                |]

        rows
        |> List.filter (fun x -> x.Index.Length = 4)
        |> List.groupBy (fun row ->
            let indexes = row.Index
            [|indexes.[0]; indexes.[1]; indexes.[2]|]
        )
        |> List.map (fun (id, subRows) ->
            let transition =
                {
                    Id = id
                    Stops = getStops subRows
                    Trailer =
                        let firstRow = subRows.[0]
                        match firstRow.Trailer with
                        | Choice2Of2 trailer -> trailer
                        | Choice1Of2 date ->
                            failwithf "Поле %A содержит дату %A, хотя должно содержать "
                                firstRow
                                date
                }
            id, transition
        )
        |> Map.ofList

    let getStop ((id: TransitionId, stopIndex: StopIndex): StopCoords) (transitions: Transitions) =
        Map.tryFind id transitions
        |> Option.map (fun transition ->
            transition.Stops.[stopIndex]
        )

type Data = { StartLoc: string option; EndLoc: string option }

type Point = System.DateTime * Trailer * Data

type FinishedTable = Point3dTable.Table<System.DateTime, Trailer, Data>
[<CompilationRepresentation(CompilationRepresentationFlags.ModuleSuffix)>]
[<RequireQualifiedAccess>]
module FinishedTable =
    let create (transition: Transitions) : FinishedTable =
        transition
        |> Seq.choose (fun (KeyValue(_, x)) ->
            let res =
                let f startStop endStop =
                    let get (stop: Stop) =
                        match stop.DateTime with
                        | Value x ->
                            Some (x.Date, stop.Location)
                        | Unknown ->
                            None

                    get startStop
                    |> Option.map (fun (date, startLoc) ->
                        let endLoc =
                            endStop
                            |> Option.bind (fun endStop ->
                                DateTimeResult.toOption endStop.DateTime
                                |> Option.bind (fun x ->
                                    if date = x.Date then
                                        Some endStop.Location
                                    else
                                        None
                                )
                            )

                        let data =
                            {
                                StartLoc = Some startLoc
                                EndLoc = endLoc
                            }
                        date, data
                    )

                match x.Stops with
                | [|startStop|] ->
                    f startStop None
                | [||] -> None
                | stops ->
                    f stops.[0] (Some stops.[stops.Length - 1])

            res
            |> Option.map (fun (date, data) ->
                date, x.Trailer, data
            )
        )
        |> Array.ofSeq
        |> Point3dTable.ofPoints

let start (xmlPath: string) =
    let worksheetName = "Поездки"

    let getWorkbook (path: string) =
        if System.IO.File.Exists path then
            try
                new XLWorkbook(path)
                |> Ok
            with
                | :? System.IO.IOException ->
                    sprintf "Доступ к таблице запрещен. Если он открыт у вас в Excel, то закройте Excel и попробуйте снова."
                    |> Error
                | e ->
                    sprintf "Непредвиденная ошибка:\n%A" e
                    |> Error
        else
            Error (sprintf "Файл '%s' не существует." path)

    let getWorksheet name (workbook: XLWorkbook) =
        match workbook.Worksheets.TryGetWorksheet name with
        | true, worksheet ->
            Ok worksheet
        | false, _ ->
            Error (sprintf "В документе не найдена \"%s\" таблица!" name)

    let createSheet transitions (workbook: XLWorkbook) =
        let finishedTable = FinishedTable.create transitions

        do
            let worksheet = workbook.AddWorksheet()

            printfn "Добавляю результат в таблицу '%s'..." worksheet.Name

            finishedTable.Headers
            |> Seq.iteri (fun columnIndex dateTime ->
                let columnIndex = 1 + columnIndex * 2

                let c = worksheet.Cell(1, columnIndex + 1)
                c.SetValue (XLCellValue.op_Implicit (dateTime.ToString "dd.MM.yyyy")) |> ignore

                worksheet.Column(columnIndex + 1).Width <- 16.71
                worksheet.Column(columnIndex + 2).Width <- 16.71
            )

            // trailers column
            worksheet.Column(1).Width <- 12

            finishedTable.Values
            |> Array.iteri (fun rowIndex (trailer, xs) ->
                let rowIndex = 1 + rowIndex

                worksheet.Row(rowIndex + 1).Height <- 85.5

                let c = worksheet.Cell(rowIndex + 1, 1)
                c.SetValue (XLCellValue.op_Implicit trailer) |> ignore

                xs
                |> Array.iteri
                    (fun columnIndex data ->
                        match data with
                        | Some data ->
                            let set columnIndex (value: string) =
                                let c = worksheet.Cell(rowIndex + 1, columnIndex + 1)
                                let alignment = c.Style.Alignment
                                alignment.WrapText <- true
                                alignment.Horizontal <- XLAlignmentHorizontalValues.Center
                                alignment.Vertical <- XLAlignmentVerticalValues.Center
                                c.SetValue (XLCellValue.op_Implicit value) |> ignore

                            let columnIndex = 1 + columnIndex * 2

                            data.StartLoc
                            |> Option.iter (set columnIndex)

                            data.EndLoc
                            |> Option.iter (set (columnIndex + 1))
                        | None ->
                            ()
                    )
            )

            worksheet.SheetView.Freeze(1, 1)

    resultComputation {
        use workbook = getWorkbook xmlPath
        let! worksheet = getWorksheet worksheetName workbook
        let rows = Rows.parseFromXLWorksheet worksheet
        let transitions = Transitions.ofRows rows

        createSheet transitions workbook

        printfn "Сохраняю документ..."

        workbook.SaveAs(xmlPath)

        return Ok "Готово!"
    }

[<EntryPoint>]
let main args =
    let resultCode =
        match args with
        | [||] ->
            printfn "Перетащите на программу файл Excel"
            2
        | [|path|] ->
            printfn "ОБрабатываю %s..." path
            match start path with
            | Ok resMsg ->
                printfn "%s" resMsg
                0
            | Error errMsg ->
                printfn "%s" errMsg
                1
        | xs ->
            printfn "Перетащите на программу файл Excel"
            2

    printfn "Для выхода из программы нажмите любую клавишу..."
    System.Console.ReadKey () |> ignore

    resultCode
