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

type Row =
    {
        /// 1.1, 1.1.1, ...
        Index: RowId
        /// в двухзначных индексах стоит дата, которая распределяется по всем остальным
        Trailer: Choice<CustomDate, string>
        StartDateTime: System.DateTime
        StartLocation: string
        EndDataTime: System.DateTime
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
                        let res =
                            cell.GetString()
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
        DateTime: System.DateTime
        Location: string
    }

type Trailer = string

type Transition =
    {
        Id: RowId
        Trailer: Trailer
        Stops: Stop []
    }

type StopIndex = int

type TransitionsId = RowId

type Transitions = Map<TransitionsId, Transition>
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

    let getStop (rowId: RowId, stopIndex: StopIndex) (transitions: Transitions) =
        Map.tryFind rowId transitions
        |> Option.map (fun transition ->
            transition.Stops.[stopIndex]
        )

type StopsByDates = Map<System.DateTime, (TransitionsId * StopIndex) list>
[<CompilationRepresentation(CompilationRepresentationFlags.ModuleSuffix)>]
[<RequireQualifiedAccess>]
module StopsByDates =
    let ofTransitions (transitions: Transitions) : StopsByDates =
        transitions
        |> Map.fold
            (fun st k x ->
                x.Stops
                |> Array.fold
                    (fun (st, index) way ->
                        let st =
                            let xs =
                                Map.tryFind way.DateTime st
                                |> Option.defaultValue []
                            Map.add way.DateTime ((k, index)::xs) st

                        st, index + 1
                    )
                    (st, 0)
                |> fst
            )
            Map.empty

    let getStop dateTime (transitions: Transitions) (stopsByDates: StopsByDates) =
        Map.tryFind dateTime stopsByDates
        |> Option.map (fun coord ->
            coord
            |> List.map (fun coord ->
                Transitions.getStop coord transitions
            )
        )

type TranstionsByTrailer = Map<Trailer, TransitionsId Set>
[<CompilationRepresentation(CompilationRepresentationFlags.ModuleSuffix)>]
[<RequireQualifiedAccess>]
module TranstionsByTrailer =
    let ofTransitions (transitions: Transitions) : TranstionsByTrailer =
        transitions
        |> Map.fold
            (fun st transitionId transition ->
                let k = transition.Trailer
                let v =
                    Map.tryFind k st
                    |> Option.defaultValue Set.empty
                Map.add k (Set.add transitionId v) st
            )
            Map.empty

type StopDateTimeIndex = int

type FinishedRow = Trailer * (StopDateTimeIndex * (TransitionsId * StopIndex)) list

type FinishedRows = FinishedRow list
[<CompilationRepresentation(CompilationRepresentationFlags.ModuleSuffix)>]
[<RequireQualifiedAccess>]
module FinishedRows =
    let create (stopsByDates: StopsByDates) (transtionsByTrailer: TranstionsByTrailer) : FinishedRows =
        transtionsByTrailer
        |> Seq.map (fun (KeyValue(trailer, transitionIds)) ->
            stopsByDates
            |> Seq.indexed
            |> Seq.choose (fun (i, (KeyValue(_, stops))) ->
                stops
                |> List.tryPick (fun (transitionId, stopIndex) ->
                    if Set.contains transitionId transitionIds then
                        Some (i, (transitionId, stopIndex))
                    else
                        None
                )
            )
            |> fun xs -> trailer, List.ofSeq xs
        )
        |> List.ofSeq

let start (xmlPath: string) =
    let getWorkbook (path: string) =
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

    let getWorksheet name (workbook: XLWorkbook) =
        match workbook.Worksheets.TryGetWorksheet name with
        | true, worksheet ->
            Ok worksheet
        | false, _ ->
            Error (sprintf "В документе не найдена \"%s\" таблица!" name)

    resultComputation {
        use workbook = getWorkbook xmlPath
        let! worksheet = getWorksheet "Поездки" workbook
        let rows = Rows.parseFromXLWorksheet worksheet
        let transitions = Transitions.ofRows rows
        let stopsByDates = StopsByDates.ofTransitions transitions
        let transtionsByTrailer = TranstionsByTrailer.ofTransitions transitions
        let finishedRows = FinishedRows.create stopsByDates transtionsByTrailer

        do
            use workbook = new XLWorkbook()

            let worksheet = workbook.AddWorksheet()

            stopsByDates
            |> Seq.iteri (fun columnIndex (KeyValue(dateTime, _)) ->
                let columnIndex = columnIndex + 1 + 1

                worksheet.Column(columnIndex).Width <- 16.144844

                let c = worksheet.Cell(1, columnIndex)
                c.SetValue (XLCellValue.op_Implicit dateTime) |> ignore
            )

            finishedRows
            |> List.iteri (fun rowIndex (trailer, xs) ->
                let rowIndex = rowIndex + 1 + 1

                worksheet.Row(rowIndex).Height <- 85.5

                let c = worksheet.Cell(rowIndex, 1)
                c.SetValue (XLCellValue.op_Implicit trailer) |> ignore

                xs
                |> List.fold
                    (fun st (columnIndex, ((transitionId, _) as coord)) ->
                        let columnIndex = columnIndex + 1 + 1
                        let c = worksheet.Cell(rowIndex, columnIndex)
                        let stop =
                            Transitions.getStop coord transitions
                            |> Option.defaultWith (fun () ->
                                failwithf "Not found %A in transitions!" coord
                            )
                        c.SetValue (XLCellValue.op_Implicit stop.Location) |> ignore

                        let setGrey (cell: IXLCell) =
                            cell.Style.Fill.BackgroundColor <- XLColor.LightGray

                        let setGreyCell (columnIndex: int) =
                            let c = worksheet.Cell(rowIndex, columnIndex)
                            setGrey c

                        let setGreyCurrentCell () =
                            setGrey c
                            Some (columnIndex, transitionId)

                        match st with
                        | None -> setGreyCurrentCell ()
                        | Some (lastColumnIndex, lastTransitionId) ->
                            if lastTransitionId = transitionId then
                                for i = lastColumnIndex + 1 to columnIndex do
                                    setGreyCell i
                                Some (columnIndex, lastTransitionId)
                            else
                                setGreyCurrentCell ()
                    )
                    None
                |> ignore
            )

            workbook.SaveAs("output.xlsx")

        return Ok "Готово!"
    }

start @"e:\downloads\2023-01-10_11-56-48.xlsx"
|> printfn "%A"
