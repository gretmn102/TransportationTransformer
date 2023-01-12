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

type Row =
    {
        /// 1.1, 1.1.1, ...
        Index: int []
        /// в двухзначных индексах стоит дата, которая распределяется по всем остальным
        Trailer: Choice<CustomDate, string>
        StartDateTime: System.DateTime
        StartLocation: string
        EndDataTime: System.DateTime
        EndLocation: string
    }

type Rows = Row []
[<CompilationRepresentation(CompilationRepresentationFlags.ModuleSuffix)>]
[<RequireQualifiedAccess>]
module Rows =
    let parseFromXLWorksheet (worksheet: IXLWorksheet) =
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

[<Struct>]
type StartOrEnd = Start | End

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
        let items = Rows.parseFromXLWorksheet worksheet

        let items =
            items
            |> List.filter (fun x -> x.Index.Length = 3)

        let allDates =
            items
            |> List.fold
                (fun st x ->
                    Set.add x.StartDateTime st
                    |> Set.add x.EndDataTime
                )
                Set.empty
            |> List.ofSeq

        let table =
            items
            |> List.groupBy (fun x -> x.Trailer)
            |> List.map (fun (trailer, xs) ->
                match trailer with
                | Choice2Of2 trailer ->
                    allDates
                    |> List.map (fun d ->
                        xs
                        |> List.tryPick (fun x ->
                            if x.StartDateTime = d then
                                Some (XLCellValue.op_Implicit x.StartLocation)
                            elif x.EndDataTime = d then
                                Some (XLCellValue.op_Implicit x.EndLocation)
                            else
                                None
                        )
                    )
                    |> fun xs -> Some (XLCellValue.op_Implicit trailer) :: xs
            )
            |> fun items ->
                let dateHeads =
                    allDates
                    |> List.groupBy (fun x -> x.Date)
                    |> List.collect (fun (d, dateTimes) ->
                        Some (XLCellValue.op_Implicit d)
                        :: List.replicate (dateTimes.Length - 1) None
                    )
                    |> fun xs -> None :: xs

                let allDates =
                    allDates
                    |> List.map (fun x -> Some <| XLCellValue.op_Implicit x)
                    |> fun xs -> None :: xs

                dateHeads :: (allDates :: items)

        do
            use workbook = new XLWorkbook()

            let worksheet = workbook.AddWorksheet()

            table
            |> List.iteri (fun y ->
                List.iteri (fun x v ->
                    v
                    |> Option.iter (fun v ->
                        let c = worksheet.Cell(y + 1, x + 1)
                        c.SetValue v
                        |> ignore
                    )
                )
            )

            workbook.SaveAs("output.xlsx")

        return Ok "Готово!"
    }

start @"e:\downloads\2023-01-10_11-56-48.xlsx"
|> printfn "%A"
