open ClosedXML.Excel

type ResultComputation () =
    member _.Bind(x, f) =
        match x with
        | Ok ok -> f ok
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

module DateTimeParser =
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
                (puint32)
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
                (opt (pdate .>> spaces))
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

type Item =
    {
        /// 1.1, 1.1.1, ...
        Index: int []
        /// в двухзначных индексах стоит дата, которая распределяется по всем остальным
        Trailer: Choice<CustomDateTime, string>
        StartDateTime: CustomDateTime
        StartLocation: string
        EndDataTime: CustomDateTime
        EndLocation: string
    }



let start (xmlPath: string) =
    // let xmlPath = @"e:\downloads\2023-01-10_11-56-48.xlsx"
    // System.IO.IOException

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

    // use workbook = new XLWorkbook(xmlPath)

    let getWorksheet name (workbook: XLWorkbook) =
        match workbook.Worksheets.TryGetWorksheet name with
        | true, worksheet ->
            Ok worksheet
        | false, _ ->
            Error (sprintf "В документе не найдена \"%s\" таблица!" name)

    let parseIndexes =
        let r = System.Text.RegularExpressions.Regex(@"\d+")
        fun input ->
            r.Matches input
            |> Seq.cast<System.Text.RegularExpressions.Match>
            |> Seq.map (fun x ->
                int x.Value
            )
            |> Array.ofSeq

    resultComputation {
        let! workbook = getWorkbook xmlPath
        let! worksheet = getWorksheet "Поездки" workbook


        // worksheet.Rows
        // let indexes = worksheet.Columns(1, 1)
        // let прицепы = worksheet.Columns(2, 2)
        // let startDateTimes = worksheet.Columns(3, 3)
        // let startLocations = worksheet.Columns(4, 4)
        // let endDateTimes = worksheet.Columns(6, 6)
        // let endLocations = worksheet.Columns(7, 7)


        let length =
            // `worksheet.RowCount()` возвращает 1048576, так что даже не надейся
            // worksheet.Rows()
            // |> Seq.iter (fun )
            10 // TODO
            // indexes.Cells()
            // |> Seq.findIndex (fun x -> x.Value.IsBlank)
            // |> fun x -> x + 1

        // Seq.initInfinite id
        // |> Seq.mapFold
        //     (fun st x ->
        //         match st with
        //         | Some st ->
        //             if st > 5 then x, None else x, Some (st + 1)
        //         | None ->
        //             failwithf ""
        //     )
        //     (Some 0)

        // printfn "length = %d" length
        let rows =
            worksheet.Rows()
            |> Seq.skip 1 // пропускает заголовки
            |> Seq.takeWhile (fun x ->
                not <| x.Cell(1).Value.IsBlank
            )
            |> List.ofSeq

        let items () =
            let rec f (acc, lastDate: System.DateTime option) = function
                | (row: IXLRow)::rows ->

                    let trailer = row.Cell(int ColumnHeader.Trailers)

                    // if index.Value.IsBlank then
                    //     None
                    // else
                    // Неизвестно

                    let dateOrTrailer =
                        let value = trailer.GetText()
                        match System.DateTime.TryParse value with
                        | true, date ->
                            Choice1Of2 date
                        | false, _ ->
                            Choice2Of2 value

                    let lastDate =
                        match dateOrTrailer with
                        | Choice1Of2 newDate -> Some newDate
                        | _ -> lastDate

                    let parseGetDateTime (cell: IXLCell) =
                        todo

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
                    // System.DateTime.ParseExact "18:34:32"
                    // printfn "%A" (trailer.Value, trailer.Value.Type)

                    f (item::acc, lastDate) rows
                | [] ->
                    List.rev acc
            f ([], None) rows
            // |> Seq.mapFold
            //     (fun (lastDate: System.DateTime) x ->
            //         let index = x.Cell(int ColumnHeader.Indexes)
            //         let trailer = x.Cell(int ColumnHeader.Trailers)
            //         let startDateTime = x.Cell(int ColumnHeader.StartDateTimes)
            //         let startLocation = x.Cell(int ColumnHeader.StartLocations)
            //         let endDateTime = x.Cell(int ColumnHeader.EndDateTimes)
            //         let endLocations = x.Cell(int ColumnHeader.EndLocations)

            //         if index.Value.IsBlank then
            //             None
            //         else

            //             index.Value.Type

            //     )
            //     // System.DateTime
            //     None
            //     // lastDate
        items ()

        return Ok "Готово!"
    }

start @"e:\downloads\2023-01-10_11-56-48.xlsx"
|> printfn "%A"
