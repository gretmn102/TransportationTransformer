module Point3dTable

type Point3d<'RowValue, 'ColumnValue, 'Data> = ('RowValue * 'ColumnValue * 'Data)

type Tuple = int * int

type Table<'RowValue, 'ColumnValue, 'Data> =
    {
        Headers: 'ColumnValue []
        Values: ('RowValue * option<'Data> []) []
    }

let ofPoints (points: Point3d<'RowValue, 'ColumnValue, 'Data> []) : Table<'RowValue, 'ColumnValue, 'Data> =
    let columns =
        points
        |> Array.fold
            (fun st (rowValue, columnValue, data) ->
                Set.add columnValue st
            )
            Set.empty

    let rows =
        points
        |> Array.fold
            (fun st (rowValue, columnValue, data) ->
                let rows =
                    Map.tryFind rowValue st
                    |> Option.defaultValue Map.empty

                Map.add rowValue (Map.add columnValue data rows) st
            )
            Map.empty

    let values =
        rows
        |> Map.map (fun row cols ->
            let res =
                columns
                |> Seq.map (fun currentCol ->
                    Map.tryFind currentCol cols
                )
                |> Array.ofSeq
            row, res
        )

    {
        Headers = columns |> Array.ofSeq
        Values = values |> Seq.map (fun x -> x.Value) |> Array.ofSeq
    }
