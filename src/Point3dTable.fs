module Point3dTable

type Point3d<'X, 'Y, 'Z> = 'X * 'Y * 'Z

type Tuple = int * int

type Table<'X, 'Y, 'Z> =
    {
        Headers: 'X []
        Values: ('Y * option<'Z> []) []
    }

let ofPoints (points: Point3d<'X, 'Y, 'Z> []) : Table<'X, 'Y, 'Z> =
    let xs =
        points
        |> Array.fold
            (fun st (x, y, z) ->
                Set.add x st
            )
            Set.empty

    let values =
        points
        |> Array.fold
            (fun st (x, y, z) ->
                let zs =
                    Map.tryFind y st
                    |> Option.defaultValue Map.empty

                Map.add y (Map.add x z zs) st
            )
            Map.empty
        |> Map.map (fun y xzs ->
            let res =
                xs
                |> Seq.map (fun x ->
                    Map.tryFind x xzs
                )
                |> Array.ofSeq
            y, res
        )

    {
        Headers = xs |> Array.ofSeq
        Values = values |> Seq.map (fun x -> x.Value) |> Array.ofSeq
    }
