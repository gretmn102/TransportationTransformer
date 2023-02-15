module Point3dTableTests
open Fuchu

[<Tests>]
let toTableTests =
    testList "toTableTests" [
        testCase "base" <| fun () ->
            let ptime = System.TimeSpan.Parse

            let points: Point3dTable.Point3d<_,_,_> [] =
                [|
                    "п1", "14:40:00", "положение1"
                    "п1", "14:50:00", "положение2"
                    "п2", "14:40:25", "положение3"
                    "п2", "14:50:00", "положение4"
                |]
                |> Array.map (fun (x, y, z) -> x, ptime y, z)

            let exp =
                {
                    Point3dTable.Headers =
                        [|ptime "14:40:00"; ptime "14:40:25"; ptime "14:50:00"|]
                    Point3dTable.Values = [|
                        ("п1", [|Some "положение1"; None;              Some "положение2"|])
                        ("п2", [|None;              Some "положение3"; Some "положение4"|])
                    |]
                }
            let act = Point3dTable.ofPoints points
            Assert.Equal("", exp, act)
    ]
