module GiniTest

open Deedle
open Gini
open ExcelHOF
open ExcelHOFTest


let test_reductions(_) = 
    let mutable test_OK = true
    //  reductions    
    let x = reductions (fun y x -> x+2) 0 [1..10]
    test_OK <- test_OK && x.Equals [3; 4; 5; 6; 7; 8; 9; 10; 11; 12]

    let y = reductions (fun y x -> (List.map (fun y -> y+2) x)) [] [[1..10];[11..20];[21..30];[31..40]]
    test_OK <- test_OK && y.Equals [[3; 4; 5; 6; 7; 8; 9; 10; 11; 12]; 
                                    [13; 14; 15; 16; 17; 18; 19; 20; 21; 22];
                                    [23; 24; 25; 26; 27; 28; 29; 30; 31; 32];
                                    [33; 34; 35; 36; 37; 38; 39; 40; 41; 42]]


    let z = reductions (fun y x -> List.fold (+) 0 x) 0 [[1..10];[11..20];[21..30];[31..40]]
    test_OK <- test_OK && z.Equals [55; 155; 255; 355]

    if test_OK then printf "reduction tests OK" else printf "reduction tests FAILED\n x = %A\n y = %A\n z = %A" x y z


// understand Frame.ofValues
printfn "\n-----"
let someData = Frame.ofValues [|(1,"x",3.0); (2,"x",4.0); (3,"x",43.0)|]
printf "%s" (someData.Format())


type Person = 
  { Name:string; Age:int; Countries:string list; }

let peopleRecds = 
  [ { Name = "Joe"; Age = 51; Countries = [ "UK"; "US"; "UK"] }
    { Name = "Tomas"; Age = 28; Countries = [ "CZ"; "UK"; "US"; "CZ" ] }
    { Name = "Eve"; Age = 2; Countries = [ "FR" ] }
    { Name = "Suzanne"; Age = 15; Countries = [ "US" ] } ]

// Turn the list of records into data frame 
let peopleList = Frame.ofRecords peopleRecds
printfn "\n-----"
printfn "peopleList\n------%s" (peopleList.Format())


let test_calcGini() =
    let x = calcGini peopleList "Age" 0.0
    printf "\n Gini of peopleList by Age : %A" x 

let c(df) = 
    printfn "\n-----"
    [|"x";"y0";"y1";"y2";"y3";"y4";"y5";"y6";"y7"|] |> Array.iter (printf "   %s   ")
    printfn "\n-----"
    let ginis = seq { 
        for sel in ["x";"y0";"y1";"y2";"y3";"y4";"y5";"y6";"y7"] do 
            yield calcGini df sel 0.0 
    }
    ginis |> Seq.iter (printf "  %.3f ")
    printfn "\n-----\n"


let test_calcGini_Wikipedia() =
    printf "Excel data preparation ..."
    pushDataToExcel()
    printfn " done."

    printf "burn some cycles ..."
    burnSomeCycles 1000 ignore (ignore 0)
    printfn " done."

    use workbook = new ExcelWorkbook (Some __SOURCE_DIRECTORY__, @"data\GiniTest.xlsx")
    workbook.setSheetByName "GiniTest"
    let dataFrame = workbook.getFrameWithStringHeader (1,1) (1000,9)
    match dataFrame with
    | Some dataFrame -> c(dataFrame)
    | _ -> ignore dataFrame



