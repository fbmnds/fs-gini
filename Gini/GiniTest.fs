module GiniTest

open Deedle
open Gini

let test_reductions(_) = 
    let mutable test_OK = true
    //  reductions    
    let x = reductions [1..10] (fun x -> x+2)
    test_OK <- test_OK && x.Equals [3; 4; 5; 6; 7; 8; 9; 10; 11; 12]

    let y = reductions [[1..10];[11..20];[21..30];[31..40]] (fun x -> (List.map (fun y -> y+2) x))
    test_OK <- test_OK && y.Equals [[3; 4; 5; 6; 7; 8; 9; 10; 11; 12]; 
                                    [13; 14; 15; 16; 17; 18; 19; 20; 21; 22];
                                    [23; 24; 25; 26; 27; 28; 29; 30; 31; 32];
                                    [33; 34; 35; 36; 37; 38; 39; 40; 41; 42]]


    let z = reductions [[1..10];[11..20];[21..30];[31..40]] (fun x -> List.fold (+) 0 x)
    test_OK <- test_OK && z.Equals [55; 155; 255; 355]

    if test_OK then printf "reduction tests OK" else printf "reduction tests FAILED"


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
printfn "%s" (peopleList.Format())

printfn "\n-----"
let excelData = Frame.ofValues [|(1,"x",3.0); (2,"x",4.0); (3,"x",43.0)|]
printf "%s" (excelData.Format())



let x = [0..999]



let test_calcGini() =
    let x = calcGini peopleList "Age" 0.0
    printf "\n%A" x 


// Expert F# 3.0, chapter 11, p. 269
// Burn some additional cycles to make sure it runs slowly enough
let rec burnSomeCycles n f s =
    if n <= 0 then f s else ignore (f s); burnSomeCycles (n - 1) f s