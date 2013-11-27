// Weitere Informationen zu F# unter "http://fsharp.net".
// Weitere Hilfe finden Sie im Projekt "F#-Lernprogramm".


open GiniTest
open ExcelHOFTest


let waitForKey() =
    printf "\n... enter key"
    System.Console.ReadKey() |> ignore


[<EntryPoint>]
let main argv = 
    printfn "%A" argv

    printfn "\n-----"
    printfn "test 'reductions'"
    test_reductions()

    printfn "\n-----"
    printfn "development test 'calcGini'"
    test_calcGini()

    printfn "\n-----"
    //printfn "test 'calcGini' on wikipedia examples"
    //test_calcGini_Wikipedia()

    //test_ExcelSheetHandling()

    test_getRangeAsArray ()

    test_getFrameWithHeader()

    waitForKey()
    0 // Exitcode aus ganzen Zahlen zurückgeben
