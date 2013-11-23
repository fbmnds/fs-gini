// Weitere Informationen zu F# unter "http://fsharp.net".
// Weitere Hilfe finden Sie im Projekt "F#-Lernprogramm".


open GiniTest
open ExcelHOF


let waitForKey() =
    printf "\n... enter key"
    System.Console.ReadKey() |> ignore


[<EntryPoint>]
let main argv = 
    printfn "%A" argv

    printf "Excel data preparation ..."
    //pushDataToExcel()
    printfn " done."

    printf "burn some cycles ..."
    burnSomeCycles 1000 ignore (ignore 0)
    printfn " done."

    test_reductions()

    test_calcGini()

    printfn "\n-----"
    let dataFrame = pullDataFromExcel()
    printfn "%s" (dataFrame.Format())

    printfn "\n-----"
    use workbook = new ExcelWorkbook (Some __SOURCE_DIRECTORY__, @"data\GiniTest.xlsx")
    workbook.setSheetByName "GiniTest"
    let mutable dataFrame2 = workbook.getFrameWithStringHeader (1,1) (10,9)
    match dataFrame2 with
    | Some dataFrame2 -> printfn "%s" (dataFrame2.Format())
    | _ -> ignore dataFrame2

    printfn "\n-----"
    printfn "Type error: header is not string"
    /// why does F# not complain as workbook is not mutable (same issue with 'let' instead of 'use')?
    use workbook = new ExcelWorkbook (Some __SOURCE_DIRECTORY__, @"data\GiniTest.xlsx")
    workbook.setSheetByName "GiniTest"
    dataFrame2 <- workbook.getFrameWithStringHeader (2,1) (10,9)
    match dataFrame2 with
    | Some dataFrame2 -> printfn "%s" (dataFrame2.Format())
    | _ -> ignore dataFrame2

    printfn "\n-----"
    printfn "File not found error:"
    printfn "current directory: %s" System.Environment.CurrentDirectory
    use workbook = new ExcelWorkbook (None, @"..\data\GiniTest.xlsx")
    workbook.setSheetByName "GiniTest"
    dataFrame2 <- workbook.getFrameWithStringHeader (1,1) (10,9)
    match dataFrame2 with
    | Some dataFrame2 -> printfn "%s" (dataFrame2.Format())
    | _ -> ignore dataFrame2

    printfn "\n-----"
    printfn "No error:"
    printfn "current directory: %s" System.Environment.CurrentDirectory
    use workbook = new ExcelWorkbook (None, @"..\..\..\data\GiniTest.xlsx")
    workbook.setSheetByName "GiniTest"
    dataFrame2 <- workbook.getFrameWithStringHeader (1,1) (10,9)
    match dataFrame2 with
    | Some dataFrame2 -> printfn "%s" (dataFrame2.Format())
    | _ -> ignore dataFrame2

    /// there is auto dispose with 'use', no need to exlicitely do either call: 
    /// (workbook :> System.IDisposable).Dispose()
    /// workbook.close()

    waitForKey()
    0 // Exitcode aus ganzen Zahlen zurückgeben
