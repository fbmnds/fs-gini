module ExcelHOFTest

open Microsoft.Office.Interop.Excel
open ExcelHOF
open Deedle

let pushDataToExcel() = 
    let excel = ApplicationClass(Visible = false)

    // Open a workbook:
    let workbookDir = __SOURCE_DIRECTORY__
    let workbooks = excel.Workbooks
    let workbook = workbooks.Open(workbookDir + @"\data\GiniTest.xlsx")
    let sheets = workbook.Worksheets

    // Get a reference to the workbook:
    let sheet = sheets.["GiniTest"] :?> Worksheet


    // column A
    sheet.Cells.[1,1] <- "x"
    for i in [2..1000] do
        sheet.Cells.[i,1] <- i-1

    // column B
    sheet.Cells.[1,2] <- "y0"
    for i in [2..1000] do
        sheet.Cells.[i,2] <- 1

    // column C
    sheet.Cells.[1,3] <- "y1"
    for i in [2..1000] do
        sheet.Cells.[i,3] <- (float (i-1))**(1./3.)

    // column D
    sheet.Cells.[1,4] <- "y2"
    for i in [2..1000] do
        sheet.Cells.[i,4] <- sqrt (float (i-1))

    // column E
    sheet.Cells.[1,5] <- "y3"
    for i in [2..1000] do
        sheet.Cells.[i,5] <- (float (i-1)) + (0.1 * 1000.0)

    // column F
    sheet.Cells.[1,6] <- "y4"
    for i in [2..1000] do
        sheet.Cells.[i,6] <- (float (i-1)) + (0.05 * 1000.0)

    // column G
    sheet.Cells.[1,7] <- "y5"
    for i in [2..1000] do
        sheet.Cells.[i,7] <- i-1

    // column H
    sheet.Cells.[1,8] <- "y6"
    for i in [2..1000] do
        sheet.Cells.[i,8] <- (float (i-1)) ** 2.0

    // column I
    sheet.Cells.[1,9] <- "y7"
    for i in [2..1000] do
        sheet.Cells.[i,9] <- (float (i-1)) ** 3.0


    // http://stackoverflow.com/questions/158706/how-to-properly-clean-up-excel-interop-objects
    System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet) |> ignore
    System.Runtime.InteropServices.Marshal.ReleaseComObject(sheets) |> ignore
    // http://stackoverflow.com/questions/19977337/closing-excel-application-with-excel-interop-without-save-message
    // http://msdn.microsoft.com/en-us/library/h1e33e36.aspx
    workbook.Save()
    workbooks.Close()
    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks) |> ignore
    excel.Quit()
    System.Runtime.InteropServices.Marshal.ReleaseComObject(excel) |> ignore


let pullDataFromExcel() : Frame<int,_> =

    let excel = ApplicationClass(Visible = false)

    // Open a workbook:
    let workbookDir =  __SOURCE_DIRECTORY__
    let workbooks = excel.Workbooks
    let workbook = workbooks.Open(workbookDir + @"\data\GiniTest.xlsx")
    let sheets = workbook.Worksheets

    // Get a reference to the workbook:
    let sheet = sheets.["GiniTest"] :?> Worksheet

    let data = seq { 
        for col in [1..9] do
            let label = cellString (sheet.Cells.[1,col] :?> Range)
            for row in [2..10] do
                yield (row-1, label, cellDouble (sheet.Cells.[row,col] :?> Range))
    } 

    let frame = Frame.ofValues data

    // http://stackoverflow.com/questions/158706/how-to-properly-clean-up-excel-interop-objects
    System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet) |> ignore
    System.Runtime.InteropServices.Marshal.ReleaseComObject(sheets) |> ignore
    // http://stackoverflow.com/questions/19977337/closing-excel-application-with-excel-interop-without-save-message
    // http://msdn.microsoft.com/en-us/library/h1e33e36.aspx
    workbook.Save()
    workbooks.Close()
    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks) |> ignore
    excel.Quit()
    System.Runtime.InteropServices.Marshal.ReleaseComObject(excel) |> ignore
    frame


// Expert F# 3.0, chapter 11, p. 269
// Burn some additional cycles to make sure it runs slowly enough
let rec burnSomeCycles n f s =
    if n <= 0 then f s else ignore (f s); burnSomeCycles (n - 1) f s


let test_ExcelSheetHandling() =

    printf "Excel data preparation ..."
    pushDataToExcel()
    printfn " done."

    printf "burn some cycles ..."
    burnSomeCycles 1000 ignore (ignore 0)
    printfn " done."

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


let test_getRangeAsArray () = 
    printfn "\n-----"
    printfn "\n workbook.getRangeAsArray 1 1 1000 9"
    use workbook = new ExcelWorkbook (Some __SOURCE_DIRECTORY__, @"data\GiniTest.xlsx")
    workbook.setSheetByName "GiniTest"
    let x = workbook.getRangeAsArray 1 1 1000 9
    printfn "%A" x



let test_intToColumn() =
    try 
        printfn "Excel index 0 : %s" (ExcelHOF.intToColumn 0)
    with
        | _ -> printfn "exception caught"
    printfn "Excel index 1 : %s" (ExcelHOF.intToColumn 1)
    printfn "Excel index 53 : %s" (ExcelHOF.intToColumn 53)
    printfn "Excel index 16384 : %s" (ExcelHOF.intToColumn 16384)
    try 
        printfn "Excel index 16385 : %s" (ExcelHOF.intToColumn 16385)
    with
        | _ -> printfn "exception caught"