module Log

// http://litemedia.info/application-logging-for-fsharp

//let private _log = log4net.LogManager.GetLogger("litemedia")
//let debug format = Printf.ksprintf _log.Debug format

let debug = eprintfn "[DEBUG]: %s"  
let error = eprintfn "[ERROR]: %s"  
let warn = eprintfn "[WARN]: %s"  
let info = eprintfn "[INFO]: %s"  