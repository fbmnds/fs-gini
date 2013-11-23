module Log

// http://litemedia.info/application-logging-for-fsharp

//let private _log = log4net.LogManager.GetLogger("litemedia")
//let debug format = Printf.ksprintf _log.Debug format

let debug = eprintfn "[DEBUG]: %s"  