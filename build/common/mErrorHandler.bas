Attribute VB_Name = "mErrorHandler"
'*************************************************************
'   Module mErrorHandler
'   Desc:
'       it logs errors when you call like:
'           ErrorHandler sModuleName, "FunctionName"
'
'       it appends to .\Documents\CloudDemoApp\log-file.txt
'
'       it also returns a string which can be optionally used in the UI
'
'       Turn on additional logging by running:
'           UpdateProperty "PrintLogs", True, ActiveDocument, "Boolean"
'
'       You can save and close the doc, then open again, and it will log AutoOpen and similar routines.
'*************************************************************
    Option Explicit
    Private Const sModule As String = "mErrorHandler"

'an error handler
Public Function ErrorHandler( _
                    eOriginal As ErrObject, _
                    Optional ByVal sModule As String = "<unknown module name>", _
                    Optional ByVal sMacro As String = "<unknown macro name>", _
                    Optional ByVal sNotes As String = "", _
                    Optional ByVal bSilenceLog As Boolean _
                    ) As Variant

    Dim iLogFile As Integer:            iLogFile = FreeFile
    Dim sMessage As String

    If eOriginal.Number = 5825 Then Resume Next
    If eOriginal.Number = 4198 Or eOriginal.Number = 91 Or eOriginal.Number = 4248 Or eOriginal.Number = 35602 Then Exit Function

    sMessage = vbCrLf + "  Module           = " + sModule + ""
    sMessage = sMessage + vbCrLf + "  Method           = " + sMacro
    sMessage = sMessage + vbCrLf + "  Document         = " + ActiveDocument.FullName
    sMessage = sMessage + vbCrLf + "  Err.Number       = " + CStr(eOriginal.Number)
    sMessage = sMessage + vbCrLf + "  Err.Source       = " + eOriginal.Source
    sMessage = sMessage + vbCrLf + "  Err.Description  = " + eOriginal.Description
    sMessage = sMessage + vbCrLf + "  Notes            = " + sNotes

    LogString sMessage, (Not bSilenceLog)
    Debug.Print vbCrLf + sMessage
    UpdateProperty "LastErrorMessage", sMessage
    UpdateProperty "LastErrorMessageTime", Now
    ErrorHandler = sMessage

End Function

'*************************************************************
'   Public Sub LogString(sText As String, bOverride As Boolean)
'   Desc:
'       Allows logging to the log file arbitrary text
'       Allows printing to immediate window if PROP("PrintLogs") = True
'       It's a PROP, not a silent bool, so you can persist it and log
'       AutoOpen and similar routines
'
'       Allows bForcePrint so that things can be printed in exceptional cases
'       even when the user has silent mode enabled (which it is by default).
'
'       Specifically, if the ErrorHandler uses this to ensure errors are shown.
'
'       So basically, errors are printed in the immediate window,
'       but not ordinary, non-error log messages, by default.
'*************************************************************
Public Sub LogString(ByVal sText As String, Optional ByVal bForcePrint As Boolean = False)
    On Error GoTo Catch
    On Error Resume Next

    sText = vbCrLf & Format(Now, "mm/dd/yyyy hh:mm:ss ") & sText
    If PROP("PrintLogs", , True) Or bForcePrint Then Debug.Print sText
    WriteFile sText, PROP("LogFileLocation")
CleanExit:
    Exit Sub
Catch:
    ErrorHandler Err, sModule, "LogString"
    Resume CleanExit
    Resume
End Sub

'todo: desc
Public Sub WriteFile(ByVal sText As String, ByVal sLocation As String)
    On Error GoTo Catch
    Dim iLogFile As Integer

    iLogFile = FreeFile
    Open sLocation For Append Access Write As #iLogFile
    Print #iLogFile, sText
    Close #iLogFile
CleanExit:
    Exit Sub
Catch:
    ErrorHandler Err, sModule, "WriteFile"
    Resume CleanExit
    Resume
End Sub
