Attribute VB_Name = "mAutoOpen"
'note: I keep having trouble saving through the Word interface. Run thisdocument.save in the immediate window as a workaround.

'*************************************************************
'   Module mAutoOpen
'   Desc:
'       This module holds the AutoOpen method which fires on open. It also defines orphan methods
'       which are called by AutoOpen and they have no other better-fitting module home.
'
'       This module also defines all VBA global variables
'*************************************************************
    Option Explicit
    Private Const sModule       As String = "mAutoOpen"
    Private Const bDebugMode    As Boolean = True

'*************************************************************
'   Public Sub AutoOpen
'   Desc:
'      This sub runs on document open and sets some custom properties.
'      Invokes webCheck to see if doc should quit itself after processing.
'*************************************************************
Public Sub AutoOpen()
    On Error GoTo Catch
    Dim oDoc As Document
    Set oDoc = ActiveDocument

    If bDebugMode Then Stop
    LoadEarlyPROPS oDoc
    ProcessMessage oDoc

Catch:
    ErrorHandler Err, sModule, "AutoOpen"
End Sub

'*************************************************************
'   Private Sub ProcessMessage(oWebCheckDoc As Document)
'   Desc:
'       Formerly Public Function webCheck
'   This sub is used to obtain information passed in the uri from the browser client to the VBA process.
'   This sub also helps determine whether Word should close itself immediately after executing the VBA process.
'*************************************************************
'todo: maybe use Application.Run PROP("WebCheckOption", oWebCheckDoc) NOT Select Case PROP("WebCheckOption", oWebCheckDoc)
Private Sub ProcessMessage(oWebCheckDoc As Document)
    On Error GoTo Catch
    Dim oWin    As Long
    Dim sAction As String

    sAction = LCase(PROP("WebCheckOption", oWebCheckDoc))
    Select Case sAction
        Case "PromptUserForInput"
            PromptUserForInput
    End Select

    Select Case sAction
        Case "focusworddocument"                                                                'don't FocusDocumentByIdTest
            oWebCheckDoc.Close False
        Case Else
            oWin = FocusOnChrome(oWebCheckDoc)
            UpdateProperty "SetWindowFocusAndDie", oWin, oWebCheckDoc
            oWebCheckDoc.Close False                                                            'close before SetForeground because close itself takes focus
    End Select

CleanExit:
    Exit Sub
Catch:
    ErrorHandler Err, sModule, "ProcessMessage"
    Resume Next
End Sub

Private Sub PromptUserForInput()
    Dim sInput As String
    Dim sURL As String

    sURL = InputBox("Enter some text. This text will be sent to the web browser.")
    'sURL = GetContext(False) & URLEncode(sURL)               'if u want to be cool u can do special encoding
    sURL = GetContext(False) & sURL
    OpenURLinBrowser sURL
End Sub

'messenger should usually focus on OC Chrome Application
'returns the window on which we need to focus
Private Function FocusOnChrome(oDoc As Document) As Long
    On Error GoTo Catch
    If PROP("DocSetID", oDoc) = "" Then                                                         'focus on the OC Console Window
        FocusOnChrome = CLng(FindWindow(vbNullString, PROP("CONSTS_ConsoleLabel")))
    Else                                                                                        'focus on specified OC extended viewer window
        FocusOnChrome = CLng(FindWindow(vbNullString, PROP("ViewerWindowLabel")))
    End If

CleanExit:
    Exit Function
Catch:
    ErrorHandler Err, sModule, "FocusOnChrome"
    Resume CleanExit
    Resume
End Function
