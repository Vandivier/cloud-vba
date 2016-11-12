Attribute VB_Name = "mAutoOpen"
'application todo:
'   1) Fix the error handler
'   2) Add GetHTTPResponse()
'   3) Better module organization and comments
'   4) UpdateProperty to include Integer
'   5) LoadEarlyPROPs, LoadLatePROPs, and default values
'   6) other code improvements
    Option Explicit
    Private Const sModule As String = "mAutoOpen"

Public Sub AutoOpen()
    DynamicallyAssociateTemplate
End Sub

Public Sub DynamicallyAssociateTemplate()
    On Error GoTo Catch
    Dim oDoc As Document

    Set oDoc = ActiveDocument
    If InStr(oDoc.FullName, "C:\") Or InStr(oDoc.FullName, "D:\") Then Exit Sub                 'dont run at design time
    oDoc.AttachedTemplate = GetContext(False) & "/files/cloud.dotm"
    Application.Run "mAutoOpen.AutoOpen"                                                        'After you attach the other template, trigger it's AutoOpen

CleanExit:
    Exit Sub
Catch:
    ErrorHandler Err, sModule, "DynamicallyAssociateTemplate"
    Resume CleanExit
    Resume
End Sub

'*************************************************************
'   GetContext(Optional RestContext As Boolean = False) As String
'   description:
'       utility routine to return web environment context
'       assumes a valid remote location string as input
'*************************************************************
Public Function GetContext(Optional RestContext As Boolean = False) As String
    Dim sFullName As String

    sFullName = ActiveDocument.FullName
    GetContext = "http://" + Split(sFullName, "/")(2)

    If RestContext = True Then GetContext = GetContext & "/rest"
End Function

'Make sure we only manipulate documents from our application
Public Function IsCloudDoc(Optional oDoc As Document) As Boolean
    If oDoc Is Nothing Then Set oDoc = ActiveDocument
    IsCloudDoc = (InStr(LCase(oDoc.AttachedTemplate), "cloud") = 1)
End Function

