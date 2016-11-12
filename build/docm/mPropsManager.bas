Attribute VB_Name = "mPropsManager"
'*************************************************************
'   Module mPropsManager
'   Desc: todo
'*************************************************************
    Option Explicit
    Private Const sModule As String = "mPropsManager"
    Private oModuleScopeDoc As Document

'*************************************************************
'   Public Function PROP(...) As Variant
'   Desc:
'       An opinionated access function for the Document.CustomDocumentProperties collection
'*************************************************************
Public Function PROP( _
                        sPropID As String, _
                        Optional oDoc As Document = Nothing, _
                        Optional bReturnFalseOnDoesntExist As Boolean = False _
                        ) As Variant

    On Error GoTo Catch

    If oDoc Is Nothing Then Set oDoc = ActiveDocument
    If bReturnFalseOnDoesntExist Then
        If Not CustomPropertyExists(sPropID) Then
            PROP = False
            Exit Function
        End If
    End If

    PROP = NullToString(oDoc.CustomDocumentProperties(sPropID))
Exit Function
Catch:
    Debug.Print "Property access error. " & sPropID & " is undefined."
    Debug.Print Err.Description
End Function

'*************************************************************
'   Public Function CustomPropertyExists(sPropertyName As String, Optional oDoc As Document) As Boolean
'   Desc:
'*************************************************************
Public Function CustomPropertyExists(sPropertyName As String, Optional oDoc As Document) As Boolean
    On Error GoTo Catch
    Dim oDocumentProperty As DocumentProperty
    If oDoc Is Nothing And ActiveDocument Is Nothing Then Exit Function
    If oDoc Is Nothing Then Set oDoc = ActiveDocument

    For Each oDocumentProperty In oDoc.CustomDocumentProperties
        If oDocumentProperty.Name = sPropertyName Then
            CustomPropertyExists = True
            Exit Function
        End If
    Next
    CustomPropertyExists = False
Exit Function
Catch:
    If Err.Number = 5825 Then Exit Function                                         ' the document reference is to a deleted doc, just exit
End Function

'*************************************************************
'   Public Sub UpdateProperty(...)
'   Description:
'       Pass in a global ID, it's data type, and it's new value to refresh
'*************************************************************
Public Sub UpdateProperty(sGlobalKey As String, sValue As Variant, Optional oDoc As Document = Nothing, Optional sValueDataType As String = "String")
    On Error GoTo Catch
    Dim iPropertyMap As Long

    If oDoc Is Nothing Then
        Set oDoc = ActiveDocument
        If Not oModuleScopeDoc Is Nothing Then Set oDoc = oModuleScopeDoc           ' oModuleScopeDoc takes precedence If oDoc Is Nothing
    End If

    Select Case sValueDataType
        Case "String", "msoPropertyTypeString"
            iPropertyMap = 4
        Case "Bool", "Boolean", "msoPropertyTypeBoolean"
            iPropertyMap = 2
        Case Else
            Debug.Print "Error: UpdateProperty: Unknown Data Type."
            Exit Sub
    End Select

    If CustomPropertyExists(sGlobalKey, oDoc) Then oDoc.CustomDocumentProperties(sGlobalKey).Delete
    oDoc.CustomDocumentProperties.Add Name:=sGlobalKey, LinkToContent:=False, Type:=iPropertyMap, Value:=sValue
Exit Sub
Catch:
    If Err.Number = 5825 Then Resume Next                                           ' object was alread deleted...probably during AutoClose() trying to set CloseCancelChecked
End Sub
