Attribute VB_Name = "mPropsManager"
'*************************************************************
'   Module mPropsManager
'   Desc:
'       Contains routines related to the manipulation of custom document properties
'       This API is a customized layer above the MS-Word Document.CustomDocumentProperties API
'
'       These custom document properies include certain kinds of variables used in leiu of globals:
'           -Constants
'           -Document-Scope Variables
'           -State Indicators
'           -Others
'
'       Document-scope variables are superior to globals because they are persistant, attached
'           to a relevant document, and they do not pollute or collide with globals related to
'           other documents, applications, or processes.
'*************************************************************
    Option Explicit
    Private Const sModule As String = "mPropsManager"
    Private oModuleScopeDoc As Document

'*************************************************************
'   Sub Name:       LoadEarlyPROPS
'   Desc:
'       We always have to get these properties, even if we are just a messenger document.
'
'       We basically got global-like variables in 3 locations:
'           1) Public Vars declared above mAutoOpen.AutoOpen
'           2) PROPS defined here both here in LoadEarlyPROPS and also in LoadLatePROPS
'           3) Uninitialized PROPS which are used sometimes but not initialized by AutoOpen or its children
'
'       Uninitialized PROPS are noted below so that we have a central reference for PROPS:
'           LastErrorMessage
'           LastErrorMessageTime
'           PERSIST_DocumentHeight
'           PERSIST_DocumentWidth
'           PERSIST_DocumentLeft
'           PERSIST_DocumentTop
'           PERSIST_DocumentWindowState
'           PERSIST_OCActiveDictionaryLocation
'           todo: note other dictionary related props
'
'           Keep an eye out for the NOTCONST function. It acts kind of like a parameter but it's not
'           Protip: Utilize and compare output from proplist and propcount while debugging errors.
'
'       There are some decorators:
'           CONSTS_     -   It is a constant so it doesn't matter which doc you read it from
'           URLS_       -   It is a constant representing a URL
'           PERSIST_    -   It won't be automatically deleted or updated by delete or update all
'
'*************************************************************
Public Sub LoadEarlyPROPS(oDoc As Document)
    On Error GoTo Catch
    Dim wShell       As Object:          Set wShell = CreateObject("WScript.Shell")
    Dim sMyDocuments As String:          sMyDocuments = wShell.SpecialFolders("MyDocuments")
    Dim sAppFolder   As String:          sAppFolder = sMyDocuments & "\CloudDemoApp\"

    Set oModuleScopeDoc = oDoc

    'main PROPs section
    UpdateProperty "Path", oDoc.Path
    UpdateProperty "WebCheckOption", ParseFullName("action")

    'CONSTS
    UpdateProperty "CONSTS_ChromePath", """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe""", bIfMissing:=True
    UpdateProperty "LogFileLocation", sAppFolder & "log-file.txt"

CleanExit:
    Exit Sub
Catch:
    Debug.Print "LoadEarlyPROPS " & Err.Description
    Resume Next
End Sub

Private Function ParseFullName(sKey As String, Optional sFallbackKey As String) As String
    On Error GoTo Catch
    Dim sFullName As String
    If CustomPropertyExists("FullName") Then
        sFullName = PROP("FullName")
    Else
        sFullName = oModuleScopeDoc.FullName
    End If

    ParseFullName = ParseQueryString(sKey, sFullName)
    If ParseFullName = "" And sFallbackKey <> "" Then ParseFullName = ParseQueryString(sFallbackKey, sFullName)
CleanExit:
    Exit Function
Catch:
    ErrorHandler Err, sModule, "ParseFullName"
    Resume CleanExit
    Resume
End Function

'*************************************************************
'   Public Function PROP(...) As Variant
'   Desc:
'       An opinionated access function for the Document.CustomDocumentProperties collection
'*************************************************************
'todo: optional bUndefinedAsEmpty
'todo: optional bRefreshOnUndefined OR optional iTimesToTryRefreshing OR optional iMaxSecondsOfRefreshing
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
CleanExit:
    Exit Function
Catch:
    Debug.Print "Property access error. " & sPropID & " is undefined."
    'Debug.Print "Property access error. " & sPropID & " is undefined. Trying refresh."
    'UpdateProperty sPropID, JsonValueLite(responseText, "applicantFullName")
    Debug.Print Err.Description
End Function

'*************************************************************
'   Public Sub DeleteCustomProps(oDoc As Document)
'   Desc:
'           Clear the properties for one document
'           It also sets CONSTS and URLS to nothing
'               maybe that is tightly coupled but for now it works bc we only invoke on AutoOpen and we need both
'*************************************************************
Public Sub DeleteCustomProps(oDoc As Document, Optional ByVal bAllowPersist As Boolean = True)
    On Error GoTo Catch
    Dim oDocumentProperty As DocumentProperty

    If bAllowPersist Then                           'put the check outside For Each for performance
        For Each oDocumentProperty In oDoc.CustomDocumentProperties
            If Len(oDocumentProperty.Name) > 7 Then
                If Not Left(oDocumentProperty.Name, 7) = "PERSIST" Then oDoc.CustomDocumentProperties(oDocumentProperty.Name).Delete
            Else
                oDoc.CustomDocumentProperties(oDocumentProperty.Name).Delete
            End If
        Next oDocumentProperty
    Else
        For Each oDocumentProperty In oDoc.CustomDocumentProperties
            oDoc.CustomDocumentProperties(oDocumentProperty.Name).Delete
        Next oDocumentProperty
    End If
CleanExit:
    Exit Sub
Catch:
    ErrorHandler Err, sModule, "DeleteCustomProps"
    Resume CleanExit
    Resume
End Sub

'*************************************************************
'   Public Function CustomPropertyExists(...) As Boolean
'   Desc:
'       Tells you whether or not a PROP has been defined.
'       Optionally, you can treat empty as undefined.
'       Note: Passing bFalseOnEmpty implies the underlying data is String
'*************************************************************
Public Function CustomPropertyExists( _
        ByVal sPropertyName As String, _
        Optional oDoc As Document, _
        Optional ByVal bFalseOnEmpty As Boolean = False) As Boolean

    On Error GoTo Catch
    Dim oDocumentProperty As DocumentProperty

    If oDoc Is Nothing And ActiveDocument Is Nothing Then Exit Function
    If oDoc Is Nothing Then Set oDoc = ActiveDocument

    For Each oDocumentProperty In oDoc.CustomDocumentProperties
        If LCase(oDocumentProperty.Name) = LCase(sPropertyName) Then
            CustomPropertyExists = True
            If bFalseOnEmpty Then
                If PROP(sPropertyName, oDoc) = "" Then CustomPropertyExists = False
            End If
        End If
    Next
CleanExit:
    Exit Function
Catch:
    If Err.Number = 5825 Then Exit Function                                         ' the document reference is to a deleted doc, just exit
End Function

' If we're on dev, prompt for the value of a property.
' Delete the property first.
Public Function IfDevPromptFor(PROP As String) As String
    On Error GoTo Catch
    If InDev Then                                                                   ' Make sure we're in dev.
        ActiveDocument.CustomDocumentProperties(PROP).Delete
        IfDevPromptFor = InputBox("Enter value for " & PROP)
    End If
CleanExit:
    Exit Function
Catch:
    ErrorHandler Err, sModule, "IfDevPromptFor"
    Resume CleanExit
    Resume
End Function

' If the reference template for this document is local, it's path won't start with "\\".
' That means we're in dev.
Public Function InDev() As Boolean
    InDev = CBool(Environ("VBA_In_Development") = "InDev")
End Function

'==============================================================================
' Public Sub UpdateProperty - Update the argument custom document property.
'
' Arguments:
'   sKey = the custom property name
'   sValue = the custom document value
'   oDoc = (optional) the document to which this custom property
'           is attached. If omitted:
'               a. if oModuleScopeDoc is set, use it.
'               b. else use ActiveDocument.
'   sValueDataType = (optional) the data type for this property.
'           If omitted, assumes "String"
'   bIfMissing = (optional) flag to do the update ONLY IF
'           the custom property is currently undefined or empty. This
'           is to prevent accidental stomping of properties.
'------------------------------------------------------------------------------
'todo: maybe pass Optional regValidityCheck As RegExp
Public Sub UpdateProperty( _
            sKey As String, _
            sValue As Variant, _
            Optional oDoc As Document = Nothing, _
            Optional sValueDataType As String = "String", _
            Optional bIfMissing As Boolean = False)

    On Error GoTo Catch
    Dim iPropertyMap As Long

    If bIfMissing Then
        If CustomPropertyExists(sKey, oDoc) Then
            If sValueDataType = "String" And PROP(sKey, oDoc) <> "" Then Exit Sub
            If sValueDataType = "Boolean" Then Exit Sub
        End If
    End If

    If oDoc Is Nothing Then
        Set oDoc = ActiveDocument
        If Not oModuleScopeDoc Is Nothing Then Set oDoc = oModuleScopeDoc           ' oModuleScopeDoc takes precedence If oDoc Is Nothing
    End If

    Select Case sValueDataType
        Case "String", "msoPropertyTypeString"
            iPropertyMap = msoPropertyTypeString
        Case "Bool", "Boolean", "msoPropertyTypeBoolean"
            iPropertyMap = msoPropertyTypeBoolean
        Case "Int", "Integer", "msoPropertyTypeNumber"
            iPropertyMap = msoPropertyTypeNumber
        Case Else
            Debug.Print "Error: UpdateProperty: Unknown Data Type."
            Exit Sub
    End Select

    If CustomPropertyExists(sKey, oDoc) Then oDoc.CustomDocumentProperties(sKey).Delete
    oDoc.CustomDocumentProperties.Add Name:=sKey, LinkToContent:=False, Type:=iPropertyMap, Value:=sValue
CleanExit:
    Exit Sub
Catch:
    If Err.Number = 5825 Then Resume Next                                           ' object was alread deleted...probably during AutoClose() trying to set CloseCancelChecked
End Sub

'todo: maybe we don't need UpdateProperty anymore and we can just use this.
'*************************************************************
'   Public Function UpdateAndReturnProperty(...) As Variant
'   Description:
'       Invokes UpdateProperty and then also returns the property.
'*************************************************************
Public Function UpdateAndReturnProperty(sGlobalKey As String, sValue As Variant, Optional oDoc As Document = Nothing, Optional sValueDataType As String = "String") As Variant
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
            Exit Function
    End Select

    If CustomPropertyExists(sGlobalKey, oDoc) Then oDoc.CustomDocumentProperties(sGlobalKey).Delete
    oDoc.CustomDocumentProperties.Add Name:=sGlobalKey, LinkToContent:=False, Type:=iPropertyMap, Value:=sValue

    UpdateAndReturnProperty = sValue
CleanExit:
    Exit Function
Catch:
    If Err.Number = 5825 Then Resume Next                                           ' object was alread deleted...probably during AutoClose() trying to set CloseCancelChecked
End Function

'Use the old size and position when user saved and close document last
Public Sub GetSizeAndPosition()
    On Error GoTo Catch

    ActiveWindow.WindowState = wdWindowStateNormal
    If CustomPropertyExists("PERSIST_DocumentHeight") Then ActiveWindow.Height = CInt(PROP("PERSIST_DocumentHeight"))
    If CustomPropertyExists("PERSIST_DocumentWidth") Then ActiveWindow.Width = CInt(PROP("PERSIST_DocumentWidth"))
    If CustomPropertyExists("PERSIST_DocumentLeft") Then ActiveWindow.Left = CInt(PROP("PERSIST_DocumentLeft"))
    If CustomPropertyExists("PERSIST_DocumentTop") Then ActiveWindow.Top = CInt(PROP("PERSIST_DocumentTop"))
    If CustomPropertyExists("PERSIST_DocumentWindowState") Then ActiveWindow.WindowState = CInt(PROP("PERSIST_DocumentWindowState"))

CleanExit:
    Exit Sub
Catch:
    ErrorHandler Err, sModule, "GetSizeAndPosition"
End Sub

'Save the new size and position when user saves document
Public Sub SetSizeAndPosition(oDoc As Document, oWindow As Window)
    On Error GoTo Catch

    'use below if they want to set size of all docs to last one
    UpdateProperty "PERSIST_DocumentHeight", oWindow.Height, oDoc
    UpdateProperty "PERSIST_DocumentWidth", oWindow.Width, oDoc
    UpdateProperty "PERSIST_DocumentLeft", oWindow.Left, oDoc
    UpdateProperty "PERSIST_DocumentTop", oWindow.Top, oDoc
    UpdateProperty "PERSIST_DocumentWindowState", oWindow.WindowState, oDoc

CleanExit:
    Exit Sub
Catch:
    ErrorHandler Err, sModule, "SetSizeAndPosition"
End Sub

'*************************************************************
'   Public Sub SetCustomProperty(sKey As String, ByVal vData As Variant, Optional oDoc As Document = Nothing)
'   Desc:
'       This is one of those cases where the sub name explains it...
'*************************************************************
Public Sub SetCustomProperty(sKey As String, ByVal vData As Variant, Optional oDoc As Document = Nothing)
    On Error GoTo Catch
    Dim oProps As DocumentProperties
    Dim oProp  As DocumentProperty

    On Error Resume Next
    If oDoc Is Nothing Then
        Set oProps = ActiveDocument.CustomDocumentProperties
    Else
        Set oProps = oDoc.CustomDocumentProperties
    End If
    Set oProp = oProps.Item(sKey)
    If Not oProp Is Nothing Then
        oProp.Delete
    End If
    Select Case VarType(vData)
        Case vbString
            oProps.Add Name:=sKey, LinkToContent:=False, Type:=msoPropertyTypeString, Value:=vData
        Case vbBoolean
            oProps.Add Name:=sKey, LinkToContent:=False, Type:=msoPropertyTypeBoolean, Value:=vData
        Case vbInteger, vbLong
            oProps.Add Name:=sKey, LinkToContent:=False, Type:=msoPropertyTypeNumber, Value:=vData
    End Select
CleanExit:
    Exit Sub
Catch:
    ErrorHandler Err, sModule, "SetCustomProperty"
    Resume CleanExit
    Resume
End Sub

'*************************************************************
'   Public Function GetCustomProperty(sKey As String, Optional ByVal vDefault As Variant = "", Optional oDoc As Document = Nothing) As Variant
'   Desc:
'       This is one of those cases where the sub name explains it...
'*************************************************************
Public Function GetCustomProperty(sKey As String, Optional ByVal vDefault As Variant = "", Optional oDoc As Document = Nothing) As Variant
    On Error GoTo Catch
    Dim oProps As DocumentProperties
    Dim oProp  As DocumentProperty

    On Error Resume Next
    GetCustomProperty = vDefault
    If oDoc Is Nothing Then
        Set oProps = ActiveDocument.CustomDocumentProperties
    Else
        Set oProps = oDoc.CustomDocumentProperties
    End If
    Set oProp = oProps.Item(sKey)
    If Not oProp Is Nothing Then
        GetCustomProperty = oProp.Value
    End If
CleanExit:
    Exit Function
Catch:
    ErrorHandler Err, sModule, "GetCustomProperty"
    Resume CleanExit
    Resume
End Function

'*************************************************************
'   Public Sub DeleteCustomProperty(sKey As String, Optional oDoc As Document = Nothing)
'   Desc:
'       This is one of those cases where the sub name explains it...
'*************************************************************
Public Sub DeleteCustomProperty(sKey As String, Optional oDoc As Document = Nothing)
    On Error GoTo Catch
    Dim oProps As DocumentProperties
    Dim oProp  As DocumentProperty

    On Error Resume Next
    If oDoc Is Nothing Then
        Set oProps = ActiveDocument.CustomDocumentProperties
    Else
        Set oProps = oDoc.CustomDocumentProperties
    End If
    Set oProp = oProps.Item(sKey)
    If Not oProp Is Nothing Then
        oProp.Delete
    End If
CleanExit:
    Exit Sub
Catch:
    ErrorHandler Err, sModule, "DeleteCustomProperty"
    Resume CleanExit
    Resume
End Sub

Public Function fAIAString(Optional oDoc As Document, Optional ByVal bBothAsFalse As Boolean) As String
    On Error GoTo Catch
    If oDoc Is Nothing Then Set oDoc = ActiveDocument

    fAIAString = CStr(PROP("CFPIsAIA", oDoc, True))                             'allow for "Both" to take precedence
    If bBothAsFalse Then
        If fAIAString = "Both" Then fAIAString = "False"
    End If

    If fAIAString = "False" Then fAIAString = CStr(PROP("IsAIA", oDoc, True))   'Give "True" one more chance
CleanExit:
    Exit Function
Catch:
    ErrorHandler Err, sModule, "fAIAString"
    Resume CleanExit
    Resume
End Function

