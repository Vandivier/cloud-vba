Attribute VB_Name = "mPublicMethods"
'*************************************************************
'   Module mPublicMethods
'   Desc:
'       This module contains public methods!
'       - HTTP transaction methods
'       - Base64 encode/decode
'       - String-based JSON manipulation (not object or script-based)
'       - Some navigation-related methods
'       - Miscellaneous other stuff
'*************************************************************
    Option Explicit
    Private Const sModule As String = "mPublicMethods"
    Private sJWTToken As String                                           'in memory bc must not be persistent; it is a session-duration param for GetHTTPResponse

'*************************************************************
'   name:       NullToString
'   description:
'       It will check the first parameter passed and do a nullcheck.
'       If it's a null it will return an empty string by default,
'       but you can change that by passing an optional second argument.
'*************************************************************
Public Function NullToString(Value As Variant, Optional valueIfNull As Variant = "") As Variant
    On Error GoTo Catch
    If VarType(Value) = vbObject Then
        If Value Is Nothing Then NullToString = valueIfNull
    ElseIf VarType(Value) = vbNull Then
        If IsNull(Value) Then NullToString = valueIfNull
    Else
        NullToString = Value
    End If
CleanExit:
    Exit Function
Catch:
    ErrorHandler Err, sModule, "NullToString"
    Resume CleanExit
    Resume
End Function

' Finds the key corresponding to the key in the dictionary.
Public Function DictFindKeyByVal(dict As Scripting.Dictionary, vValue As Variant) As String
    On Error GoTo Catch
    Dim vKey As Variant
    With dict
        For Each vKey In .Keys
            If .Item(vKey) = vValue Then
                DictFindKeyByVal = vKey
                Exit Function
            End If
        Next vKey
    End With
CleanExit:
    Exit Function
Catch:
    ErrorHandler Err, sModule, "DictFindKeyByVal"
    Resume CleanExit
    Resume
End Function


'*************************************************************
'   name:       currentField
'   description:
'       returns the current field as a field object, where current field is defined
'       as either the field that the cursor is inside of, or the first field in the selected range
'*************************************************************
Public Function currentField() As Field
    On Error GoTo Catch
    Dim iStart As Long
    Dim iEnd   As Long
    Dim oField As Field

    iStart = Selection.Range.Start
    iEnd = Selection.Range.End

    For Each oField In ActiveDocument.Fields
        If ((iStart <= oField.Result.Start) Or (iStart <= oField.Code.Start)) _
        And ((oField.Result.End <= iEnd) Or (oField.Code.End <= iEnd)) Then
                Set currentField = oField
                Exit Function
        End If
    Next oField
    Set currentField = Nothing
CleanExit:
    Exit Function
Catch:
    ErrorHandler Err, sModule, "currentField"
    Resume CleanExit
    Resume
End Function

'===============================================================================
' Function currentFPRange - returns the Range for the FP where the cursor is.
'
' This loops through all the bookmarks in the document to find the bookmark
'   whose beginning is <= cursor's left and whose end is >= cursor's left.
'   (YES, that's left again.)
'
' NOTE: As of 2016-08-26 FP bookmarks span the range of the corresponding FP,
'   and no longer just mark the beginning of the FP. This *will* have side
'   effects.
'
' NOTE2: FP bookmarks are abutted directly next to each other. That is,
'   Bookmark(FP(i)).End == Bookmark(FP(i+1)).Begin
' As a result, we *don't* want the bookmark when oBook.Range.End = iPos, but
'   only the range where oBook.Range.End > iPos.
'
'-------------------------------------------------------------------------------
Public Function currentFPRange() As Range
    On Error GoTo Catch
    Dim iPos  As Long
    Dim oBook As Bookmark

    iPos = Selection.Range.Start

    For Each oBook In ActiveDocument.Bookmarks
        If ((oBook.Range.Start <= iPos) And (oBook.Range.End > iPos)) Then
            Set currentFPRange = oBook.Range
            Exit Function
        End If
    Next oBook
    Set currentFPRange = Nothing
CleanExit:
    Exit Function
Catch:
    ErrorHandler Err, sModule, "currentFPRange"
    Resume CleanExit
    Resume
End Function

'*************************************************************
'   name:       JsonValueLite
'   description:
'       Given a JSON string or response text, find the value of a given key.
'*************************************************************
Public Function JsonValueLite(ByVal JsonString As String, FieldName As String, Optional sDataType As String)
    On Error GoTo Catch
    Dim index1 As Integer
    Dim index2 As Integer
    Dim Value  As String
    If sDataType = "number" Or sDataType = "boolean" Then
        index1 = InStr(JsonString, """" & FieldName & """:")
        If index1 = 0 Then
            JsonValueLite = ""
            Exit Function
        End If
        index1 = index1 + Len(FieldName) + 3
        index2 = InStr(index1, JsonString, ",")
    Else
        index1 = InStr(JsonString, """" & FieldName & """:""")
        If index1 = 0 Then
            JsonValueLite = ""
            Exit Function
        End If
        index1 = index1 + Len(FieldName) + 4
        index2 = InStr(index1, JsonString, """")
    End If
    Value = Mid(JsonString, index1, index2 - index1)
    JsonValueLite = Value
CleanExit:
    Exit Function
Catch:
    ErrorHandler Err, sModule, "JsonValueLite"
    Resume CleanExit
    Resume
End Function

'*************************************************************
'   name:       OpenURLinBrowser
'   description:
'       It opens a URL. In a browser.
'*************************************************************
Public Sub OpenURLinBrowser(DestURL As String, Optional bSmallWindow As Boolean = True)
    Dim sHTML As String
    If bSmallWindow = True Then
        sHTML = " --profile-directory=""Default"""
        sHTML = sHTML & " --app=""data:text/html,<html><body><script>"
        sHTML = sHTML & "window.moveTo(5000,5000);"
        sHTML = sHTML & "window.resizeTo(10,10);"
        sHTML = sHTML & "window.location='" & URLEncode(DestURL) & "';"
        sHTML = sHTML & "</script></body></html>"""
        Shell PROP("CONSTS_ChromePath") & sHTML, vbMinimizedNoFocus             'or vbHide ?
    Else
        Shell PROP("CONSTS_ChromePath") & " -url " & DestURL, vbNormalFocus
    End If
End Sub

'todo: merge GetBaseUrl and GetContext
'Perhaps implement as CONSTS, NOTCONST, or PROP
'*************************************************************
'   name:       GetBaseUrl
'   description:
'*************************************************************
Public Function GetBaseUrl()
    On Error GoTo Catch
    Dim sFullName As String
    Dim baseUrl   As String

    sFullName = IIf(CustomPropertyExists("FullName"), PROP("FullName"), ActiveDocument.FullName)

    If InStr(sFullName, "/ocweb/") > 0 Then                                 'rest context
        GetBaseUrl = Left(sFullName, InStr(sFullName, "/ocweb/") - 1)
    ElseIf InStr(sFullName, ".gov/oc") > 0 Then                             'non-rest context...probably a messenger.docx
        GetBaseUrl = Left(sFullName, InStr(sFullName, ".gov/oc") + 3)
    Else
        GetBaseUrl = "http://dav.dev.uspto.gov"                             'fallback
    End If
CleanExit:
    Exit Function
Catch:
    ErrorHandler Err, sModule, "GetBaseUrl"
    Resume CleanExit
    Resume
End Function

'todo: merge GetBaseUrl and GetContext
'Perhaps implement as CONSTS, NOTCONST, or PROP
'*************************************************************
'   name:       GetContext
'   description:
'       utility routine to return web environment context
'       pass optional boolean because there is a different web context for rest service endpoints
'*************************************************************
Public Function GetContext(Optional RestContext As Boolean = False) As String
    On Error GoTo Catch
    Dim sFullName As String

    If PROP("PERSIST_IsCFP", , True) Then
        GetContext = PROP("CFPBaseUrl")                                  'inhereted from template at CFP creation
    Else
        If Not CustomPropertyExists("FullName") Then Exit Function          'todo: throw GUI err? maybe MN is making a 1pt story for S8 @vandivier
        sFullName = PROP("FullName")
        GetContext = "http://" + Split(sFullName, "/")(2)
    End If

    If InStr(GetContext, "localhost") = 0 And RestContext = False Then GetContext = GetContext & "/oc"
    If RestContext = True Then GetContext = GetContext & "/ocweb/rest"
CleanExit:
    Exit Function
Catch:
    ErrorHandler Err, sModule, "GetContext"
    Resume CleanExit
    Resume
End Function

'If a document is in the documents collection return it
'otherwise make one in the TempFolder and return it
Public Function GetOrCreateDoc( _
            ByVal sName As String, _
            Optional ByVal iFormat As WdSaveFormat = wdFormatRTF, _
            Optional ByVal sExtension As String = ".rtf", _
            Optional ByVal bVisible As Boolean, _
            Optional ByVal bNameAsPath As Boolean _
            ) As Document

    On Error GoTo Catch
    Dim sFullName As String
    Dim oDoc      As Document

    For Each oDoc In Application.Documents
        If oDoc.Name = (sName & sExtension) Then Set GetOrCreateDoc = oDoc
    Next oDoc

    If GetOrCreateDoc Is Nothing Then
        If bNameAsPath Then
            sFullName = sName
        Else
            sFullName = PROP("TempFolder") & sName & sExtension
        End If

        Set GetOrCreateDoc = Documents.Add(Visible:=bVisible)
        GetOrCreateDoc.SaveAs2 sFullName, iFormat
    End If
CleanExit:
    Exit Function
Catch:
    ErrorHandler Err, sModule, "GetOrCreateDoc"
End Function

'*
'The function encodes the string so that the spaces and tabs are preserved when sent using xmlhttp request
'*
Function URLEncode(EncodeStr As String) As String
    On Error GoTo Catch
    Dim i As Integer
    Dim erg As String

    erg = EncodeStr
    erg = Replace(erg, "%", Chr(1)) ' *** First replace '%' chr
    erg = Replace(erg, "+", Chr(2)) ' *** then '+' chr
    For i = 0 To 255
        Select Case i
            Case 37, 43, 48 To 57, 65 To 90, 97 To 122 ' *** Allowed 'regular' characters
            Case 1  ' *** Replace original %
                erg = Replace(erg, Chr(i), "%25")
            Case 2  ' *** Replace original +
                erg = Replace(erg, Chr(i), "%2B")
            Case 32
                erg = Replace(erg, Chr(i), "+")
            Case 3 To 15
                erg = Replace(erg, Chr(i), "%0" & Hex(i))
            Case Else
                erg = Replace(erg, Chr(i), "%" & Hex(i))
        End Select
    Next
    URLEncode = erg
CleanExit:
    Exit Function
Catch:
    ErrorHandler Err, sModule, "URLEncode"
    Resume CleanExit
    Resume
End Function
