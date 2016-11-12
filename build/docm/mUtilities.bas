Attribute VB_Name = "mUtilities"
'*************************************************************
'   Module mPublicMethods
'   Desc:
'       This module contains public methods!
'*************************************************************
    Option Explicit
    Private Const sModule As String = "mPublicMethods"
    Private sJWTToken As String                                           'in memory bc must not be persistent; it is a session-duration param for GetHTTPResponse

'*************************************************************
'   Public Function Parse(...) As Variant
'   Desc:
'       - Use to pass multiple arguments to forms via OpenArgs in MS Access
'       - Keep multiple arguments in the Tag property of forms and controls.
'       - Use to parse a user-entered search string.
'       - Defaults to using connection string formatted key-value pairs.
'       - Specifying a ReturnType guarantees the type of the result and allows the
'           function to be safely called in certain situations.
'*************************************************************
Public Function Parse(Txt As Variant, _
                Key As String, _
                Optional ReturnType As VbVarType = vbVariant, _
                Optional AssignChar As String = "=", _
                Optional Delimiter As String = ";" _
                ) As Variant

    On Error GoTo Catch
    Dim StartPos As Integer
    Dim EndPos As Integer
    Dim Result As Variant

    Result = Null
    If IsNull(Txt) Then
        Parse = Null
    ElseIf Len(Key) = 0 Then
        EndPos = InStr(Txt, AssignChar)
        If EndPos = 0 Then
            Result = Trim(Txt)
        Else
            If InStrRev(Txt, " ", EndPos) = EndPos - 1 Then
                EndPos = InStrRev(Txt, Delimiter, EndPos - 2)
            Else
                EndPos = InStrRev(Txt, Delimiter, EndPos)
            End If
            Result = Trim(Left(Txt, EndPos))
        End If
    Else
        StartPos = InStr(Txt, Key & AssignChar)
        'Allow for space between Key and Assignment Character
        If StartPos = 0 Then
            StartPos = InStr(Txt, Key & " " & AssignChar)
            If StartPos > 0 Then StartPos = StartPos + Len(Key & " " & AssignChar)
        Else
            StartPos = StartPos + Len(Key & AssignChar)
        End If
        If StartPos = 0 Then
            Parse = Null
        Else
            EndPos = InStr(StartPos, Txt, AssignChar)
            If EndPos = 0 Then
                If Right(Txt, Len(Delimiter)) = Delimiter Then
                    Result = Trim(Mid(Txt, StartPos, _
                                      Len(Txt) - Len(Delimiter) - StartPos + 1))
                Else
                    Result = Trim(Mid(Txt, StartPos))
                End If
            Else
                If InStrRev(Txt, Delimiter, EndPos) = EndPos - 1 Then
                    EndPos = InStrRev(Txt, Delimiter, EndPos - 2)
                Else
                    EndPos = InStrRev(Txt, Delimiter, EndPos)
                End If
                If EndPos < StartPos Then
                    Result = Trim(Mid(Txt, StartPos))
                Else
                    Result = Trim(Mid(Txt, StartPos, EndPos - StartPos))
                End If
            End If

        End If
    End If
    Select Case ReturnType
    Case vbBoolean
        If IsNull(Result) Or Len(Result) = 0 Or Result = "False" Then
            Parse = False
        Else
            Parse = True
            If IsNumeric(Result) Then
                If Val(Result) = 0 Then Parse = False
            End If
        End If

    Case vbCurrency, vbDecimal, vbDouble, vbInteger, vbLong, vbSingle
        If IsNumeric(Result) Then
            Select Case ReturnType
            Case vbCurrency: Parse = CCur(Result)
            Case vbDecimal: Parse = CDec(Result)
            Case vbDouble: Parse = CDbl(Result)
            Case vbInteger: Parse = CInt(Result)
            Case vbLong: Parse = CLng(Result)
            Case vbSingle: Parse = CSng(Result)
            End Select
        Else
            Select Case ReturnType
            Case vbCurrency: Parse = CCur(0)
            Case vbDecimal: Parse = CDec(0)
            Case vbDouble: Parse = CDbl(0)
            Case vbInteger: Parse = CInt(0)
            Case vbLong: Parse = CLng(0)
            Case vbSingle: Parse = CSng(0)
            End Select
        End If

    Case vbDate
        If IsDate(Result) Then
            Parse = CDate(Result)
        ElseIf IsNull(Result) Then
            Parse = 0
        ElseIf IsDate(Replace(Result, "#", "")) Then
            Parse = CDate(Replace(Result, "#", ""))
        Else
            Parse = 0
        End If

    Case vbString
        Parse = Result

    Case Else
        If IsNull(Txt) Then
            Parse = Null
        ElseIf Result = "True" Then
            Parse = True
        ElseIf Result = "False" Then
            Parse = False
        ElseIf IsNumeric(Result) Then
            Parse = Val(Result)
        Else
            Parse = Result
        End If
    End Select

Catch:
End Function

'*************************************************************
'   name:       NullToString
'   description:
'       It will check the first parameter passed and do a nullcheck.
'       If it's a null it will return an empty string by default,
'       but you can change that by passing an optional second argument.
'*************************************************************
Public Function NullToString(Value As Variant, Optional valueIfNull As Variant = "") As Variant
    If VarType(Value) = vbObject Then
        If Value Is Nothing Then NullToString = valueIfNull
    ElseIf VarType(Value) = vbNull Then
        If IsNull(Value) Then NullToString = valueIfNull
    Else
        NullToString = Value
    End If
End Function

'desc: common string parsing algorithm. We use it for node keys, querystrings, and more.
Public Function ParseQueryString(sKey As String, sQueryString) As String
    ParseQueryString = NullToString(Parse(sQueryString, sKey, vbString, "=", "&"))
End Function

'*************************************************************
'   name:       GetJsonValue
'   description:
'       Given a JSON string or response text, find the value of a given key.
'*************************************************************
Public Function GetJsonValue(ByVal JsonString As String, FieldName As String, Optional sDataType As String)
    Dim index1 As Integer
    Dim index2 As Integer
    Dim Value As String
    If sDataType = "number" Or sDataType = "boolean" Then
        index1 = InStr(JsonString, """" & FieldName & """:")
        If index1 = 0 Then
            GetJsonValue = ""
            Exit Function
        End If
        index1 = index1 + Len(FieldName) + 3
        index2 = InStr(index1, JsonString, ",")
    Else
        index1 = InStr(JsonString, """" & FieldName & """:""")
        If index1 = 0 Then
            GetJsonValue = ""
            Exit Function
        End If
        index1 = index1 + Len(FieldName) + 4
        index2 = InStr(index1, JsonString, """")
    End If
    Value = Mid(JsonString, index1, index2 - index1)
    GetJsonValue = Value
End Function

'todo: desc; it takes a hex str and outputs RGB value string
'The input at this point could be HexColor = "#00FF1F"
Public Function HEXCOL2RGB(ByVal HexColor As String) As String
    Dim Red As String
    Dim Green As String
    Dim Blue As String

    HexColor = Replace(HexColor, "#", "")
    Red = Val("&H" & Mid(HexColor, 1, 2))
    Green = Val("&H" & Mid(HexColor, 3, 2))
    Blue = Val("&H" & Mid(HexColor, 5, 2))
    HEXCOL2RGB = RGB(Red, Green, Blue)
End Function

Public Function Max(iNum1 As Integer, iNum2 As Integer) As Integer
    Max = IIf(iNum1 > iNum2, iNum1, iNum2)
End Function

Public Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
  IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function

'read binary file As a string value
Function GetFile(FileName As String) As String
  Dim FileContents() As Byte, FileNumber As Integer
  ReDim FileContents(FileLen(FileName) - 1)
  FileNumber = FreeFile
  Open FileName For Binary As FileNumber
    Get FileNumber, , FileContents
  Close FileNumber
  GetFile = StrConv(FileContents, vbUnicode)
End Function

'*************************************************************
'   Public Function GetFileBytes(ByVal path As String) As Byte()
'   Desc:
'*************************************************************
Public Function GetFileBytes(ByVal Path As String) As Byte()
    Dim lngFileNum As Long
    Dim bytRtnVal() As Byte
    lngFileNum = FreeFile
    If LenB(Dir(Path)) Then ''// Does file exist?
        Open Path For Binary Access Read As lngFileNum
        ReDim bytRtnVal(LOF(lngFileNum) - 1&) As Byte
        Get lngFileNum, , bytRtnVal
        Close lngFileNum
    Else
        Err.Raise 53
    End If
    GetFileBytes = bytRtnVal
    Erase bytRtnVal
End Function

'Delete all files and subfolders
'Be sure that no file is open in the folder
'*************************************************************
'   Public Sub ClearOCTempFolder()
'   Desc:
'*************************************************************
Public Sub ClearTempFolder()
    On Error Resume Next
    Dim FSO As New FileSystemObject

    If FSO.FolderExists(PROP("TempFolder")) = False Then Exit Sub
    FSO.DeleteFile PROP("TempFolder") & "\*.*", True                                'Delete all files
    FSO.DeleteFolder PROP("TempFolder") & "\*.*", True                              'Delete subfolders
End Sub

'todo: desc
'creates OC temp folder(s) if not created already
Public Sub MakeFolders()
    Dim FSO As New FileSystemObject
    If Dir(PROP("TempFolder"), vbDirectory) = "" Then FSO.CreateFolder PROP("TempFolder")
End Sub

'todo: do this better
Public Function DecodeHTMLString(sSourceString) As String
    DecodeHTMLString = Replace(sSourceString, "&amp;", "&")
End Function

'a syntactic sugar for typing json lines in the middle of a json object eg:
'       {
'           ToJsonLine(sKey, sValue)
'       }
'
' If you pass bQuotelessValue = True then it won't wrap the value in quotes. eg if it is a number or bool not str.
Public Function ToJsonLine(sKey As String, sValue As String, Optional ByVal bComma As Boolean = True, Optional ByVal bQuotelessValue As Boolean = False) As String
    ToJsonLine = IIf(bQuotelessValue, sValue, """" & sValue & """")
    ToJsonLine = """" & sKey & """" & ": " & ToJsonLine & IIf(bComma, ",", "")
End Function

'a syntactic sugar for wrapping a json string eg:
'   WrapJson(sOptionalKey, ToJsonLine(sLineKey, sValue))
'
'   instead of
'       sOptionalKey : {
'           sLineKey : sValue
'       }
'
'if no key then pass empty
Public Function WrapJson(sKey As String, sValue As String, Optional ByVal bComma As Boolean = False) As String
    If sKey <> "" Then WrapJson = """" & sKey & """:"
    WrapJson = WrapJson & "{" & sValue & IIf(bComma, """,", "") & "}"
End Function
