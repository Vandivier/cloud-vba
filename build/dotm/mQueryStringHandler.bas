Attribute VB_Name = "mQueryStringHandler"
'*************************************************************
' Procedure : Parse
' DateTime  : 7/16/2009 11:32
' Author    : Mike
' Purpose   : Parse a string of keys and values (such as a connection string) and return
'               the value of a specific key.
' Usage     - Use to pass multiple arguments to forms via OpenArgs in MS Access
'           - Keep multiple arguments in the Tag property of forms and controls.
'           - Use to parse a user-entered search string.
' Notes     - Defaults to using connection string formatted key-value pairs.
'           - Specifying a ReturnType guarantees the type of the result and allows the
'               function to be safely called in certain situations.
'*************************************************************
Option Explicit
Private Const sModule As String = "mQueryStringHandler"

Public Function Parse( _
            Txt As Variant, _
            Key As String, _
            Optional ReturnType As VbVarType = vbVariant, _
            Optional AssignChar As String = "=", _
            Optional Delimiter As String = ";" _
            ) As Variant
    On Error GoTo Catch

Dim StartPos As Integer, EndPos As Integer, Result As Variant
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
                If val(Result) = 0 Then Parse = False
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
        'Parse = Nz(Result, vbNullString)
        Parse = Result

    Case Else
        If IsNull(Txt) Then
            Parse = Null
        ElseIf Result = "True" Then
            Parse = True
        ElseIf Result = "False" Then
            Parse = False
        ElseIf IsNumeric(Result) Then
            Parse = val(Result)
        Else
            Parse = Result
        End If
    End Select
CleanExit:
    Exit Function
Catch:
    ErrorHandler Err, sModule, "Parse"
    Resume CleanExit
    Resume
End Function

'desc: common string parsing algorithm. We use it for node keys, querystrings, and more.
Public Function ParseQueryString(ByVal sKey As String, ByVal sQueryString As String) As String
    ParseQueryString = NullToString(Parse(sQueryString, sKey, vbString, "=", "&"))
End Function
