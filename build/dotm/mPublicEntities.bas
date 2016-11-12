Attribute VB_Name = "mPublicEntities"
'*************************************************************
'   Module mPublicEntities
'   Desc:
'       This module contains OC custom Enums, Types,
'       public variables, public constants, and the
'       Windows API methods we call.
'*************************************************************
'todo: better module name
'todo: SOP for PROP vs Public Var...PROP if it needs to be persistant
    Option Explicit
    Private Const sModule As String = "mPublicEntities"

'Windows API

'idk if it's actually PtrSafe, I just declared it to compile
Public Declare PtrSafe Function SetForegroundWindow _
    Lib "user32" (ByVal hWnd As Long) As LongPtr

Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Public Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#If Win64 Then
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
#Else
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
#End If

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type WINDOWPLACEMENT
    Length As Long
    flags As Long
    showCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECT
End Type

'Public Consts
Public Const GWL_STYLE = (-16)                          'Remove native MS borders
Public Const WS_CAPTION = &HC00000                      'Remove native MS borders: WS_BORDER Or WS_DLGFRAME
Public Const SW_SHOWMINIMIZED   As Integer = 2
Public Const SW_SHOWMAXIMIZED   As Integer = 3
Public Const SW_SHOWNORMAL      As Integer = 1

'Public Variables
Public oPrimaryWindow                       As Window
