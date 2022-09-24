Attribute VB_Name = "ModulProses"
Option Explicit

' *** Modul Proses Background Virus ***

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Public Const SW_HIDE = 0
Public Const SW_Maximize = 3
Public Const SW_Minimize = 6
Public Const SW_Normal = 1
Public Const SW_SHOW = 5
Public Const WM_CLOSE = &H10

Sub WindowHandle(hWindow, mCase As Long)
Dim X As Long
Select Case mCase
    Case 0
        X = SendMessage(hWindow, WM_CLOSE, 0, 0)
    Case 1
        X = ShowWindow(hWindow, SW_SHOW)
    Case 2
        X = ShowWindow(hWindow, SW_HIDE)
    Case 3
        X = ShowWindow(hWindow, SW_Maximize)
    Case 4
        X = ShowWindow(hWindow, SW_Minimize)
    Case 5
        X = ShowWindow(hWindow, SW_Normal)
End Select
End Sub

Public Function GetWindowTitle(ByVal hWnd As Long) As String
On Error Resume Next
Dim L As Long
Dim S As String

L = GetWindowTextLength(hWnd)
S = Space(L + 1)

GetWindowText hWnd, S, L + 1
GetWindowTitle = Left$(S, L)
End Function

' *** Fungsi untuk mencari window yang hidden dan non hidden ***

Public Sub RefreshDaftarWindow(vForm As Form, vListbox As ListBox)

    Dim RefreshD As Boolean
    Dim i, z, APPCap As Integer
    Dim A As String
    Dim hW As Long
    
    vListbox.Clear
    
    DoEvents
    
    For i = 1 To 10000
        A$ = GetWindowTitle(i)
        z = FindWindow(vbNullString, A$)
        hW = vForm.hWnd
        If z <> 0 Then
            If A$ <> vbNullString And LCase(A$) <> LCase(APPCap) And LCase(A$) <> "FeeLCoMz Destroyer" And i <> hW Then
                If IsWindowEnabled(z) = 0 Then
                    If IsWindowVisible(z) = 0 Then
                        vListbox.AddItem i & vbTab & "[Aktif] " + A$
                    End If
                ElseIf IsWindowEnabled(z) = 1 Then
                    If IsWindowVisible(z) = 0 Then
                        vListbox.AddItem i & vbTab & A$
                    End If
                End If
            End If
        End If
    Next i
    
    DoEvents
    
End Sub

