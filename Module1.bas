Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public l As Long
Public k As Long
Public j(1 To 1000) As Long
Public Function myproc(ByVal a As Long) As Boolean
If k < 1000 Then k = k + 1
myproc = True
j(k) = a
End Function
