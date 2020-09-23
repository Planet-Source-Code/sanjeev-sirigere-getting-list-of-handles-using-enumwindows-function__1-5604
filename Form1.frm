VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8295
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   435
      Left            =   6270
      TabIndex        =   5
      Top             =   4140
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   3195
      Left            =   4290
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "Form1.frx":0000
      Top             =   570
      Width           =   3945
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   405
      Left            =   210
      TabIndex        =   1
      Top             =   4080
      Width           =   1725
   End
   Begin VB.TextBox Text1 
      Height          =   3195
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Form1.frx":0006
      Top             =   570
      Width           =   3945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "All the window handles"
      Height          =   240
      Left            =   4380
      TabIndex        =   4
      Top             =   150
      Width           =   1980
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Windows With Caption"
      Height          =   240
      Left            =   210
      TabIndex        =   3
      Top             =   150
      Width           =   1980
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Want to help?
'If you've fixed a bug in this code, or added a new feature, or made any other
'improvement, feel free to let me know, and I can put the update on my site
'(giving you credit of course). If you submit an improved version to me it must
'be public domain, just as this version is. You can however change it yourself,
'claim a copyright on your part, and not submit it to me. It's your choice.
'If you find a bug and don't have the knowledge or time to fix it let me know
'and I'll try to fix it. Sending an image file that it can't read would be helpful, as I have tested it on thousands of files without a problem yet.
'Hope you find this code useful,
'
'SANJEEV SIRIGERE
'

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Dim ii As Integer
Dim jj As Integer
Dim strName As String

l = 5
k = 0
EnumWindows AddressOf myproc, l
Text1 = ""
Text2 = ""
jj = 0
For ii = 1 To k
    strName = String(100, Chr$(0))
    GetWindowText j(ii), strName, 100
    strName = Left$(strName, InStr(strName, Chr$(0)) - 1)
    If Len(strName) > 0 Then
        Text1.Text = Text1.Text & "  : " & j(ii) & " : " & strName & vbCrLf
        jj = jj + 1
    End If
    Text2.Text = Text2.Text & j(ii) & vbCrLf
Next
Label1.Caption = Label1.Caption & "(" & jj & ")"
Label2.Caption = Label2.Caption & "(" & k & ")"
End Sub

Private Sub Form_Load()
Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2             'To center the form in the middle of the screen
Command2.Caption = "Fill the text box with Window Text"
Command2.Width = TextWidth(Command2.Caption) + 200
End Sub
