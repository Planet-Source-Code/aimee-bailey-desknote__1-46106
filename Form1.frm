VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DeskNote"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4920
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "About http://"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start DeskNote"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00C0FFFF&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   4860
      TabIndex        =   0
      Top             =   0
      Width           =   4920
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "By Steve Bailey"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   645
         Width           =   4695
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DeskNote v1"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   4695
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sfile As String
Public nocom As Boolean
Public opt1 As Boolean
Public opt2 As Boolean
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const SWP_NOMOVE = &H2

Private Const SWP_NOSIZE = &H1

Private Const HWND_TOPMOST = -1
Private Sub Command1_Click()
note.Show
AutoFormShape note, RGB(255, 0, 255)
note.Left = Screen.Width - 3315 - 1000
note.opt1 = opt1
note.opt2 = opt2
Unload Me
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Shell "c:\windows\explorer.exe " & Chr(34) & "http://www.hybrid-factor.co.uk/dn/index.htm" & Chr(34), vbMaximizedFocus
End Sub

Public Function LoadSettings()
'On Error GoTo err
If Right(App.Path, 1) <> "\" Then x = "\"
Open App.Path & x & "desknote.ini" For Input As #1
Input #1, a$
opt1 = a$
Input #1, a$
opt2 = a$
err:
Close #1
End Function
Public Function SaveSettings()
'On Error Resume Next
If Right(App.Path, 1) <> "\" Then x = "\"
Open App.Path & x & "desknote.ini" For Output As #1
Print #1, opt1
Print #1, opt2
Close #1
'MsgBox opt1 & " " & opt2
End Function

Private Sub Command4_Click()
Form4.Show
SetWindowPos Form4.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_Load()
Dim i As Integer
'On Error GoTo err
LoadSettings
If nocom = True Then GoTo aa
If Trim(Mid(LCase(Command$), 1, 2)) = "-s" Then
sfile = Replace(Trim(Mid(Trim(Command$), 3, 255)), Chr(34), "")
If sfile = "" Then MsgBox "Sfile was called but not used!", vbInformation: GoTo err
Close #3
Open sfile For Input As #3
Do Until EOF(3)
DoEvents
Input #3, a$
i = i + 1
NewNote a$, i, False
Loop
Close #3


MsgBox sfile & " was successfully loaded!"
Unload Me
ElseIf Trim(Mid(LCase(Command$), 1, 2)) = "-f" Then
sfile = Replace(Trim(Mid(Trim(Command$), 3, 255)), Chr(34), "")
NewNote sfile, 1, False
Else


If opt1 = True Then LoadRecent


End If
GoTo aa:
err:
MsgBox err.Description & " -- form_load"
aa:
Close #3
End Sub

Public Function LoadSFile(file As String, rndcol As Boolean)
Dim i As Integer
Open file For Input As #2
Do Until EOF(2)
DoEvents
Input #2, a$
i = i + 1
NewNote a$, i, rndcol
Loop
Close #2
End Function

Public Function LoadRecent()
Dim v As Integer
If Right(App.Path, 1) <> "\" Then x = "\" Else x = ""
Open App.Path & x & "recent.nlf" For Input As #15
Do Until EOF(15)
DoEvents

v = v + 1
Input #15, a$
NewNote a$, v, opt2
If v = 10 Then GoTo aa
Loop
aa:
Close #15

End Function

Public Function test()
Dim i As Integer
sfile = "c:\notes\n.nlf"
Open sfile For Input As #2
Do Until EOF(2)
DoEvents
Input #2, a$
i = i + 1
NewNote a$, i, False
Loop
Close #2
End Function

Public Function NewNote(file As String, Index As Integer, rndcol As Boolean)
Dim vv As Long
Debug.Print file & " " & Index & " " & rndcol
On Error Resume Next
Dim x As New note
x.Visible = True
x.Left = Me.Left
x.Top = Me.Top + Me.Height
AutoFormShape x, RGB(255, 0, 255)
x.Label2.Caption = "#" & Format(Int(Mid(x.Label2.Caption, 2, 255)) + 1, "00")
x.Top = Me.Top + (460 * Index)
x.Left = Screen.Width - 3315 - 1000
Close #1
x.Text1.Text = ""
Open file For Input As #1
Do Until EOF(1)
Input #1, a$
Debug.Print a$
If Mid(a$, 1, 3) = "<t>" Then
    xx = Replace(a$, "<t>", "")
    xx = Replace(xx, "</t>", "")
    x.Label1.Caption = xx
ElseIf Mid(a$, 1, 3) = "<d>" Then
    xx = Replace(a$, "<d>", "")
    xx = Replace(xx, "</d>", "")
    Rdate = xx
    x.Label7.Caption = Rdate
Else
    x.Text1.Text = x.Text1.Text & a$ & vbCrLf
End If

If rndcol = True Then
vv = GetRNDC
    x.Label1.BackColor = vv
    x.Label2.BackColor = vv
    x.Label8.BackColor = vv
    x.Label3.BackColor = vv
    x.Text1.BackColor = vv
End If

x.Minimize
Loop
GoTo aa
err:
MsgBox err.Description & " -- NewNote"
aa:
Close #1

End Function

Public Function GetRNDC() As Long
Dim x As Integer
Randomize
x = Rnd * 6
Select Case x
Case 0: GetRNDC = &HC0FFFF
Case 1: GetRNDC = &HC0E0FF
Case 2: GetRNDC = &HC0C0FF
Case 3: GetRNDC = &HFFC0C0
Case 4: GetRNDC = &HFFC0FF
Case 5: GetRNDC = &HC0FFC0
Case 6: GetRNDC = &HFFFFC0
End Select
End Function
