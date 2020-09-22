VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form note 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0080FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3315
   ForeColor       =   &H8000000B&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   1845
   ScaleWidth      =   3315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CMD 
      Left            =   1680
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "Form2.frx":B852
      Top             =   480
      Width           =   2895
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "&Options"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   165
      Left            =   1440
      TabIndex        =   9
      ToolTipText     =   "Change some options for this note"
      Top             =   1440
      Width           =   540
   End
   Begin VB.Line Line4 
      X1              =   3120
      X2              =   3120
      Y1              =   0
      Y2              =   720
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2960
      TabIndex        =   2
      ToolTipText     =   "close this note"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2800
      TabIndex        =   8
      ToolTipText     =   "Back to main menu!"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/03"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   165
      Left            =   2280
      TabIndex        =   7
      ToolTipText     =   "the current date"
      Top             =   1440
      Width           =   780
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   165
      Left            =   960
      TabIndex        =   6
      ToolTipText     =   "Save this note to file"
      Top             =   1440
      Width           =   420
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "&Open"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   165
      Left            =   480
      TabIndex        =   5
      ToolTipText     =   "Open a note from file"
      Top             =   1440
      Width           =   420
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   160
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Create a new note"
      Top             =   1440
      Width           =   300
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "[DBL Click To Edit]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      MousePointer    =   15  'Size All
      TabIndex        =   0
      ToolTipText     =   "The title of the note"
      Top             =   120
      Width           =   2775
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   2520
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   0
      Y1              =   240
      Y2              =   960
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   960
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "#01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "(tecky stuff) the window index"
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "note"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const SWP_NOMOVE = &H2

Private Const SWP_NOSIZE = &H1

Private Const HWND_TOPMOST = -1

Dim CurRgn, TempRgn As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private MouseDownForm
Private MouseDownFormX
Private MouseDownFormY
Private Type POINTAPI
    x As Long
    y As Long
End Type
Dim oldx, oldy As Integer
Dim Rdate As String
Public opt1 As Boolean
Public opt2 As Boolean

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
Rdate = Realdate
Label7.Caption = Realdate
End Sub

Private Sub form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ReleaseCapture  ' This releases the mouse communication with the form so it can communicate with the operating system to move the form
Result& = SendMessage(Me.hwnd, &H112, &HF012, 0)  ' This tells the OS to pick up the form to be moved

End Sub

Private Sub Label1_DblClick()
Dim x As String
On Error GoTo err
x = InputBox("Please enter new title", "Edit Title", Label1.Caption)
If Trim(x) = "" Then x = "[DBL Click To Edit]"
Label1.Caption = x
err:
End Sub

Private Sub label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

MouseDownForm = 1
    MouseDownFormX = x + Label1.Left
    MouseDownFormY = y + Label1.Top
    
End Sub

Private Sub label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim z As POINTAPI
Call GetCursorPos(z)
If Button = 1 Then

        Me.Top = (z.y * 15) - MouseDownFormY
        Me.Left = (z.x * 15) - MouseDownFormX
End If
End Sub




Private Sub label1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
MouseDownForm = 0

If Button = 2 Then

If Me.Height = 1845 Then
    Me.Height = 500
    Text1.Visible = False
    Line3.Y1 = Me.Height - 22: Line3.Y2 = Me.Height - 22: Line3.X2 = Me.Width
Else
    Me.Height = 1845
    Text1.Visible = True
    Line3.Y1 = Me.Height - 12: Line3.Y2 = Me.Height - 12: Line3.X2 = Me.Width
End If

End If

End Sub
Public Function Minimize()
    Me.Height = 500
    Text1.Visible = False
    Line3.Y1 = Me.Height - 22: Line3.Y2 = Me.Height - 22: Line3.X2 = Me.Width
End Function
Private Sub Form_Resize()
On Error Resume Next

Line1.X2 = Me.Width
Line2.Y2 = Me.Height
Line3.Y1 = Me.Height - 12: Line3.Y2 = Me.Height - 12: Line3.X2 = Me.Width
Line4.X1 = Me.Width - 10: Line4.X2 = Me.Width - 10: Line4.Y2 = Me.Height
Label1.Width = Me.Width
Label3.Left = Line4.X1 - 140
Text1.Width = Me.Width - 10 - 240
Text1.Height = Me.Height - Text1.Top - 130 - Label4.Height
Label4.Top = Text1.Top + Text1.Height
Label5.Top = Label4.Top
Label6.Top = Label4.Top
Label7.Top = Label4.Top
Label7.Left = Text1.Left + Text1.Width - Label7.Width
Label8.Left = Label3.Left - 160
Label9.Top = Label5.Top
End Sub

Private Sub Label3_Click()
Unload Me
End Sub

Private Sub Label4_Click()
Dim x As New note
Dim vv As Long
x.Visible = True
x.Left = Me.Left
x.Top = Me.Top + Me.Height
AutoFormShape x, RGB(255, 0, 255)
x.Label2.Caption = "#" & Format(Int(Mid(Me.Label2.Caption, 2, 255)) + 1, "00")
x.Top = Me.Top + 600
x.opt1 = opt1
x.opt2 = opt2
If opt2 = True Then
vv = GetRNDC
    x.Label1.BackColor = vv
    x.Label2.BackColor = vv
    x.Label8.BackColor = vv
    x.Label3.BackColor = vv
    x.Text1.BackColor = vv
End If

End Sub

Private Sub Label5_Click()
CMD.Filter = "Note File(s) [*.note]|*.note|Text Files [*.txt]|*.txt*"
CMD.ShowOpen
Reset
Text1.Text = ""
Open CMD.FileName For Input As #3
Do Until EOF(3)
Input #3, a$
Debug.Print a$
If Mid(a$, 1, 3) = "<t>" Then
    xx = Replace(a$, "<t>", "")
    xx = Replace(xx, "</t>", "")
    Label1.Caption = xx
ElseIf Mid(a$, 1, 3) = "<d>" Then
    xx = Replace(a$, "<d>", "")
    xx = Replace(xx, "</d>", "")
    Rdate = xx
    Label7.Caption = Rdate
Else
    Text1.Text = Text1.Text & a$ & vbCrLf
End If
    
Loop
Close #3
End Sub

Private Sub Label6_Click()
CMD.Filter = "Note File(s) [*.note]|*.note|Text Files [*.txt]|*.txt*"
CMD.ShowSave
LogNote CMD.FileName
On Error GoTo err
Open CMD.FileName For Output As #11
x = "<t>" & Label1.Caption & "</t>" & vbCrLf
x = "<d>" & Label7.Caption & "</d>" & vbCrLf
x = x & Text1.Text
Print #11, x

err:
Close #11
End Sub

Public Function Realdate() As String
Dim x As String
Dim m, d, y As String
x = Date$
m = Mid(x, 1, 2)
d = Mid(x, 4, 2)
y = Right(x, 4)
Realdate = d & "/" & m & "/" & y
End Function

Public Function LogNote(name_ta As String)
MsgBox name_ta
'on error goto err
List1.Clear
If Right(App.Path, 1) <> "\" Then x = "\" Else x = ""
FileName = App.Path & x & "recent.nlf"
If name_ta = "" Then GoTo err:
Open FileName For Output As #6
Do Until EOF(6)
DoEvents
Input #6, a$
List1.AddItem a$
Loop
Close #6
List1.AddItem name_ta
Open FileName For Output As #6
For i = 0 To List1.ListCount - 1
Print #6, List1.List(i)
Next i
Close #6

err:

End Function

Private Sub Label8_Click()
Form1.nocom = True
Form1.Show
Form1.nocom = True
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Label9_Click()
Form3.Show
Set Form3.owner = Me
Form3.docol Label1.BackColor
End Sub
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

