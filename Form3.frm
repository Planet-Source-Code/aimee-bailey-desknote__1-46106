VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4455
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Always On Top"
      Height          =   375
      Left            =   1800
      TabIndex        =   17
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Keep On Desktop"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   1800
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFC0&
      Height          =   255
      Index           =   6
      Left            =   2640
      ScaleHeight     =   195
      ScaleWidth      =   1635
      TabIndex        =   14
      Top             =   1440
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   5
      Left            =   2640
      ScaleHeight     =   195
      ScaleWidth      =   1635
      TabIndex        =   12
      Top             =   1080
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   5
      Left            =   2280
      TabIndex        =   13
      Top             =   1080
      Width           =   735
   End
   Begin MSComDlg.CommonDialog ColCMD 
      Left            =   240
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0FF&
      Height          =   255
      Index           =   4
      Left            =   2640
      ScaleHeight     =   195
      ScaleWidth      =   1635
      TabIndex        =   10
      Top             =   720
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   4
      Left            =   2280
      TabIndex        =   11
      Top             =   720
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Index           =   3
      Left            =   2640
      ScaleHeight     =   195
      ScaleWidth      =   1635
      TabIndex        =   8
      Top             =   360
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   3
      Left            =   2280
      TabIndex        =   9
      Top             =   360
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   2
      Left            =   480
      ScaleHeight     =   195
      ScaleWidth      =   1635
      TabIndex        =   6
      Top             =   1080
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0E0FF&
      Height          =   255
      Index           =   1
      Left            =   480
      ScaleHeight     =   195
      ScaleWidth      =   1635
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   0
      Left            =   480
      ScaleHeight     =   195
      ScaleWidth      =   1635
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save && Exit"
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
      Left            =   3240
      TabIndex        =   0
      Top             =   1800
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   6
      Left            =   2280
      TabIndex        =   15
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Title Color"
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
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public owner As Form
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const SWP_NOMOVE = &H2

Private Const SWP_NOSIZE = &H1

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Private Sub Command1_Click()
For i = 0 To 6
If Option1(i).Value = True Then
sendcol Picture1(i).BackColor
End If
Next i
Unload Me
End Sub

Private Sub Command2_Click()
SetWindowPos owner.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
Label2.Caption = "disabled"
End Sub

Public Function sendcol(col As Long)
owner.Label1.BackColor = col
owner.Label2.BackColor = col
owner.Label8.BackColor = col
owner.Label3.BackColor = col
owner.Text1.BackColor = col
End Function

Public Function docol(col As Long)
For i = 0 To 6
If col = Picture1(i).BackColor Then
Option1(i).Value = True
End If
Next i
End Function

Private Sub Picture3_Click(Index As Integer)

End Sub

Private Sub Command3_Click()
SetWindowPos owner.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
Label2.Caption = "enabled"
End Sub

Private Sub Form_Load()
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub
