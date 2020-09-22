VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Save And Exit"
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
      Left            =   2760
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Randomise title colors"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   3855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Load Last 10 Recent Notes On Startup"
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
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "info:  Changes on restart"
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
      Left            =   600
      TabIndex        =   3
      Top             =   720
      Width           =   3975
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.opt1 = Check1.Value
Form1.opt2 = Check2.Value
Form1.SaveSettings
Unload Me
End Sub

Private Sub Form_Load()
Check1.Value = Inte(Form1.opt1)
Check2.Value = Inte(Form1.opt2)
End Sub
Public Function Inte(arg As Boolean) As Integer
If arg = True Then Inte = 1 Else Inte = 0
End Function
