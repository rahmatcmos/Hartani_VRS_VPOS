VERSION 5.00
Begin VB.Form frmscantbl 
   Caption         =   "scan key"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   1440
      TabIndex        =   8
      Top             =   1920
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "exit"
      Height          =   375
      Left            =   4200
      TabIndex        =   7
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "clear"
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox text3 
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   525
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "char"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Key Ascii"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Key Code"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Test disini"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmscantbl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Text1.Text = ""
    Text2.Text = ""
    text3.Text = ""
    Text4.Text = ""
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
text3.Text = ""
Text4.Text = ""
Text2.Text = Trim(CStr(KeyCode))
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
text3.Text = Trim(CStr(KeyAscii))
Text4.Text = Chr(KeyAscii)
End Sub
