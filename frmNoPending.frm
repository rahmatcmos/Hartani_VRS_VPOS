VERSION 5.00
Begin VB.Form frmPesan 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   ScaleHeight     =   1500
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3630
      Top             =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1170
      TabIndex        =   0
      Top             =   1680
      Width           =   1395
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   1245
      Left            =   150
      Top             =   120
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   1365
      Left            =   90
      Top             =   60
      Width           =   4575
   End
   Begin VB.Label lbPesan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tidak ada Transaksi Pending"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   -60
      TabIndex        =   1
      Top             =   570
      Width           =   4605
   End
End
Attribute VB_Name = "frmPesan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub Form_Activate()
Timer1.interval = 9000
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Unload Me
End Sub
