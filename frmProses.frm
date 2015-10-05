VERSION 5.00
Begin VB.Form frmProses 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1485
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2430
      Top             =   180
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sedang proses copy data dari server"
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
      Left            =   -30
      TabIndex        =   1
      Top             =   720
      Width           =   4605
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   1245
      Left            =   120
      Top             =   120
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   1365
      Left            =   60
      Top             =   60
      Width           =   4575
   End
   Begin VB.Label lbTanya 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tunggu................................................"
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
      Left            =   -30
      TabIndex        =   0
      Top             =   330
      Width           =   4605
   End
End
Attribute VB_Name = "frmProses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
'Dim db As Access.Application
On Error GoTo Akhir
  If (Dir(DirUpdate & "\temp.txt") <> "") Then
'    Set db = New Access.Application
'    db.DBEngine.CompactDatabase DirDatabase & "\dbmaster.mdb", DirDatabase & "\dbmasterNew.mdb"
'    Kill DirDatabase & "\dbmaster.mdb"
'    Name DirDatabase & "\dbmasterNew.mdb" As DirDatabase & "\dbmaster.mdb"
'    db.Quit acQuitSaveNone
'    Set db = Nothing
    FileCopy DirUpdate & "\dbmaster.mdb", DirDatabase & "\dbmaster.mdb"
    Kill DirUpdate & "\temp.txt"
    Kill DirUpdate & "\dbmaster.mdb"
  End If
DoEvents
Unload Me
Exit Sub
Akhir:
  Unload Me
'  db.Quit acQuitSaveNone
'  Set db = Nothing
End Sub
