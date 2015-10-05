VERSION 5.00
Begin VB.Form frmGetSetial 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   555
      Left            =   360
      TabIndex        =   2
      Top             =   2190
      Width           =   1875
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   2280
      TabIndex        =   1
      Top             =   480
      Width           =   1995
   End
   Begin VB.CommandButton Command1 
      Height          =   465
      Left            =   390
      TabIndex        =   0
      Top             =   1140
      Width           =   1095
   End
End
Attribute VB_Name = "frmGetSetial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x As New clsSerial
Private Sub Command1_Click()
On Error Resume Next
  Text1.Text = Replace(x.MBSerialNumber, " ", "") & "-" & Replace(x.GetCpuID, " ", "")
End Sub

