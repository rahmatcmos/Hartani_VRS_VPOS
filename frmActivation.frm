VERSION 5.00
Begin VB.Form frmActivation 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ACTIVATION NOW"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Silahkan hubungi Kami di vpointindonesia@yahoo.co.id, Telp : 031-81111918, Cell : 087862405489"
      Top             =   2550
      Width           =   7305
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Versi Demo dapat di gunakan dengan limitasi pemakaian"
      Top             =   2310
      Width           =   7305
   End
   Begin VB.CommandButton cmdDemo 
      Caption         =   "&Demo Version"
      Height          =   405
      Left            =   3930
      TabIndex        =   3
      Top             =   1680
      Width           =   1845
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "frmActivation.frx":0000
      Top             =   450
      Width           =   2205
   End
   Begin VB.TextBox txtActivation 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4020
      TabIndex        =   2
      Top             =   1260
      Width           =   5535
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   405
      Left            =   7710
      TabIndex        =   5
      Top             =   1680
      Width           =   1845
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   405
      Left            =   5820
      TabIndex        =   4
      Top             =   1680
      Width           =   1845
   End
   Begin VB.TextBox txtIDSoftware 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4020
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   5535
   End
   Begin VB.TextBox txtIDHardware 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4020
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   420
      Width           =   5535
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PENDAFTARAN APLIKASI VPOS"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   3195
      TabIndex        =   10
      Top             =   30
      Width           =   3210
   End
   Begin VB.Label Label3 
      Caption         =   "Activation Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2340
      TabIndex        =   8
      Top             =   1290
      Width           =   1725
   End
   Begin VB.Label Label2 
      Caption         =   "ID Software"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2340
      TabIndex        =   7
      Top             =   870
      Width           =   1725
   End
   Begin VB.Label Label1 
      Caption         =   "ID Hardware"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2340
      TabIndex        =   6
      Top             =   450
      Width           =   1725
   End
End
Attribute VB_Name = "frmActivation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xIDHardware As String
Dim xIDSoftware As String
Dim xActivasi As String

Private Sub cmdCancel_Click()
  End
End Sub

Private Sub cmdDemo_Click()
  xActivasi = ""
  Unload Me
End Sub

Private Sub cmdOK_Click()
  If Trim(Replace(txtActivation.Text, "-", "")) = xActivasi Then
    MsgBox "Terima Kasih telah menggunakan produk kami.", vbInformation, App.Title
    xActivasi = Trim(Replace(txtActivation.Text, "-", ""))
    Unload Me
  Else
    MsgBox "Activasi Number salah.", vbInformation, App.Title
  End If
End Sub

Private Sub Form_Load()
Dim i As Integer
On Error GoTo Trace
  txtIDHardware.Enabled = True
  txtIDSoftware.Enabled = True
  For i = 1 To 25 Step 5
    txtIDHardware.Text = txtIDHardware.Text & IIf(i = 1, "", "-") & Mid(xIDHardware, i, 5)
    txtIDSoftware.Text = txtIDSoftware.Text & IIf(i = 1, "", "-") & Mid(xIDSoftware, i, 5)
  Next
Trace:
  If Err.Number <> 0 Then
    MsgBox "Kesalahan : " & Err.Number & " " & Err.Description, vbCritical, App.Title
    Err.Clear
  End If
End Sub

Public Sub Tampil(ByRef IDHardware As String, IDSoftware As String, Activasi As String)
  xActivasi = Activasi
  xIDSoftware = IDSoftware
  xIDHardware = IDHardware
  Me.Show 1
  Activasi = xActivasi
  IDSoftware = xIDSoftware
  IDHardware = xIDHardware
End Sub
