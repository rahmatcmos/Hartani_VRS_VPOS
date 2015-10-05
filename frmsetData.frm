VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSetData 
   Caption         =   "SETTING DATA"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   8625
   Icon            =   "frmsetData.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5025
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   2610
      TabIndex        =   18
      Top             =   3510
      Width           =   5325
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   2610
      TabIndex        =   10
      Top             =   1800
      Width           =   5325
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   2610
      TabIndex        =   1
      Top             =   120
      Width           =   5325
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   2610
      TabIndex        =   7
      Top             =   1230
      Width           =   5325
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   2610
      TabIndex        =   4
      Top             =   660
      Width           =   5325
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   2610
      TabIndex        =   13
      Top             =   2370
      Width           =   5325
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   2610
      TabIndex        =   15
      Top             =   2940
      Width           =   5325
   End
   Begin VB.CommandButton Command6 
      Caption         =   "..."
      Height          =   465
      Left            =   7950
      TabIndex        =   11
      Top             =   1830
      Width           =   465
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4620
      Top             =   3090
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&TUTUP"
      Height          =   615
      Left            =   7110
      TabIndex        =   17
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&SIMPAN"
      Height          =   615
      Left            =   5580
      TabIndex        =   16
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   465
      Left            =   7950
      TabIndex        =   5
      Top             =   690
      Width           =   465
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   465
      Left            =   7950
      TabIndex        =   8
      Top             =   1260
      Width           =   465
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   465
      Left            =   7950
      TabIndex        =   2
      Top             =   150
      Width           =   465
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "AUTO COMPACT (Y/T)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   60
      TabIndex        =   19
      Top             =   3600
      Width           =   2595
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "BACKUP 2 "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   60
      TabIndex        =   9
      Top             =   1890
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "NO ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   60
      TabIndex        =   14
      Top             =   3030
      Width           =   2505
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "KODE PENGAWAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   60
      TabIndex        =   12
      Top             =   2460
      Width           =   2505
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "BACKUP 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   60
      TabIndex        =   3
      Top             =   750
      Width           =   1275
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "RESET 2 (Lokal)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   60
      TabIndex        =   6
      Top             =   1290
      Width           =   2385
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "RESET 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   60
      TabIndex        =   0
      Top             =   150
      Width           =   1185
   End
End
Attribute VB_Name = "frmSetData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    CommonDialog1.ShowSave
    Text1.Text = DirSaja(CommonDialog1.FileName)
End Sub

Private Sub Command2_Click()
    CommonDialog1.ShowSave
    Text2.Text = DirSaja(CommonDialog1.FileName)
End Sub

Private Sub Command3_Click()
    CommonDialog1.ShowSave
    Text3.Text = DirSaja(CommonDialog1.FileName)
End Sub

Private Sub Command4_Click()
    SetRegistry "Reset1", Text1.Text, "Data"
    SetRegistry "Reset2", Text2.Text, "Data"
    SetRegistry "Backup", Text3.Text, "Data"
    SetRegistry "Backup2", Text6.Text, "Data"
    SetRegistry "Kode", Text4.Text, "Pengawas"
    SetRegistry "Prosen", Text5.Text, "Pengawas"
    SetRegistry "AutoDelete", Text7.Text, "Data"
    MsgBox "Setting sudah tersimpan!", vbOKOnly + vbInformation, "SET DATA"
End Sub

Private Sub Command5_Click()
    Unload Me
End Sub

Private Sub Command6_Click()
    CommonDialog1.ShowSave
    Text6.Text = DirSaja(CommonDialog1.FileName)
End Sub

Private Sub Form_Load()
    Text1.Text = getRegistry("Reset1", "Data")
    Text2.Text = getRegistry("Reset2", "Data")
    Text3.Text = getRegistry("Backup", "Data")
    Text6.Text = getRegistry("Backup2", "Data")
    Text7.Text = getRegistry("AutoDelete", "Data")
    Text4.Text = getRegistry("Kode", "Pengawas")
    Text5.Text = getRegistry("Prosen", "Pengawas")
End Sub

