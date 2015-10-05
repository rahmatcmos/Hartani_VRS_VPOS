VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Begin VB.Form frmReset 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4005
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmReset.frx":0000
      Left            =   1290
      List            =   "frmReset.frx":002E
      TabIndex        =   30
      Top             =   330
      Width           =   1035
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmReset.frx":006D
      Left            =   3900
      List            =   "frmReset.frx":007A
      TabIndex        =   29
      Top             =   330
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PROSES"
      Height          =   675
      Left            =   5760
      TabIndex        =   28
      Top             =   420
      Width           =   1035
   End
   Begin VB.Data DataLokalDTL 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   405
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data DataLokal 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   405
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      IMEMode         =   3  'DISABLE
      Left            =   210
      TabIndex        =   24
      Text            =   "\\Kassa02\pos\Database"
      Top             =   3180
      Width           =   6645
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      IMEMode         =   3  'DISABLE
      Left            =   3810
      TabIndex        =   22
      Top             =   2460
      Width           =   1275
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      IMEMode         =   3  'DISABLE
      Left            =   3810
      TabIndex        =   20
      Top             =   1920
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      IMEMode         =   3  'DISABLE
      Left            =   3810
      TabIndex        =   0
      Top             =   1380
      Width           =   1275
   End
   Begin TDBDate6Ctl.TDBDate TDBDate1 
      Height          =   315
      Left            =   300
      TabIndex        =   31
      Top             =   780
      Width           =   2595
      _Version        =   65536
      _ExtentX        =   4577
      _ExtentY        =   556
      Calendar        =   "frmReset.frx":0098
      Caption         =   "frmReset.frx":01B0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmReset.frx":0213
      Keys            =   "frmReset.frx":0231
      Spin            =   "frmReset.frx":028F
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd/mm/yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "dd/mm/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "05/03/2010"
      ValidateMode    =   0
      ValueVT         =   1414725639
      Value           =   40242
      CenturyMode     =   0
   End
   Begin TDBDate6Ctl.TDBDate TDBDate2 
      Height          =   315
      Left            =   2970
      TabIndex        =   32
      Top             =   780
      Width           =   2595
      _Version        =   65536
      _ExtentX        =   4577
      _ExtentY        =   556
      Calendar        =   "frmReset.frx":02B7
      Caption         =   "frmReset.frx":03CF
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmReset.frx":0430
      Keys            =   "frmReset.frx":044E
      Spin            =   "frmReset.frx":04AC
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd/mm/yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "dd/mm/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "05/03/2010"
      ValidateMode    =   0
      ValueVT         =   1414725639
      Value           =   40242
      CenturyMode     =   0
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Kassa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   270
      TabIndex        =   34
      Top             =   390
      Width           =   945
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Shift"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2940
      TabIndex        =   33
      Top             =   360
      Width           =   945
   End
   Begin VB.Label lblNama 
      BackStyle       =   0  'Transparent
      Caption         =   "Mesin-Kasir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   9
      Left            =   120
      TabIndex        =   27
      Top             =   5490
      Width           =   2205
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Mesin-Kasir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   9
      Left            =   120
      TabIndex        =   26
      Top             =   5130
      Width           =   3405
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "LETAK DIREKTORI POS:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   210
      TabIndex        =   25
      Top             =   2880
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "No Mesin (2 angka)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   210
      TabIndex        =   23
      Top             =   2520
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal-Bulan-Tahun (ddMMyyyy)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   210
      TabIndex        =   21
      Top             =   1980
      Width           =   3735
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Masukkan SHIFT (1 atau 2)                Cetak <ENTER>, Keluar <ESC>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   210
      TabIndex        =   19
      Top             =   1380
      Width           =   3735
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Transaksi Pending"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   8
      Left            =   11430
      TabIndex        =   18
      Top             =   3870
      Width           =   3405
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Bank"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   7
      Left            =   11430
      TabIndex        =   17
      Top             =   3480
      Width           =   3405
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Tunai"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   6
      Left            =   11430
      TabIndex        =   16
      Top             =   3120
      Width           =   3405
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   5
      Left            =   11430
      TabIndex        =   15
      Top             =   2760
      Width           =   3405
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Diskon"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1305
      Index           =   4
      Left            =   11430
      TabIndex        =   14
      Top             =   1920
      Width           =   7365
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah SubTotal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   3
      Left            =   11430
      TabIndex        =   13
      Top             =   1560
      Width           =   3405
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Nota"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   2
      Left            =   11430
      TabIndex        =   12
      Top             =   1200
      Width           =   3405
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Mesin-Kasir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   1
      Left            =   11430
      TabIndex        =   11
      Top             =   840
      Width           =   3405
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   0
      Left            =   11430
      TabIndex        =   10
      Top             =   480
      Width           =   3405
   End
   Begin VB.Label lblNama 
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   8
      Left            =   9450
      TabIndex        =   9
      Top             =   3840
      Width           =   2205
   End
   Begin VB.Label lblNama 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Bank"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   7
      Left            =   9450
      TabIndex        =   8
      Top             =   3480
      Width           =   2205
   End
   Begin VB.Label lblNama 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Tunai"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   6
      Left            =   9450
      TabIndex        =   7
      Top             =   3120
      Width           =   2205
   End
   Begin VB.Label lblNama 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   5
      Left            =   9450
      TabIndex        =   6
      Top             =   2760
      Width           =   2205
   End
   Begin VB.Label lblNama 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Diskon"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   4
      Left            =   9450
      TabIndex        =   5
      Top             =   1920
      Width           =   2205
   End
   Begin VB.Label lblNama 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah SubTotal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   3
      Left            =   9450
      TabIndex        =   4
      Top             =   1560
      Width           =   2205
   End
   Begin VB.Label lblNama 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Nota"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   2
      Left            =   9450
      TabIndex        =   3
      Top             =   1200
      Width           =   2205
   End
   Begin VB.Label lblNama 
      BackStyle       =   0  'Transparent
      Caption         =   "Mesin-Kasir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   1
      Left            =   9450
      TabIndex        =   2
      Top             =   840
      Width           =   2205
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   3585
      Left            =   150
      Top             =   270
      Width           =   6795
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   3735
      Left            =   60
      Top             =   180
      Width           =   6915
   End
   Begin VB.Label lblNama 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   0
      Left            =   9450
      TabIndex        =   1
      Top             =   480
      Width           =   2205
   End
End
Attribute VB_Name = "frmReset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim jawab As Boolean
Dim tanya As String
Dim dbs As Database
Dim rs As Recordset
Dim KeyOk As Integer
Dim isDataOK As Boolean
Dim NamaShiftReset As Integer
Dim namaMesinReset As String
Dim DIRPOS As String
Dim TglReset As String
  Dim DayReset As Integer
  Dim MonthReset As Integer
  Dim YearReset As Integer
Dim KodeKasirReset As String
    Dim namakasirReset As String
Private Sub Command1_Click()
On Error GoTo pesan
Dim ks As Integer
If Combo1.Text <> "Semua" Then
    Dim i As Integer
    Dim tgl As Date
    For i = 0 To (DateDiff("d", TDBDate1, TDBDate2))
    tgl = DateAdd("d", i, TDBDate1)
    Text2.Text = Format(tgl, "ddMMyyyy")
    DoEvents
    If Combo2.ListIndex = 0 Or Combo2.ListIndex = 1 Then
        Text1.Text = 1
        NamaShiftReset = 1
        DoEvents
        Text3.Text = Combo1.Text
        DoEvents
        namaMesinReset = Text3.Text
        DoEvents
        Text4.Text = "\\KASSA" & Format(Val(namaMesinReset), "0#") & "\pos"
        Text4.Locked = True
        If Text4.Text = "" Then Exit Sub
        If Dir(Text4.Text & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb") = "" Then
          frmPesan.lbPesan = "MESIN TIDAK ONLINE!!"
          frmPesan.Show 1
          Text4.Locked = False
          Exit Sub
        End If
        DoEvents
        DIRPOS = Text4.Text
        DayReset = Format(tgl, "dd") 'Val(Mid(Text2.Text, 1, 2))
        MonthReset = Format(tgl, "MM") ' Val(Mid(Text2.Text, 3, 2))
        YearReset = Format(tgl, "yyyy") 'Val(Mid(Text2.Text, 5, 4))
        DoEvents
        AmbilPersenPerMesin
'        If Text3.Text = "09" Or Text3.Text = "10" Or Text3.Text = "20" Or Text3.Text = "99" Then
'            CekTransaksiBermasalahBusana
'        Else
            CekTransaksiBermasalah
'        End If
        If isDataOK Then
            Screen.MousePointer = vbHourglass
            If Text3.Text = "09" Or Text3.Text = "10" Or Text3.Text = "20" Or Text3.Text = "99" Then
                ResetBusana
            Else
                ResetRetail
            End If
            ResendServerOnline
          
          If (UCase(Trim(getRegistry("AutoDelete", "Data"))) = "Y") And (KodeKasir = KodeUserDua) Then
            HAPUSTRANSAKSI
          End If
          Screen.MousePointer = vbDefault
        Else
    '      frmPesan.lbPesan = "Ada Transaksi Pending!!!!"
            MsgBox "Ada Transaksi Pending pada tanggal: " & Format(tgl, "dd-MM-yyyy") & "!!!!"
        End If
    End If
    DoEvents
    If Combo2.ListIndex = 0 Or Combo2.ListIndex = 2 Then
        Text1.Text = 2
        NamaShiftReset = 2
        DoEvents
        Text3.Text = Combo1.Text
        DoEvents
        namaMesinReset = Text3.Text
        DoEvents
        Text4.Text = "\\KASSA" & Format(Val(namaMesinReset), "0#") & "\pos"
        Text4.Locked = True
        If Text4.Text = "" Then Exit Sub
        If Dir(Text4.Text & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb") = "" Then
          frmPesan.lbPesan = "MESIN TIDAK ONLINE!!"
          frmPesan.Show 1
          Text4.Locked = False
          Exit Sub
        End If
        DoEvents
        DIRPOS = Text4.Text
        DayReset = Format(tgl, "dd") 'Val(Mid(Text2.Text, 1, 2))
        MonthReset = Format(tgl, "MM") ' Val(Mid(Text2.Text, 3, 2))
        YearReset = Format(tgl, "yyyy") 'Val(Mid(Text2.Text, 5, 4))
        DoEvents
        AmbilPersenPerMesin
        If Text3.Text = "09" Or Text3.Text = "10" Or Text3.Text = "20" Or Text3.Text = "99" Then
            CekTransaksiBermasalahBusana
        Else
            CekTransaksiBermasalah
        End If
        If isDataOK Then
            Screen.MousePointer = vbHourglass
            If Text3.Text = "09" Or Text3.Text = "10" Or Text3.Text = "20" Or Text3.Text = "99" Then
                ResetBusana
            Else
                ResetRetail
            End If
            ResendServerOnline
          
          If (UCase(Trim(getRegistry("AutoDelete", "Data"))) = "Y") And (KodeKasir = KodeUserDua) Then
            HAPUSTRANSAKSI
          End If
          Screen.MousePointer = vbDefault
        Else
    '      frmPesan.lbPesan = "Ada Transaksi Pending!!!!"
            MsgBox "Ada Transaksi Pending pada tanggal: " & Format(tgl, "dd-MM-yyyy") & "!!!!"
        End If
    End If
    Next
    MsgBox "RESET SELESAI!", vbOKOnly + vbInformation
Else 'Semua kasir
For ks = 1 To 13
DoEvents
Combo1.Text = Format(ks, "00")
DoEvents
    For i = 0 To (DateDiff("d", TDBDate1, TDBDate2))
    tgl = DateAdd("d", i, TDBDate1)
    Text2.Text = Format(tgl, "ddMMyyyy")
    DoEvents
    If Combo2.ListIndex = 0 Or Combo2.ListIndex = 1 Then
        Text1.Text = 1
        NamaShiftReset = 1
        DoEvents
        Text3.Text = Format(ks, "00")
        DoEvents
        namaMesinReset = Text3.Text
        DoEvents
        Text4.Text = "\\KASSA" & Format(Val(namaMesinReset), "0#") & "\pos"
        Text4.Locked = True
        If Text4.Text = "" Then Exit Sub
        If Dir(Text4.Text & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb") = "" Then
          frmPesan.lbPesan = "MESIN TIDAK ONLINE!!"
          frmPesan.Show 1
          Text4.Locked = False
          Exit Sub
        End If
        DoEvents
        DIRPOS = Text4.Text
        DayReset = Format(tgl, "dd") 'Val(Mid(Text2.Text, 1, 2))
        MonthReset = Format(tgl, "MM") ' Val(Mid(Text2.Text, 3, 2))
        YearReset = Format(tgl, "yyyy") 'Val(Mid(Text2.Text, 5, 4))
        DoEvents
        AmbilPersenPerMesin
        If Text3.Text = "09" Or Text3.Text = "10" Or Text3.Text = "20" Or Text3.Text = "99" Then
            CekTransaksiBermasalahBusana
        Else
            CekTransaksiBermasalah
        End If
        If isDataOK Then
            Screen.MousePointer = vbHourglass
            If Text3.Text = "09" Or Text3.Text = "10" Or Text3.Text = "20" Or Text3.Text = "99" Then
                ResetBusana
            Else
                ResetRetail
            End If
            ResendServerOnline
          
          If (UCase(Trim(getRegistry("AutoDelete", "Data"))) = "Y") And (KodeKasir = KodeUserDua) Then
            HAPUSTRANSAKSI
          End If
          Screen.MousePointer = vbDefault
        Else
    '      frmPesan.lbPesan = "Ada Transaksi Pending!!!!"
            MsgBox "Ada Transaksi Pending pada tanggal: " & Format(tgl, "dd-MM-yyyy") & "!!!!"
        End If
    End If
    DoEvents
    If Combo2.ListIndex = 0 Or Combo2.ListIndex = 2 Then
        Text1.Text = 2
        NamaShiftReset = 2
        DoEvents
        Text3.Text = Format(ks, "00")
        DoEvents
        namaMesinReset = Text3.Text
        DoEvents
        Text4.Text = "\\KASSA" & Format(Val(namaMesinReset), "0#") & "\pos"
        Text4.Locked = True
        If Text4.Text = "" Then Exit Sub
        If Dir(Text4.Text & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb") = "" Then
          frmPesan.lbPesan = "MESIN TIDAK ONLINE!!"
          frmPesan.Show 1
          Text4.Locked = False
          Exit Sub
        End If
        DoEvents
        DIRPOS = Text4.Text
        DayReset = Format(tgl, "dd") 'Val(Mid(Text2.Text, 1, 2))
        MonthReset = Format(tgl, "MM") ' Val(Mid(Text2.Text, 3, 2))
        YearReset = Format(tgl, "yyyy") 'Val(Mid(Text2.Text, 5, 4))
        DoEvents
        AmbilPersenPerMesin
        If Text3.Text = "09" Or Text3.Text = "10" Or Text3.Text = "20" Or Text3.Text = "99" Then
            CekTransaksiBermasalahBusana
        Else
            CekTransaksiBermasalah
        End If
        If isDataOK Then
            Screen.MousePointer = vbHourglass
            If Text3.Text = "09" Or Text3.Text = "10" Or Text3.Text = "20" Or Text3.Text = "99" Then
                ResetBusana
            Else
                ResetRetail
            End If
            ResendServerOnline
          
          If (UCase(Trim(getRegistry("AutoDelete", "Data"))) = "Y") And (KodeKasir = KodeUserDua) Then
            HAPUSTRANSAKSI
          End If
          Screen.MousePointer = vbDefault
        Else
    '      frmPesan.lbPesan = "Ada Transaksi Pending!!!!"
            MsgBox "Ada Transaksi Pending pada tanggal: " & Format(tgl, "dd-MM-yyyy") & "!!!!"
        End If
    End If
    Next

Next
    MsgBox "RESET SELESAI!", vbOKOnly + vbInformation

End If

Exit Sub
pesan:
MsgBox "Ada kesalahan :" & Err.Description, vbInformation + vbOKOnly
End Sub

Private Sub Form_Load()
TDBDate1 = Date
TDBDate2 = Date

  isHasilKonversi = False
   If Format(Time, "HHnnss") > "150000" Then
    Text1.Text = 2
  Else
    Text1.Text = 1
  End If
  Text2.Text = Format(Date, "DDMMYYYY")
  Text3.Text = NamaMesin
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim hasil As String
hasil = Trim(SendByCode(KeyCode))
KeyCode = 0
Select Case hasil
Case "1", "2"
  Text1.Text = Text1.Text & hasil
  Text1.SelStart = Len(Text1.Text)
 Case "SPC"
  Text1.Text = Text1.Text & " "
  Text1.SelStart = Len(Text1.Text)
Case "BKS"
  If Len(Text1.Text) > 0 Then Text1.Text = Left(Text1.Text, Len(Text1.Text) - 1)
   Text1.SelStart = Len(Text1.Text)
Case "CLR"
    Text1.Text = ""
Case "ENT"
    If Text1.Text = "" Then Exit Sub
    If IsNumeric(Text1.Text) Then
        NamaShiftReset = Text1.Text
    Else
        NamaShiftReset = 1
    End If
    Text2.SetFocus
'    CekTransaksiBermasalah
'    If isDataOK Then
'      ResetNew
'      frmPesan.lbPesan = "Selesai !!!!"
'
'    Else
'      frmPesan.lbPesan = "Ada Transaksi Pending!!!!"
'      frmPesan.Show 1
'      Unload Me
'    End If
Case "ESC"
    Unload Me
End Select
End Sub

Function CekKassaShift1() As String
Dim dbc As Database
Dim rsc As Recordset
Dim NamaFileReset1 As String
Dim TotalReset1 As Double
Dim totalhasil1 As Double
NamaFileReset1 = DIRPOS & "\Reset\K" & Trim(NamaMesin) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & "01.mdb"
If Dir(NamaFileReset1) = "" Then
CekKassaShift1 = "-"
Exit Function
End If
Set dbc = OpenDatabase(NamaFileReset1)
Set rsc = dbc.OpenRecordset("SELECT Sum(SubTotal) as Total From MSales")
If rsc.EOF And rsc.BOF Then
Else
  TotalReset1 = IIf(IsNull(rsc!Total), 0, rsc!Total)
End If
Set dbc = OpenDatabase(DIRPOS & "\Database\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
Set rsc = dbc.OpenRecordset("SELECT Sum(SubTotal) as Total From MSales WHERE " & _
          "day(tanggal)=" & DayReset & " AND month(tanggal)=" & MonthReset & " and year(tanggal)=" & YearReset & " AND Shift=1")
If rsc.EOF And rsc.BOF Then
Else
  totalhasil1 = IIf(IsNull(rsc!Total), 0, rsc!Total)
End If
If KodeKasir = KodeUserDua Then
    CekKassaShift1 = ""
Else
    If totalhasil1 <> TotalReset1 Then
      CekKassaShift1 = "!=!=!=!=!=RESET-SELESAI!=!=!=!=!="
    Else
      CekKassaShift1 = "======RESET-SELESAI======="
    End If
End If
End Function
Sub CekTransaksiBermasalah()
Dim dbc As Database
Dim rsc As Recordset

Set dbc = OpenDatabase(DIRPOS & "\Database\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
Set rsc = dbc.OpenRecordset("SELECT MSales.NoID, Sum([Qty]*[Harga]) AS QSubTotal, MSales.SubTotal, MSales.DiscNota,MSales.Pembulatan, MSales.HargaTotal, MSales.UangMuka, MSales.ISPending " & _
    "FROM MSales INNER JOIN MSalesD ON MSales.NoID = MSalesD.IDSales " & _
    "WHERE Day(Tanggal)=" & DayReset & " AND Month(Tanggal)=" & MonthReset & " AND Year(Tanggal)=" & YearReset & " and SubTotal<>0 AND Shift= " & NamaShiftReset & " " & _
    "GROUP BY MSales.NoID, MSales.SubTotal, MSales.DiscNota, MSales.HargaTotal, MSales.UangMuka,MSales.Pembulatan, MSales.ISPending " & _
    "HAVING (((Sum([Qty]*[Harga]))<>[SubTotal])) OR (((Sum([Qty]*[Harga]))>[HargaTotal]+[DiscNota]+[Pembulatan])) OR (((Sum([Qty]*[Harga]))>[UangMuka]+[DiscNota]+[Pembulatan]))")
If rsc.EOF And rsc.BOF Then
Else
  rsc.MoveFirst
  Do While Not rsc.EOF
    'dbc.Execute "Update MSales set SubTotal=" & CCur(IIf(IsNull(rsc!QSubTotal), 0, rsc!QSubTotal)) & ", HargaTotal=" & CCur(IIf(IsNull(rsc!QSubTotal), 0, rsc!QSubTotal)) - CCur(IIf(IsNull(rsc!DiscNota), 0, rsc!DiscNota)) & ", ISpending=" & IIf((rsc!QSubTotal = (rsc!UangMuka + rsc!DiscNota)), False, True) & " Where NoId=" & rsc!NoId
    dbc.Execute "Update MSales set SubTotal=" & CCur(IIf(IsNull(rsc!QSubTotal), 0, rsc!QSubTotal)) & ", ISpending=" & IIf((rsc!QSubTotal <= (rsc!UangMuka + rsc!DiscNota + rsc!Pembulatan)), False, True) & " Where NoId=" & rsc!NoID
  rsc.MoveNext
  Loop
End If
Set rsc = dbc.OpenRecordset("Select * From MSales Where IsPending=TRUE AND Day(Tanggal)=" & DayReset & " AND Month(Tanggal)=" & MonthReset & " AND Year(Tanggal)=" & YearReset & " and SubTotal<>0 AND Shift= " & NamaShiftReset)
If rsc.EOF And rsc.BOF Then
  isDataOK = True
Else
  isDataOK = False
End If
rsc.Close
Set rsc = Nothing
dbc.Close
Set dbc = Nothing
End Sub

Sub CekTransaksiBermasalahBusana()
Dim dbc As Database
Dim rsc As Recordset

Set dbc = OpenDatabase(DIRPOS & "\Database\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
'dbc.Execute "Update MSales set IsPending=False where IsPending=True"
'Set rsc = dbc.OpenRecordset("SELECT MSales.NoID, Sum([Qty]*[Harga]) AS QSubTotal, MSales.SubTotal, MSales.DiscNota,MSales.Pembulatan,MSales.DiscIntern, MSales.HargaTotal, MSales.UangMuka, MSales.ISPending " & _
'    "FROM MSales INNER JOIN MSalesD ON MSales.NoID = MSalesD.IDSales " & _
'    "WHERE Day(Tanggal)=" & DayReset & " AND Month(Tanggal)=" & MonthReset & " AND Year(Tanggal)=" & YearReset & " and SubTotal<>0 AND Shift= " & NamaShiftReset & " " & _
'    "GROUP BY MSales.NoID, MSales.SubTotal, MSales.DiscNota,MSales.DiscIntern,  MSales.HargaTotal, MSales.UangMuka,MSales.Pembulatan,MSales.ISPending " & _
'    "HAVING (((Sum([Qty]*[Harga]))<>[SubTotal])) OR (((Sum([Qty]*[Harga]))>[HargaTotal]+[DiscNota]+[DiscIntern]+[Pembulatan])) OR (((Sum([Qty]*[Harga]))>[UangMuka]+[DiscNota]+[DiscIntern]+[Pembulatan]))")
Set rsc = dbc.OpenRecordset("SELECT MSales.NoID, Sum([Qty]*[Harga]) AS QSubTotal, MSales.SubTotal, MSales.DiscNota,MSales.Pembulatan,MSales.DiscIntern,MSales.JumDiscInternRp, MSales.HargaTotal, MSales.UangMuka, MSales.ISPending,MSales.TotalDiscount " & _
    "FROM MSales INNER JOIN MSalesD ON MSales.NoID = MSalesD.IDSales " & _
    "WHERE Day(Tanggal)=" & DayReset & " AND Month(Tanggal)=" & MonthReset & " AND Year(Tanggal)=" & YearReset & " and SubTotal<>0 AND Shift= " & NamaShiftReset & " " & _
    "GROUP BY MSales.NoID, MSales.SubTotal, MSales.DiscNota,MSales.DiscIntern,MSales.JumDiscInternRp,  MSales.HargaTotal, MSales.UangMuka,MSales.Pembulatan,MSales.ISPending,Msales.TotalDiscount " & _
    "HAVING (abs(Sum([Qty]*[Harga])-([SubTotal]))>1 OR ([SubTotal]>[HargaTotal]+[DiscNota]+[DiscIntern]+[Pembulatan]) )")
If rsc.EOF And rsc.BOF Then
Else
  rsc.MoveFirst
  Do While Not rsc.EOF
    'dbc.Execute "Update MSales set SubTotal=" & CCur(IIf(IsNull(rsc!QSubTotal), 0, rsc!QSubTotal)) & ", HargaTotal=" & CCur(IIf(IsNull(rsc!QSubTotal), 0, rsc!QSubTotal)) - CCur(IIf(IsNull(rsc!DiscNota), 0, rsc!DiscNota)) & ", ISpending=" & IIf((rsc!QSubTotal = (rsc!UangMuka + rsc!DiscNota)), False, True) & " Where NoId=" & rsc!NoId
    dbc.Execute "Update MSales set SubTotal=" & CCur(IIf(IsNull(rsc!QSubTotal), 0, rsc!QSubTotal)) & ", ISpending=" & IIf((rsc!QSubTotal <= (rsc!UangMuka + rsc!DiscNota + rsc!DiscIntern + rsc!Pembulatan)), False, True) & " Where NoId=" & rsc!NoID
  rsc.MoveNext
  Loop
End If
Set rsc = dbc.OpenRecordset("Select * From MSales Where IsPending=TRUE AND Day(Tanggal)=" & DayReset & " AND Month(Tanggal)=" & MonthReset & " AND Year(Tanggal)=" & YearReset & " and SubTotal<>0 AND Shift= " & NamaShiftReset)
If rsc.EOF And rsc.BOF Then
  isDataOK = True
Else
  isDataOK = False
End If
rsc.Close
Set rsc = Nothing
dbc.Close
Set dbc = Nothing
End Sub

Sub Tampil(ByRef jawaban As Boolean, Key As Integer)
  KeyOk = Key
  Me.Show 1
  jawaban = jawab
End Sub
Sub ResetNewDiskonDihitung()
On Error GoTo pesan

    Dim dbs As Database
    Dim rs As Recordset
    Dim DbsReset As Database
    Dim RsReset As Recordset
    Dim psn As String
  Dim i As Integer
  Dim NoID As Integer
  Dim idkasirreset As Integer
  Dim NamaFileReset As String
  Dim NamaFileBackup As String
  Dim TunaiPajak As Long
  Dim cTunai As Long
  Dim cJumlahNota As Long
  Dim cbank As Double
  Dim cSubtotal As Double
  Dim cUangMuka As Double
  Dim cTotal As Double
  Dim cDiskonNota As Double
  Dim cDiskonBrg As Double
  Dim cVoucher As Double
  Dim dbsClear As Database
  Dim strqry As String
  Dim MaxTotal As Double
  Dim curTotal As Double
  Dim curQty As Long
  Dim CurhargaSatuan As Double
  Dim cekFile As String
  Dim JmlSalah As Integer
    lblNama(9).Caption = "Biaya Credit Card:"
    Set dbs = OpenDatabase(DIRPOS & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
    dbs.Execute "Update Msales Set IDbank =2 where isnull(idbank)"
    strqry = "SELECT MSales.IDUser,Count(MSales.Kode) AS JumlahNota, Sum(MSales.SubTotal) AS JumlahSubTotal, Sum(MSales.HargaTotal) AS JumlahTotal, sum(MSales.DiscNota) as JumDiskon, sum(MSales.TotalDiscount) as DiskonBrg, Sum(MSales.UangMuka-MSales.Bank-MSales.Voucher) AS JumUangMuka, Sum(MSales.Bank) AS JumBank , Sum(MSales.Voucher) AS jumVoucher " & _
            "From MSales " & _
            "WHERE Day(Tanggal)=" & DayReset & " AND Month(Tanggal)=" & MonthReset & " AND Year(Tanggal)=" & YearReset & " and SubTotal<>0 AND Shift= " & NamaShiftReset & " " & _
            "GROUP BY MSales.IDUser"
    Set rs = dbs.OpenRecordset(strqry)
    If rs.EOF And rs.BOF Then
      Exit Sub
    Else
    rs.MoveFirst
       PrinBigChar Chr(13) & Chr(10) & Chr(13) & "   KASSA : " & Str(namaMesinReset)
       psn = Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "----------------RESET----------------" & Chr(13) & Chr(10) & Chr(13) & Chr(10)
       Prin psn
       psn = ""
    Do While Not rs.EOF
      idkasirreset = rs!IDUser
        cariKodeNama KodeKasirReset, namakasirReset, rs!IDUser
        If KodeKasir = KodeUserDua Then
            TunaiPajak = ((PersenLap / 100) * rs!JumUangMuka) - (((PersenLap / 100) * rs!JumUangMuka) Mod 100)
            cTunai = rs!JumUangMuka
            cJumlahNota = rs!JumlahNota
            cbank = rs!Jumbank
            'cVoucher = rs!JumVoucher
            cVoucher = 0
            cDiskonNota = rs!JumDIskon
            cSubtotal = cTunai + cbank + cDiskonNota + cVoucher + rs!DIskonBrg
            lbl(0) = Format(DayReset, "0#") & "-" & Format(MonthReset, "0#") & "-" & Format(YearReset, "####") & ", " & IIf(NamaShiftReset = 1, "PAGI", "SORE") 'Format(Now, "dd-MM-yyyy, hh:mm:ss")
            lbl(1) = namaMesinReset & " - " & namakasirReset
            lbl(2) = Format(rs!JumlahNota, "###,##0")
            lbl(3) = Format(rs!Jumbank + TunaiPajak + cVoucher + rs!JumDIskon + rs!DIskonBrg, "###,###,##0")
            lbl(4) = Format(rs!JumDIskon + rs!DIskonBrg, "###,###,##0")
            lbl(5) = Format(rs!Jumbank + TunaiPajak, "###,###,##0")
            lbl(6) = Format(TunaiPajak, "###,###,##0")
            lbl(7) = Format(rs!Jumbank, "###,###,##0")
            lbl(8) = Format(cVoucher, "###,###,##0")
            'lbl(9) = (rs!Jumbank + TunaiPajak)
        Else
            'lbl(0) = Format(Now, "dd-MM-yyyy, hh:mm:ss")
            lbl(0) = Format(DayReset, "0#") & "-" & Format(MonthReset, "0#") & "-" & Format(YearReset, "####") & ", " & IIf(NamaShiftReset = 1, "PAGI", "SORE") 'Format(Now, "dd-MM-yyyy, hh:mm:ss")
            lbl(1) = namaMesinReset & " - " & namakasirReset
            lbl(2) = Format(rs!JumlahNota, "###,##0")
            lbl(3) = Format(rs!JumlahSubTotal + rs!DIskonBrg, "###,###,##0")
            lbl(4) = Format(rs!JumDIskon + rs!DIskonBrg, "###,###,##0")
            lbl(5) = Format(rs!JumlahTotal, "###,###,##0")
            lbl(6) = Format(rs!JumUangMuka, "###,###,##0")
            lbl(7) = Format(rs!Jumbank, "###,###,##0")
            lbl(8) = Format(rs!JumVoucher, "###,###,##0")
            lbl(9) = Format(rs!JumlahTotal - rs!JumlahSubTotal + rs!JumDIskon, "###,###,##0")
            TunaiPajak = ((PersenLap / 100) * rs!JumUangMuka) - (((PersenLap / 100) * rs!JumUangMuka) Mod 100)
            cVoucher = rs!JumVoucher
            cTunai = rs!JumUangMuka
            cJumlahNota = rs!JumlahNota
            cbank = rs!Jumbank
            cDiskonNota = rs!JumDIskon
            cDiskonBrg = rs!DIskonBrg
            cSubtotal = cTunai + cbank + cDiskonNota + cVoucher + cDiskonBrg
        End If
            If KodeKasir = KodeUserDua Then
            For i = 0 To 7
               psn = psn & lblNama(i).Caption & Space(18 - Len(Trim(lblNama(i).Caption))) & ": " & lbl(i).Caption & Chr(13) & Chr(10) & Chr(13) & Chr(10)
              Next
            Else
              For i = 0 To 9
               psn = psn & lblNama(i).Caption & Space(18 - Len(Trim(lblNama(i).Caption))) & ": " & lbl(i).Caption & Chr(13) & Chr(10) & Chr(13) & Chr(10)
              Next
              End If
              Prin psn
              psn = ""
        rs.MoveNext
    Loop
End If
    psn = psn & "----------------------------------------" & Chr(13) & Chr(10)
    psn = psn & CekKassaShift1 & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10)
    
    Prin psn
    PrinBigChar Space(15 - Min(15, Len(NamaToko))) & UCase(NamaToko)
    papercut
    Prin Judulstruk & Chr(13) & Chr(10) & "----------------------------------------"
    'MODIFIKASI 20-03-2007
    'Dijadikan 2 , server dan lokal
    '
    If KodeKasir = KodeUserDua Then
        NamaFileReset = Replace(getRegistry("Reset2", "Data") & "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb", "*", Text3.Text)
        NamaFileBackup = Replace(getRegistry("Backup2", "Data") & "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb", "*", Text3.Text)
    Else
        NamaFileReset = Replace(getRegistry("Reset1", "Data") & "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb", "*", Text3.Text)
    'Text2.Text = getRegistry("Reset2", "Data")
        NamaFileBackup = Replace(getRegistry("Backup", "Data") & "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb", "*", Text3.Text)
    End If
    
'    NamaFileReset = DIRPOS & "\Reset\K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb"
'    NamaFileBackup = DIRPOS & "\Backup\K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb"
    'RESET
    cekFile = "Reset"
    If Dir(NamaFileReset) <> "" Then
     Set dbsClear = OpenDatabase(NamaFileReset)
     dbsClear.Execute "Delete * FROM MSALES"
     dbsClear.Execute "Delete * FROM MSALESD"
     dbsClear.Close
    Else
     If KodeKasir = KodeUserDua Then
        Create_DatabaseResetDua NamaFileReset
     Else
        Create_DatabaseReset NamaFileReset
      End If
    End If
    cekFile = "Backup"
    'If KodeKasir <> KodeUserDua Then
        If Dir(NamaFileBackup) <> "" Then
         Set dbsClear = OpenDatabase(NamaFileBackup)
         dbsClear.Execute "Delete * FROM MSALES"
         dbsClear.Execute "Delete * FROM MSALESD"
         dbsClear.Close
        Else
          Create_DatabaseBackup NamaFileBackup
        End If
               
        'BackUp
        dbs.Execute "INSERT INTO MSales ( Shift, IDUser,IDBank, Tanggal, NoID, Kode, TotalPajak, TotalDiscount, SubTotal, DiscNota, HargaTotal, IDPayment, UangMuka, IsSend, Bank, ISPending,Voucher ) IN '" & NamaFileBackup & "' " & _
                  "SELECT MSales.Shift, MSales.IDUser,MSales.IDBank, MSales.Tanggal, MSales.NoID, MSales.Kode, MSales.TotalPajak, MSales.TotalDiscount, MSales.SubTotal, MSales.DiscNota, MSales.HargaTotal, MSales.IDPayment, MSales.UangMuka, MSales.IsSend, MSales.Bank, MSales.ISPending , MSales.Voucher " & _
                  "From MSales " & _
                  "WHERE (((MSales.Shift)=" & NamaShiftReset & ") AND  (Day(MSales.Tanggal)=" & DayReset & " and Month(MSales.Tanggal)=" & MonthReset & " and Year(MSales.Tanggal)=" & YearReset & "))"
        
        dbs.Execute "INSERT INTO MSalesD IN '" & NamaFileBackup & "' " & _
                  "SELECT MSalesD.*,MsalesD.Qty*MSalesD.Harga as Jumlah " & _
                  "FROM MSales INNER JOIN MSalesD ON MSales.NoID = MSalesD.IDSales " & _
                  "WHERE (((MSales.Shift)=" & NamaShiftReset & ") AND (Day(MSales.Tanggal)=" & DayReset & " and Month(MSales.Tanggal)=" & MonthReset & " and Year(MSales.Tanggal)=" & YearReset & "))"
         dbs.Close
    'End If
      
  'RESET
    Set DbsReset = OpenDatabase(NamaFileReset)
    Set RsReset = DbsReset.OpenRecordset("MSales")
    Set dbs = OpenDatabase(DIRPOS & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
    Set rs = dbs.OpenRecordset("SELECT   sum(SubTotal) as SubTotala,sum(DiscNota) as DiscNotaa,sum(TotalDiscount) as DiscBrga, " & _
            "sum(HargaTotal) as HargaTotala, sum(UangMuka)-sum(Bank) as Tunaia, sum(Voucher) as Vouchera," & _
            "sum(Bank) as Banka,IDBank,IDUser,count(NoID) as JumlahNotaa , Max(SubTotal) as NotaMax,Min(SubTotal) as NotaMin " & _
            "From MSales " & _
              "WHERE (((MSales.Shift)=" & NamaShiftReset & ") AND  (Day(MSales.Tanggal)=" & DayReset & " and Month(MSales.Tanggal)=" & MonthReset & " and Year(MSales.Tanggal)=" & YearReset & ") AND (SubTotal<>0)) " & _
              "Group By day(Tanggal),month(Tanggal),year(Tanggal),IDBank,IDUser HAVING sum(HargaTotal)<>0")
     If rs.EOF And rs.BOF Then
    
     Else
        rs.MoveFirst
        NoID = 1
        Do While Not rs.EOF
            RsReset.AddNew
            RsReset!kode = "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & Format(NoID, "0##0")
            RsReset!NoID = NoID
            RsReset!TANGGAL = Mid(Text2.Text, 3, 2) & "/" & Mid(Text2.Text, 1, 2) & "/" & Mid(Text2.Text, 5, 4)
            RsReset!SubTotal = rs!SubTotala
            RsReset!DiscNota = rs!DiscNotaa
            RsReset!Hargatotal = rs!HargaTotala
            RsReset!Tunai = rs!Tunaia
            RsReset!Voucher = rs!Vouchera
            RsReset!Bank = rs!Banka
            RsReset!IDBank = rs!IDBank
            RsReset!IDUser = rs!IDUser
            RsReset!NamaUser = namakasirReset
            RsReset!Shift = NamaShiftReset
            RsReset!IdPengawas = IDUser
            RsReset!NamaPengawas = NamaKasir
            If KodeKasir = KodeUserDua Then
            MaxTotal = PersenLap * rs!Tunaia / 100
            Else
                RsReset!JumlahNota = rs!JumlahNotaa
                RsReset!PajakPersen = PersenLap
                RsReset!NotaMin = rs!NotaMin
                RsReset!NotaMax = rs!NotaMax
                RsReset!TunaiPajak = PersenLap * rs!Tunaia / 100
                MaxTotal = rs!Tunaia '100 persen
            End If
            
            
            RsReset.Update
            RsReset.Bookmark = RsReset.LastModified
            
            If rs!Banka = 0 And KodeKasir = KodeUserDua Then 'Tunai dan kodeuserdua
                Dim RSp As Recordset
                Set RSp = dbs.OpenRecordset("SELECT " & NoID & " As IDSales, MSalesD.IDInventor, Sum(MSalesD.Qty) AS Qty, Sum((MSalesD.Harga*MSalesD.Qty)) AS Jumlah, Sum((MSalesD.HargaPokok*MSalesD.Qty)) AS HPP, MSalesD.KodeInv, MSalesD.NamaInv, MSalesD.Satuan, MSalesD.Barcode, MSalesD.IDSatuan, MSalesD.Konversi " & _
                "FROM MSalesD INNER JOIN MSales ON MSalesD.IDSales = MSales.NoID " & _
                "Where day(MSales.Tanggal) = " & DayReset & " and Month(MSales.Tanggal) = " & MonthReset & " and Year(MSales.Tanggal) = " & YearReset & "  And MSales.ShifT = " & NamaShiftReset & " " & _
                "GROUP BY MSales.IDUser,Msales.IDBank,MSalesD.IDInventor, MSalesD.KodeInv, MSalesD.NamaInv, MSalesD.Satuan, MSalesD.Barcode, MSalesD.IDSatuan, MSalesD.Konversi " & _
                "HAVING (((Sum(MSalesD.Qty))<>0) AND MSales.IDUser=" & rs!IDUser & " AND MSales.IDBank=" & NullToNol(rs!IDBank) & ") " & _
                "ORDER BY Max(MSales.Tanggal) DESC ")
                RSp.MoveFirst
                curTotal = 0
                Do While Not RSp.EOF
                    If MaxTotal <= curTotal Then Exit Do
                    curQty = ((RSp!QTY * PersenLap - RSp!QTY * PersenLap Mod 100) \ 100) + IIf(RSp!QTY * PersenLap Mod 100 > 1, 1, 0)
                    CurhargaSatuan = CLng(IIf(RSp!QTY <> 0, RSp!Jumlah / RSp!QTY, 0) * curQty)
                    If CurhargaSatuan Mod 50 > 0 Then CurhargaSatuan = CurhargaSatuan + (50 - (CurhargaSatuan Mod 50))
                    curTotal = curTotal + CurhargaSatuan
                      dbs.Execute "INSERT INTO MSalesD ( IDSales,IDInventor, Qty,HargaSat, Harga, HargaPokok," & _
                                  "KodeInv, NamaInv, Satuan, Barcode, IDSatuan, Konversi ) " & _
                                  "IN '" & NamaFileReset & "' " & _
                                  "VALUES(" & _
                                   NoID & "," & RSp!IdInventor & "," & curQty & "," & Replace(Format(IIf(curQty = 0, CurhargaSatuan, Abs(CurhargaSatuan / curQty)), "##0.00"), ",", ".") & "," & Replace(Format(CurhargaSatuan, "##0.00"), ",", ".") & ",0,'" & _
                                  Replace(RSp!KodeInv, "'", "''") & "','" & Replace(RSp!NamaInv, "'", "''") & "','" & Replace(RSp!Satuan, "'", "''") & "','" & RSp!Barcode & "'," & RSp!idSatuan & "," & RSp!Konversi & ")"

                    RSp.MoveNext
                Loop
                RsReset.Edit
                
                RsReset!SubTotal = curTotal
                'RsReset!DiscNota = curTotal
                RsReset!Hargatotal = curTotal
                RsReset!Tunai = curTotal - rs!DiscNotaa
                RsReset.Update
            Else
              dbs.Execute "INSERT INTO MSalesD ( IDSales,IDInventor, Qty, Harga, HargaPokok, KodeInv, NamaInv, Satuan, Barcode, IDSatuan, Konversi ) IN '" & NamaFileReset & "'" & _
                "SELECT " & NoID & " As IDSales, MSalesD.IDInventor, Sum(MSalesD.Qty) AS Qty, Sum((MSalesD.Harga*MSalesD.Qty)) AS Jumlah, Sum((MSalesD.HargaPokok*MSalesD.Qty)) AS HPP, MSalesD.KodeInv, MSalesD.NamaInv, MSalesD.Satuan, MSalesD.Barcode, MSalesD.IDSatuan, MSalesD.Konversi " & _
                "FROM MSalesD INNER JOIN MSales ON MSalesD.IDSales = MSales.NoID " & _
                "Where day(MSales.Tanggal) = " & DayReset & " and Month(MSales.Tanggal) = " & MonthReset & " and Year(MSales.Tanggal) = " & YearReset & "  And MSales.ShifT = " & NamaShiftReset & " " & _
                "GROUP BY MSales.IDUser,Msales.IDBank,MSalesD.IDInventor, MSalesD.KodeInv, MSalesD.NamaInv, MSalesD.Satuan, MSalesD.Barcode, MSalesD.IDSatuan, MSalesD.Konversi " & _
                "HAVING (((Sum(MSalesD.Qty))<>0) AND MSales.IDUser=" & rs!IDUser & " AND MSales.IDBank=" & NullToNol(rs!IDBank) & ")"
            DbsReset.Execute "UPDATE MSalesD SET MSalesD.HargaSat = [msalesd].[harga]/Abs([msalesd].[qty]) WHERE (((MSalesD.Qty)<>0)) AND IDSales=" & NoID
            End If
            DbsReset.Execute "Update MSALES  set Tanggal=#" & Mid(Text2.Text, 3, 2) & "/" & Mid(Text2.Text, 1, 2) & "/" & Mid(Text2.Text, 5, 4) & "#"
    '            If rs!Banka = 0 Then
    '                If KodeKasir = KodeUserDua Then
    '                  DbsReset.Execute "UPDATE MSalesD SET MSalesD.Qty=(MSalesD.Qty*" & PersenLap & ")\100,MSalesD.[harga]=((MSalesD.Qty*" & PersenLap & ")\100)*MSalesD.HargaSat WHERE  IDSales=" & NoId
    '                 End If
    '            End If
                NoID = NoID + 1
        rs.MoveNext
        Loop
     End If
 
    
    Set rs = Nothing
    dbs.Close
    Set RsReset = Nothing
    DbsReset.Close
    Set dbs = Nothing
    Set DbsReset = Nothing
    Exit Sub
pesan:
'   MsgBox "Ada kesalahan : " & Err.Description & vbCrLf & "Silahkan Tekan OK", vbCritical + vbOKOnly, "Pesan!"
   If Err.Number = 52 Then
    If cekFile = "Backup" Then
        NamaFileBackup = NMBACKUPGAGAL & "\" & "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb"
        Resume
    ElseIf cekFile = "Reset" Then
        NamaFileBackup = NMRESETGAGAL & "\" & "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb"
        Resume
    End If
   Else
    Resume Next
    End If
'    Unload Me
End Sub
Sub ResetRetail()
On Error GoTo pesan

    Dim dbs As Database
    Dim rs As Recordset
    Dim DbsReset As Database
    Dim RsReset As Recordset
    Dim psn As String
  Dim i As Integer
  Dim NoID As Integer
  Dim idkasirreset As Integer
  Dim NamaFileReset As String
  Dim NamaFileBackup As String
  Dim TunaiPajak As Long
  Dim cTunai As Long
  Dim cJumlahNota As Long
  Dim cbank As Double
  Dim cSubtotal As Double
  Dim cUangMuka As Double
  Dim cTotal As Double
  Dim cDiskonNota As Double
  Dim cDiskonBrg As Double
  Dim cVoucher As Double
  Dim dbsClear As Database
  Dim strqry As String
  Dim MaxTotal As Double
  Dim curTotal As Double
  Dim curQty As Long
  Dim CurhargaSatuan As Double
  Dim cekFile As String
  Dim JmlSalah As Integer
  Dim Bulat As Long
  
    lblNama(9).Caption = "Biaya Credit Card"
    Set dbs = OpenDatabase(DIRPOS & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
    dbs.Execute "Update Msales Set IDbank =2 where isnull(idbank)"
    strqry = "SELECT MSales.IDUser,Count(MSales.Kode) AS JumlahNota, Sum(MSales.SubTotal) AS JumlahSubTotal, Sum(MSales.HargaTotal) AS JumlahTotal, sum(MSales.DiscNota) as JumDiskon, sum(MSales.TotalDiscount) as DiskonBrg, Sum(MSales.UangMuka-MSales.Bank-MSales.Voucher) AS JumUangMuka, Sum(MSales.Bank) AS JumBank , Sum(MSales.Voucher) AS jumVoucher, " & _
            "Sum(MSales.Pembulatan) as Bulat " & _
            "From MSales " & _
            "WHERE Day(Tanggal)=" & DayReset & " AND Month(Tanggal)=" & MonthReset & " AND Year(Tanggal)=" & YearReset & " and SubTotal<>0 AND Shift= " & NamaShiftReset & " " & _
            "GROUP BY MSales.IDUser"
    Set rs = dbs.OpenRecordset(strqry)
    If rs.EOF And rs.BOF Then
      Exit Sub
    Else
    rs.MoveFirst
       PrinBigChar Chr(13) & Chr(10) & Chr(13) & "   KASSA : " & Str(namaMesinReset)
       psn = Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "----------------RESET----------------" & Chr(13) & Chr(10) & Chr(13) & Chr(10)
       Prin psn
       psn = ""
    Do While Not rs.EOF
      idkasirreset = rs!IDUser
        cariKodeNama KodeKasirReset, namakasirReset, rs!IDUser
        If KodeKasir = KodeUserDua Then
            TunaiPajak = ((PersenLap / 100) * rs!JumUangMuka) - (((PersenLap / 100) * rs!JumUangMuka) Mod 100)
            cTunai = rs!JumUangMuka
            cJumlahNota = rs!JumlahNota
            cbank = rs!Jumbank
            'cVoucher = rs!JumVoucher
            cVoucher = 0
            cDiskonNota = rs!JumDIskon
            cSubtotal = cTunai + cbank + cDiskonNota + cVoucher + rs!DIskonBrg
            lbl(0) = Format(DayReset, "0#") & "-" & Format(MonthReset, "0#") & "-" & Format(YearReset, "####") & ", " & IIf(NamaShiftReset = 1, "PAGI", "SORE") 'Format(Now, "dd-MM-yyyy, hh:mm:ss")
            lbl(1) = namaMesinReset & " - " & namakasirReset
            lbl(2) = Format(rs!JumlahNota, "###,##0")
            lbl(3) = Format(rs!Jumbank + TunaiPajak + cVoucher + rs!JumDIskon + rs!DIskonBrg, "###,###,##0")
            lbl(4) = Format(rs!JumDIskon + rs!DIskonBrg, "###,###,##0")
            lbl(5) = Format(rs!Jumbank + TunaiPajak, "###,###,##0")
            lbl(6) = Format(TunaiPajak, "###,###,##0")
            lbl(7) = Format(rs!Jumbank, "###,###,##0")
            lbl(8) = Format(cVoucher, "###,###,##0")
            'lbl(9) = (rs!Jumbank + TunaiPajak)
            Bulat = Format(rs!Bulat, "###,###,##0")
        Else
            'lbl(0) = Format(Now, "dd-MM-yyyy, hh:mm:ss")
            lbl(0) = Format(DayReset, "0#") & "-" & Format(MonthReset, "0#") & "-" & Format(YearReset, "####") & ", " & IIf(NamaShiftReset = 1, "PAGI", "SORE") 'Format(Now, "dd-MM-yyyy, hh:mm:ss")
            lbl(1) = namaMesinReset & " - " & namakasirReset
            lbl(2) = Format(rs!JumlahNota, "###,##0")
            lbl(3) = Format(rs!JumlahSubTotal + rs!DIskonBrg + rs!JumDIskon, "###,###,##0") '+ rs!DIskonBrg + rs!JumDIskon
            lbl(4) = Format(rs!JumDIskon + rs!DIskonBrg, "###,###,##0")
            lbl(5) = Format(rs!JumlahTotal + rs!JumDIskon + rs!DIskonBrg, "###,###,##0")
            lbl(6) = Format(rs!JumUangMuka, "###,###,##0")
            lbl(7) = Format(rs!Jumbank, "###,###,##0")
            lbl(8) = Format(rs!JumVoucher, "###,###,##0")
            lbl(9) = Format(rs!JumlahTotal - rs!JumlahSubTotal + rs!JumDIskon + rs!Bulat, "###,###,##0")
            TunaiPajak = ((PersenLap / 100) * rs!JumUangMuka) - (((PersenLap / 100) * rs!JumUangMuka) Mod 100)
            cVoucher = rs!JumVoucher
            cTunai = rs!JumUangMuka
            cJumlahNota = rs!JumlahNota
            cbank = rs!Jumbank
            cDiskonNota = rs!JumDIskon
            cDiskonBrg = rs!DIskonBrg
            Bulat = Format(rs!Bulat, "###,###,##0")
            cSubtotal = cTunai + cbank + cDiskonNota + cVoucher + cDiskonBrg + rs!Bulat
        End If
            If KodeKasir = KodeUserDua Then
             For i = 0 To 7
                  psn = psn & lblNama(i).Caption & Space(18 - Len(Trim(lblNama(i).Caption))) & ": " & lbl(i).Caption & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                    'If i = 4 Then
                    '    psn = psn & "Pembulatan" & Space(18 - Len("Pembulatan")) & ": " & Format(Bulat, "###,##0") & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                    'End If
             Next
            Else
             For i = 0 To 9
               psn = psn & lblNama(i).Caption & Space(18 - Len(Trim(lblNama(i).Caption))) & ": " & lbl(i).Caption & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                    If i = 4 Then
                        psn = psn & "Pembulatan" & Space(18 - Len("Pembulatan")) & ": " & Format(Bulat, "###,##0") & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                    End If
             Next
            End If
             Prin psn
             psn = ""
             rs.MoveNext
    Loop
End If
    psn = psn & "----------------------------------------" & Chr(13) & Chr(10)
    psn = psn & CekKassaShift1 & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10)
    
    Prin psn
'    PrinBigChar Space(15 - Min(15, Len(NamaToko))) & UCase(NamaToko)
'    papercut
'    Prin Judulstruk & Chr(13) & Chr(10) & "----------------------------------------"
    'MODIFIKASI 20-03-2007
    'Dijadikan 2 , server dan lokal
    '
    If KodeKasir = KodeUserDua Then
        NamaFileReset = Replace(getRegistry("Reset2", "Data") & "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb", "*", Text3.Text)
        NamaFileBackup = Replace(getRegistry("Backup2", "Data") & "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb", "*", Text3.Text)
    Else
        NamaFileReset = Replace(getRegistry("Reset1", "Data") & "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb", "*", Text3.Text)
    'Text2.Text = getRegistry("Reset2", "Data")
        NamaFileBackup = Replace(getRegistry("Backup", "Data") & "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb", "*", Text3.Text)
    End If
    
'    NamaFileReset = DIRPOS & "\Reset\K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb"
'    NamaFileBackup = DIRPOS & "\Backup\K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb"
    'RESET
    cekFile = "Reset"
    If Dir(NamaFileReset) <> "" Then
     Set dbsClear = OpenDatabase(NamaFileReset)
     dbsClear.Execute "Delete * FROM MSALES"
     dbsClear.Execute "Delete * FROM MSALESD"
     dbsClear.Close
    Else
     If KodeKasir = KodeUserDua Then
        Create_DatabaseResetDua NamaFileReset
     Else
        Create_DatabaseReset NamaFileReset
      End If
    End If
    cekFile = "Backup"
    'If KodeKasir <> KodeUserDua Then
    
     If Dir(NamaFileBackup) <> "" Then
         On Error Resume Next
         Kill (NamaFileBackup)
Create_DatabaseBackup NamaFileBackup
         Set dbsClear = OpenDatabase(NamaFileBackup)
         dbsClear.Execute "Delete * FROM MSALES"
         dbsClear.Execute "Delete * FROM MSALESD"
         dbsClear.Close
         On Error GoTo 0
        Else
          Create_DatabaseBackup NamaFileBackup
        End If
               
        'BackUp
        dbs.Execute "INSERT INTO MSales ( Shift, IDUser,IDBank, Tanggal, NoID, Kode, TotalPajak, TotalDiscount, SubTotal, DiscIntern,DiscNota,JumDiscInternRp, HargaTotal, IDPayment, UangMuka, IsSend, Bank, ISPending,Voucher,IsUpload,IDMember,BarangKSB,SisaKSB,Pembulatan ) IN '" & NamaFileBackup & "' " & _
                  "SELECT MSales.Shift, MSales.IDUser,MSales.IDBank, MSales.Tanggal, MSales.NoID, MSales.Kode, MSales.TotalPajak, MSales.TotalDiscount, MSales.SubTotal,MSales.DiscIntern,MSales.JumDiscInternRp, MSales.DiscNota, MSales.HargaTotal, MSales.IDPayment, MSales.UangMuka, MSales.IsSend, MSales.Bank, MSales.ISPending , MSales.Voucher,MSales.IsUpload,MSales.IDMember,MSales.BarangKSB,MSales.SisaKSB,MSales.Pembulatan " & _
                  "From MSales " & _
                  "WHERE (((MSales.Shift)=" & NamaShiftReset & ") AND  (Day(MSales.Tanggal)=" & DayReset & " and Month(MSales.Tanggal)=" & MonthReset & " and Year(MSales.Tanggal)=" & YearReset & "))"
        
        dbs.Execute "INSERT INTO MSalesD IN '" & NamaFileBackup & "' " & _
                  "SELECT MSalesD.*,MsalesD.Qty*MSalesD.Harga as Jumlah " & _
                  "FROM MSales INNER JOIN MSalesD ON MSales.NoID = MSalesD.IDSales " & _
                  "WHERE (((MSales.Shift)=" & NamaShiftReset & ") AND (Day(MSales.Tanggal)=" & DayReset & " and Month(MSales.Tanggal)=" & MonthReset & " and Year(MSales.Tanggal)=" & YearReset & "))"
         dbs.Close
         
      'End If
      
  'RESET
    Set DbsReset = OpenDatabase(NamaFileReset)
    Set RsReset = DbsReset.OpenRecordset("MSales")
    Set dbs = OpenDatabase(DIRPOS & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
    Set rs = dbs.OpenRecordset("SELECT   sum(SubTotal) as SubTotala,sum(DiscNota) as DiscNotaa,sum(TotalDiscount) as DiscBrga, " & _
            "sum(HargaTotal) as HargaTotala, sum(UangMuka)-sum(Bank) as Tunaia, sum(Voucher) as Vouchera," & _
            "sum(Bank) as Banka,IDBank,IDUser,count(NoID) as JumlahNotaa , Max(HargaTotal+DiscNota+TotalDiscount) as NotaMax,Min(HargaTotal+DiscNota+TotalDiscount) as NotaMin " & _
            "From MSales " & _
              "WHERE (((MSales.Shift)=" & NamaShiftReset & ") AND  (Day(MSales.Tanggal)=" & DayReset & " and Month(MSales.Tanggal)=" & MonthReset & " and Year(MSales.Tanggal)=" & YearReset & ") AND (SubTotal<>0)) " & _
              "Group By day(Tanggal),month(Tanggal),year(Tanggal),IDBank,IDUser HAVING sum(HargaTotal)<>0")
     If rs.EOF And rs.BOF Then
    
     Else
        rs.MoveFirst
        NoID = 1
        Do While Not rs.EOF
            RsReset.AddNew
            RsReset!kode = "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & Format(NoID, "0##0")
            RsReset!NoID = NoID
            RsReset!TANGGAL = Mid(Text2.Text, 3, 2) & "/" & Mid(Text2.Text, 1, 2) & "/" & Mid(Text2.Text, 5, 4)
            RsReset!SubTotal = rs!HargaTotala + rs!DiscNotaa + rs!DiscBrga 'rs!SubTotala + rs!DiscBrga
            RsReset!DiscNota = 0 'rs!DiscNotaa diskon dipaksa 0
            RsReset!Hargatotal = rs!HargaTotala + rs!DiscNotaa + rs!DiscBrga
            RsReset!Tunai = rs!Tunaia + rs!DiscBrga + rs!DiscNotaa
            RsReset!Voucher = rs!Vouchera
            RsReset!Bank = rs!Banka
            RsReset!IDBank = rs!IDBank
            RsReset!IDUser = rs!IDUser
            RsReset!NamaUser = namakasirReset
            RsReset!Shift = NamaShiftReset
            RsReset!IdPengawas = IDUser
            RsReset!NamaPengawas = NamaKasir
            If KodeKasir = KodeUserDua Then
            MaxTotal = PersenLap * (rs!Tunaia + rs!DiscBrga + rs!DiscNotaa) / 100
            Else
                RsReset!JumlahNota = rs!JumlahNotaa
                RsReset!PajakPersen = PersenLap
                RsReset!NotaMin = rs!NotaMin
                RsReset!NotaMax = rs!NotaMax
                RsReset!TunaiPajak = PersenLap * (rs!Tunaia + rs!DiscBrga + rs!DiscNotaa) / 100
                MaxTotal = (rs!Tunaia + rs!DiscBrga + rs!DiscNotaa) '100 persen
            End If
            
            
            RsReset.Update
            RsReset.Bookmark = RsReset.LastModified
            
            If rs!Banka = 0 And KodeKasir = KodeUserDua Then 'Tunai dan kodeuserdua
                Dim RSp As Recordset
                Set RSp = dbs.OpenRecordset("SELECT " & NoID & " As IDSales, MSalesD.IDInventor, Sum(MSalesD.Qty) AS Qty, Sum((MSalesD.HargaBruto*MSalesD.Qty)) AS Jumlah, Sum((MSalesD.HargaPokok*MSalesD.Qty)) AS HPP, MSalesD.KodeInv, MSalesD.NamaInv, MSalesD.Satuan, MSalesD.Barcode, MSalesD.IDSatuan, MSalesD.Konversi " & _
                "FROM MSalesD INNER JOIN MSales ON MSalesD.IDSales = MSales.NoID " & _
                "Where day(MSales.Tanggal) = " & DayReset & " and Month(MSales.Tanggal) = " & MonthReset & " and Year(MSales.Tanggal) = " & YearReset & "  And MSales.ShifT = " & NamaShiftReset & " " & _
                "GROUP BY MSales.IDUser,Msales.IDBank,MSalesD.IDInventor, MSalesD.KodeInv, MSalesD.NamaInv, MSalesD.Satuan, MSalesD.Barcode, MSalesD.IDSatuan, MSalesD.Konversi " & _
                "HAVING (((Sum(MSalesD.Qty))<>0) AND MSales.IDUser=" & rs!IDUser & " AND MSales.IDBank=" & NullToNol(rs!IDBank) & ") " & _
                "ORDER BY Max(MSales.Tanggal) DESC ")
                RSp.MoveFirst
                curTotal = 0
                Do While Not RSp.EOF
                    If MaxTotal <= curTotal Then Exit Do
                    curQty = ((RSp!QTY * PersenLap - RSp!QTY * PersenLap Mod 100) \ 100) + IIf(RSp!QTY * PersenLap Mod 100 > 1, 1, 0)
                    CurhargaSatuan = CLng(IIf(RSp!QTY <> 0, RSp!Jumlah / RSp!QTY, 0) * curQty)
                    If CurhargaSatuan Mod 50 > 0 Then CurhargaSatuan = CurhargaSatuan + (50 - (CurhargaSatuan Mod 50))
                    curTotal = curTotal + CurhargaSatuan
                      dbs.Execute "INSERT INTO MSalesD ( IDSales,IDInventor, Qty,HargaSat, Harga, HargaPokok," & _
                                  "KodeInv, NamaInv, Satuan, Barcode, IDSatuan, Konversi ) " & _
                                  "IN '" & NamaFileReset & "' " & _
                                  "VALUES(" & _
                                   NoID & "," & RSp!IdInventor & "," & curQty & "," & Replace(Format(IIf(curQty = 0, CurhargaSatuan, Abs(CurhargaSatuan / curQty)), "##0.00"), ",", ".") & "," & Replace(Format(CurhargaSatuan, "##0.00"), ",", ".") & ",0,'" & _
                                  Replace(RSp!KodeInv, "'", "''") & "','" & Replace(RSp!NamaInv, "'", "''") & "','" & Replace(RSp!Satuan, "'", "''") & "','" & RSp!Barcode & "'," & RSp!idSatuan & "," & RSp!Konversi & ")"

                    RSp.MoveNext
                Loop
                RsReset.Edit
                
                RsReset!SubTotal = curTotal
                'RsReset!DiscNota = curTotal
                RsReset!Hargatotal = curTotal
                RsReset!Tunai = curTotal '- rs!DiscNotaa
                RsReset.Update
            Else
              dbs.Execute "INSERT INTO MSalesD ( IDSales,IDInventor, Qty, Harga, HargaPokok, KodeInv, NamaInv, Satuan, Barcode, IDSatuan, Konversi ) IN '" & NamaFileReset & "'" & _
                "SELECT " & NoID & " As IDSales, MSalesD.IDInventor, Sum(MSalesD.Qty) AS Qty, Sum((MSalesD.HargaBruto*MSalesD.Qty)) AS Jumlah, Sum((MSalesD.HargaPokok*MSalesD.Qty)) AS HPP, MSalesD.KodeInv, MSalesD.NamaInv, MSalesD.Satuan, MSalesD.Barcode, MSalesD.IDSatuan, MSalesD.Konversi " & _
                "FROM MSalesD INNER JOIN MSales ON MSalesD.IDSales = MSales.NoID " & _
                "Where day(MSales.Tanggal) = " & DayReset & " and Month(MSales.Tanggal) = " & MonthReset & " and Year(MSales.Tanggal) = " & YearReset & "  And MSales.ShifT = " & NamaShiftReset & " " & _
                "GROUP BY MSales.IDUser,Msales.IDBank,MSalesD.IDInventor, MSalesD.KodeInv, MSalesD.NamaInv, MSalesD.Satuan, MSalesD.Barcode, MSalesD.IDSatuan, MSalesD.Konversi " & _
                "HAVING (((Sum(MSalesD.Qty))<>0) AND MSales.IDUser=" & rs!IDUser & " AND MSales.IDBank=" & NullToNol(rs!IDBank) & ")"
            DbsReset.Execute "UPDATE MSalesD SET MSalesD.HargaSat = [msalesd].[harga]/Abs([msalesd].[qty]) WHERE (((MSalesD.Qty)<>0)) AND IDSales=" & NoID
            End If
            DbsReset.Execute "Update MSALES  set Tanggal=#" & Mid(Text2.Text, 3, 2) & "/" & Mid(Text2.Text, 1, 2) & "/" & Mid(Text2.Text, 5, 4) & "#"
    '            If rs!Banka = 0 Then
    '                If KodeKasir = KodeUserDua Then
    '                  DbsReset.Execute "UPDATE MSalesD SET MSalesD.Qty=(MSalesD.Qty*" & PersenLap & ")\100,MSalesD.[harga]=((MSalesD.Qty*" & PersenLap & ")\100)*MSalesD.HargaSat WHERE  IDSales=" & NoId
    '                 End If
    '            End If
                NoID = NoID + 1
        rs.MoveNext
        Loop
     End If
 
    
    Set rs = Nothing
    dbs.Close
    Set RsReset = Nothing
    DbsReset.Close
    Set dbs = Nothing
    Set DbsReset = Nothing
    Exit Sub
pesan:
'   MsgBox "Ada kesalahan : " & Err.Description & vbCrLf & "Silahkan Tekan OK", vbCritical + vbOKOnly, "Pesan!"
   MsgBox "Ada kesalahan : " & Err.Description & vbCrLf & "Silahkan Tekan OK", vbCritical + vbOKOnly, "Pesan!"
    
   If Err.Number = 52 Then
    If cekFile = "Backup" Then
        NamaFileBackup = NMBACKUPGAGAL & "\" & "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb"
        Resume
    ElseIf cekFile = "Reset" Then
        NamaFileBackup = NMRESETGAGAL & "\" & "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb"
        Resume
    End If
   Else
  
    Resume Next
    End If
'    Unload Me
End Sub
Sub ResetBusana() 'diskon sebagai diskon supplier jadi dianggap tidak ada diskon (diskon dikembaikan)
'On Error GoTo pesan
'Dim KodeKasirReset As String
'    Dim namakasirReset As String
'    Dim dbs As Database
'    Dim rs As Recordset
'    Dim DbsReset As Database
'    Dim RsReset As Recordset
'    Dim psn As String
'  Dim i As Integer
'  Dim NoId As Integer
'  Dim idkasirreset As Integer
'  Dim NamaFileReset As String
'  Dim NamaFileBackup As String
'  Dim TunaiPajak As Long
'  Dim cTunai As Long
'  Dim cJumlahNota As Long
'  Dim cbank As Double
'  Dim cSubtotal As Double
'  Dim cUangMuka As Double
'  Dim cTotal As Double
'  Dim cDiskonNota As Double
'  Dim cDiskonBrg As Double
'  Dim cVoucher As Double
'  Dim dbsClear As Database
'  Dim strqry As String
'  Dim MaxTotal As Double
'  Dim curTotal As Double
'  Dim curQty As Long
'  Dim CurhargaSatuan As Double
'  Dim cekFile As String
'  Dim JmlSalah As Integer
'  Dim Bulat As Long
'    lblNama(9).Caption = "Biaya Credit Card"
'    Set dbs = OpenDatabase(DIRPOS & "\database\TempDB"& format(now,"_yyyyMM") &".mdb")
'    dbs.Execute "Update Msales Set IDbank =2 where isnull(idbank)"
'    'JumlahSubTotal adalah bruto (sebelum diskon)
'    strqry = "SELECT MSales.IDUser,Count(MSales.Kode) AS JumlahNota, sum(MSales.JumDiscInternRp)+ sum(MSales.TotalDiscount)+ Sum(MSales.SubTotal) AS JumlahSubTotal,Sum(MSales.HargaTotal) AS JumlahTotal, sum(MSales.DiscNota) as JumDiskon, sum(MSales.TotalDiscount) as DiskonBrg, Sum(MSales.UangMuka-MSales.Bank-MSales.Voucher) AS JumUangMuka, Sum(MSales.Bank) AS JumBank , Sum(MSales.Voucher) AS jumVoucher, " & _
'            "sum(MSales.JumDiscInternRp) as DiscBarangIntern,sum(MSales.DiscIntern) as DiscNotaIntern, " & _
'            "Sum(MSales.Pembulatan) as Bulat " & _
'            "From MSales " & _
'            "WHERE Day(Tanggal)=" & DayReset & " AND Month(Tanggal)=" & MonthReset & " AND Year(Tanggal)=" & YearReset & " and SubTotal<>0 AND Shift= " & NamaShiftReset & " " & _
'            "GROUP BY MSales.IDUser"
'    Set rs = dbs.OpenRecordset(strqry)
'    If rs.EOF And rs.BOF Then
'      Exit Sub
'    Else
'    rs.MoveFirst
'       PrinBigChar Chr(13) & Chr(10) & Chr(13) & "   KASSA : " & Str(namaMesinReset)
'       psn = Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "----------------RESET----------------" & Chr(13) & Chr(10) & Chr(13) & Chr(10)
'       Prin psn
'       psn = ""
'    Do While Not rs.EOF
'      idkasirreset = rs!IDUser
'        cariKodeNama KodeKasirReset, namakasirReset, rs!IDUser
'        If KodeKasir = KodeUserDua Then
'            TunaiPajak = ((PersenLap / 100) * rs!JumUangMuka) - (((PersenLap / 100) * rs!JumUangMuka) Mod 100)
'            cTunai = rs!JumUangMuka
'            cJumlahNota = rs!JumlahNota
'            cbank = rs!Jumbank
'            'cVoucher = rs!JumVoucher
'            cVoucher = 0
'            cDiskonNota = rs!JumDIskon
'            cSubtotal = cTunai + cbank + cDiskonNota + cVoucher + rs!DIskonBrg + rs!Bulat
'            lbl(0) = Format(DayReset, "0#") & "-" & Format(MonthReset, "0#") & "-" & Format(YearReset, "####") & ", " & IIf(NamaShiftReset = 1, "PAGI", "SORE") 'Format(Now, "dd-MM-yyyy, hh:mm:ss")
'            lbl(1) = namaMesinReset & " - " & namakasirReset
'            lbl(2) = Format(rs!JumlahNota, "###,##0")
'            'lbl(3) = Format(rs!Jumbank + TunaiPajak + cVoucher + rs!JumDIskon + rs!DIskonBrg + rs!DiscBarangIntern + rs!DiscNotaIntern, "###,###,##0")
'            lbl(3) = Format(rs!Jumbank + TunaiPajak + cVoucher + rs!JumDIskon + rs!DIskonBrg, "###,###,##0")
'            lbl(4) = Format(rs!JumDIskon + rs!DIskonBrg, "###,###,##0") & " + " & vbCrLf & _
'                     Format(rs!DiscBarangIntern + rs!DiscNotaIntern, "###,###,##0") & " = " & _
'                     Format(rs!JumDIskon + rs!DIskonBrg + rs!DiscBarangIntern + rs!DiscNotaIntern, "###,###,##0")
'
'            lbl(5) = Format(rs!Jumbank + TunaiPajak, "###,###,##0")
'            lbl(6) = Format(TunaiPajak, "###,###,##0")
'            lbl(7) = Format(rs!Jumbank, "###,###,##0")
'            lbl(8) = Format(cVoucher, "###,###,##0")
'            'lbl(9) = (rs!Jumbank + TunaiPajak)
'            Bulat = Format(rs!Bulat, "###,###,##0")
'
'        Else
'            'lbl(0) = Format(Now, "dd-MM-yyyy, hh:mm:ss")
'            lbl(0) = Format(DayReset, "0#") & "-" & Format(MonthReset, "0#") & "-" & Format(YearReset, "####") & ", " & IIf(NamaShiftReset = 1, "PAGI", "SORE") 'Format(Now, "dd-MM-yyyy, hh:mm:ss")
'            lbl(1) = namaMesinReset & " - " & namakasirReset
'            lbl(2) = Format(rs!JumlahNota, "###,##0")
'            'Diskon intern dan diskon extern di anggap dari supplier
'            'lbl(3) = Format(rs!JumlahSubTotal + rs!DIskonBrg + rs!DiscBarangIntern + rs!DiscNotaIntern, "###,###,##0")
'             ''Subtotal sudah dipotong disko intern DARI BRUTO DAN DITAMBAH DISKON EXTERN
'             lbl(3) = Format(rs!JumlahSubTotal, "###,###,##0") '- rs!DiscBarangIntern - rs!DiscNotaIntern
'             lbl(4) = Format(rs!JumDIskon + rs!DIskonBrg, "###,###,##0") & " + " & vbCrLf & _
'                     Format(rs!DiscBarangIntern + rs!DiscNotaIntern, "###,###,##0") & " = " & _
'                     Format(rs!JumDIskon + rs!DIskonBrg + rs!DiscBarangIntern + rs!DiscNotaIntern, "###,###,##0")
'            lbl(5) = Format(rs!JumlahTotal + rs!JumDIskon + rs!DIskonBrg, "###,###,##0")
'            lbl(6) = Format(rs!JumUangMuka, "###,###,##0")
'            lbl(7) = Format(rs!Jumbank, "###,###,##0")
'            lbl(8) = Format(rs!JumVoucher, "###,###,##0")
'            lbl(9) = Format(rs!JumlahTotal - (rs!JumlahSubTotal - rs!JumDIskon) + rs!DIskonBrg + rs!DiscNotaIntern + rs!Bulat, "###,###,##0")
'            TunaiPajak = ((PersenLap / 100) * rs!JumUangMuka) - (((PersenLap / 100) * rs!JumUangMuka) Mod 100)
'            cVoucher = rs!JumVoucher
'            cTunai = rs!JumUangMuka
'            cJumlahNota = rs!JumlahNota
'            cbank = rs!Jumbank
'            cDiskonNota = rs!JumDIskon
'            cDiskonBrg = rs!DIskonBrg
'            cSubtotal = cTunai + cbank + cDiskonNota + cVoucher + cDiskonBrg + rs!Bulat
'            Bulat = Format(rs!Bulat, "###,###,##0")
'
'        End If
'            If KodeKasir = KodeUserDua Then
'              For i = 0 To 7
'                 psn = psn & lblNama(i).Caption & Space(18 - Len(Trim(lblNama(i).Caption))) & ": " & lbl(i).Caption & Chr(13) & Chr(10) & Chr(13) & Chr(10)
'                  'If i = 4 Then
'                  '    psn = psn & "Pembulatan" & Space(18 - Len("Pembulatan")) & ": " & Format(Bulat, "###,##0") & Chr(13) & Chr(10) & Chr(13) & Chr(10)
'                  'End If
'              Next
'            Else
'              For i = 0 To 9
'               psn = psn & lblNama(i).Caption & Space(18 - Len(Trim(lblNama(i).Caption))) & ": " & lbl(i).Caption & Chr(13) & Chr(10) & Chr(13) & Chr(10)
'               If i = 4 Then
'                      psn = psn & "Pembulatan" & Space(18 - Len("Pembulatan")) & ": " & Format(Bulat, "###,##0") & Chr(13) & Chr(10) & Chr(13) & Chr(10)
'                  End If
'              Next
'              End If
'              Prin psn
'              psn = ""
'        rs.MoveNext
'    Loop
'End If
'    psn = psn & "----------------------------------------" & Chr(13) & Chr(10)
'    psn = psn & CekKassaShift1 & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10)
'
'    Prin psn
'    PrinBigChar Space(15 - Min(15, Len(NamaToko))) & UCase(NamaToko)
'    papercut
'    Prin Judulstruk & Chr(13) & Chr(10) & "----------------------------------------"
'    'MODIFIKASI 20-03-2007
'    'Dijadikan 2 , server dan lokal
'    '
'    If KodeKasir = KodeUserDua Then
'        NamaFileReset = Replace(getRegistry("Reset2", "Data") & "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb", "*", Text3.Text)
'        NamaFileBackup = Replace(getRegistry("Backup2", "Data") & "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb", "*", Text3.Text)
'    Else
'        NamaFileReset = Replace(getRegistry("Reset1", "Data") & "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb", "*", Text3.Text)
'    'Text2.Text = getRegistry("Reset2", "Data")
'        NamaFileBackup = Replace(getRegistry("Backup", "Data") & "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb", "*", Text3.Text)
'    End If
'
''    NamaFileReset = DIRPOS & "\Reset\K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb"
''    NamaFileBackup = DIRPOS & "\Backup\K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb"
'    'RESET
'    cekFile = "Reset"
'    If Dir(NamaFileReset) <> "" Then
'     Set dbsClear = OpenDatabase(NamaFileReset)
'     dbsClear.Execute "Delete * FROM MSALES"
'     dbsClear.Execute "Delete * FROM MSALESD"
'     dbsClear.Close
'    Else
'     If KodeKasir = KodeUserDua Then
'        Create_DatabaseResetDua NamaFileReset
'     Else
'        Create_DatabaseReset NamaFileReset
'      End If
'    End If
'    cekFile = "Backup"
'    'If KodeKasir <> KodeUserDua Then
'        If Dir(NamaFileBackup) <> "" Then
'         On Error Resume Next
'         Kill (NamaFileBackup)
'         Create_DatabaseBackup NamaFileBackup
'         Set dbsClear = OpenDatabase(NamaFileBackup)
'         dbsClear.Execute "Delete * FROM MSALES"
'         dbsClear.Execute "Delete * FROM MSALESD"
'         dbsClear.Close
'         On Error GoTo 0
'        Else
'          Create_DatabaseBackup NamaFileBackup
'        End If
'
'        'BackUp
'        dbs.Execute "INSERT INTO MSales ( Shift, IDUser,IDBank, Tanggal, NoID, Kode, TotalPajak, TotalDiscount, SubTotal, DiscIntern,DiscNota,JumDiscInternRp, HargaTotal, IDPayment, UangMuka, IsSend, Bank, ISPending,Voucher,IsUpload,IDMember,BarangKSB,SisaKSB,Pembulatan ) IN '" & NamaFileBackup & "' " & _
'                  "SELECT MSales.Shift, MSales.IDUser,MSales.IDBank, MSales.Tanggal, MSales.NoID, MSales.Kode, MSales.TotalPajak, MSales.TotalDiscount, MSales.SubTotal,MSales.DiscIntern,MSales.JumDiscInternRp, MSales.DiscNota, MSales.HargaTotal, MSales.IDPayment, MSales.UangMuka, MSales.IsSend, MSales.Bank, MSales.ISPending , MSales.Voucher,MSales.IsUpload,MSales.IDMember,MSales.BarangKSB,MSales.SisaKSB,MSales.Pembulatan " & _
'                  "From MSales " & _
'                  "WHERE (((MSales.Shift)=" & NamaShiftReset & ") AND  (Day(MSales.Tanggal)=" & DayReset & " and Month(MSales.Tanggal)=" & MonthReset & " and Year(MSales.Tanggal)=" & YearReset & "))"
'
'        dbs.Execute "INSERT INTO MSalesD IN '" & NamaFileBackup & "' " & _
'                  "SELECT MSalesD.*,MsalesD.Qty*MSalesD.Harga as Jumlah " & _
'                  "FROM MSales INNER JOIN MSalesD ON MSales.NoID = MSalesD.IDSales " & _
'                  "WHERE (((MSales.Shift)=" & NamaShiftReset & ") AND (Day(MSales.Tanggal)=" & DayReset & " and Month(MSales.Tanggal)=" & MonthReset & " and Year(MSales.Tanggal)=" & YearReset & "))"
'         dbs.Close
'    'End If
'
'  'RESET
'    Set DbsReset = OpenDatabase(NamaFileReset)
'    Set RsReset = DbsReset.OpenRecordset("MSales")
'    Set dbs = OpenDatabase(DIRPOS & "\database\TempDB"& format(now,"_yyyyMM") &".mdb")
'    Set rs = dbs.OpenRecordset("SELECT   sum(SubTotal) as SubTotala,sum(DiscNota) as DiscNotaa,sum(TotalDiscount) as DiscBrga, " & _
'            "sum(HargaTotal) as HargaTotala, sum(UangMuka)-sum(Bank) as Tunaia, sum(Voucher) as Vouchera," & _
'            "sum(Bank) as Banka,IDBank,IDUser,count(NoID) as JumlahNotaa , Max(HargaTotal+DiscNota+TotalDiscount) as NotaMax,Min(HargaTotal+DiscNota+TotalDiscount) as NotaMin " & _
'            "From MSales " & _
'              "WHERE (((MSales.Shift)=" & NamaShiftReset & ") AND  (Day(MSales.Tanggal)=" & DayReset & " and Month(MSales.Tanggal)=" & MonthReset & " and Year(MSales.Tanggal)=" & YearReset & ") AND (SubTotal<>0)) " & _
'              "Group By day(Tanggal),month(Tanggal),year(Tanggal),IDBank,IDUser HAVING sum(HargaTotal)<>0")
'     If rs.EOF And rs.BOF Then
'
'     Else
'        rs.MoveFirst
'        NoId = 1
'        Do While Not rs.EOF
'            RsReset.AddNew
'            RsReset!kode = "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & Format(NoId, "0##0")
'            RsReset!NoId = NoId
'            RsReset!TANGGAL = Mid(Text2.Text, 3, 2) & "/" & Mid(Text2.Text, 1, 2) & "/" & Mid(Text2.Text, 5, 4)
'            RsReset!SubTotal = rs!HargaTotala + rs!DiscNotaa + rs!DiscBrga 'rs!SubTotala + rs!DiscBrga
'            RsReset!DiscNota = 0 'rs!DiscNotaa diskon dipaksa 0
'            RsReset!Hargatotal = rs!HargaTotala + rs!DiscNotaa + rs!DiscBrga
'            RsReset!Tunai = rs!Tunaia + rs!DiscBrga + rs!DiscNotaa
'            RsReset!Voucher = rs!Vouchera
'            RsReset!Bank = rs!Banka
'            RsReset!IDBank = rs!IDBank
'            RsReset!IDUser = rs!IDUser
'            RsReset!NamaUser = namakasirReset
'            RsReset!Shift = NamaShiftReset
'            RsReset!IdPengawas = IDUser
'            RsReset!NamaPengawas = NamaKasir
'            If KodeKasir = KodeUserDua Then
'            MaxTotal = PersenLap * (rs!Tunaia + rs!DiscBrga + rs!DiscNotaa) / 100
'            Else
'                RsReset!JumlahNota = rs!JumlahNotaa
'                RsReset!PajakPersen = PersenLap
'                RsReset!NotaMin = rs!NotaMin
'                RsReset!NotaMax = rs!NotaMax
'                RsReset!TunaiPajak = PersenLap * (rs!Tunaia + rs!DiscBrga + rs!DiscNotaa) / 100
'                MaxTotal = (rs!Tunaia + rs!DiscBrga + rs!DiscNotaa) '100 persen
'            End If
'
'
'            RsReset.Update
'            RsReset.Bookmark = RsReset.LastModified
'
'            If rs!Banka = 0 And KodeKasir = KodeUserDua Then 'Tunai dan kodeuserdua
'                Dim RSp As Recordset
'                Set RSp = dbs.OpenRecordset("SELECT " & NoId & " As IDSales, MSalesD.IDInventor, Sum(MSalesD.Qty) AS Qty, Sum(MSalesD.Qty*((MSalesD.HargaBruto-MSalesD.DiscInternRp)-iif((MSales.SubTotal+MSales.TotalDiscount+Msales.JumDiscInternRp)<>0,MSalesD.HargaBruto*(MSales.DiscIntern/(MSales.SubTotal+MSales.TotalDiscount+Msales.JumDiscInternRp)),0))) AS Jumlah, Sum((MSalesD.HargaPokok*MSalesD.Qty)) AS HPP, MSalesD.KodeInv, MSalesD.NamaInv, MSalesD.Satuan, MSalesD.Barcode, MSalesD.IDSatuan, MSalesD.Konversi " & _
'                "FROM MSalesD INNER JOIN MSales ON MSalesD.IDSales = MSales.NoID " & _
'                "Where day(MSales.Tanggal) = " & DayReset & " and Month(MSales.Tanggal) = " & MonthReset & " and Year(MSales.Tanggal) = " & YearReset & "  And MSales.ShifT = " & NamaShiftReset & " " & _
'                "GROUP BY MSales.IDUser,Msales.IDBank,MSalesD.IDInventor, MSalesD.KodeInv, MSalesD.NamaInv, MSalesD.Satuan, MSalesD.Barcode, MSalesD.IDSatuan, MSalesD.Konversi " & _
'                "HAVING (((Sum(MSalesD.Qty))<>0) AND MSales.IDUser=" & rs!IDUser & " AND MSales.IDBank=" & NullToNol(rs!IDBank) & ") " & _
'                "ORDER BY Max(MSales.Tanggal) DESC ")
'                If RSp.EOF And RSp.BOF Then
'                Else
'                    RSp.MoveFirst
'                    curTotal = 0
'                    Do While Not RSp.EOF
'                        If MaxTotal <= curTotal Then Exit Do
'                        curQty = ((RSp!QTY * PersenLap - RSp!QTY * PersenLap Mod 100) \ 100) + IIf(RSp!QTY * PersenLap Mod 100 > 1, 1, 0)
'                        CurhargaSatuan = CLng(IIf(RSp!QTY <> 0, RSp!Jumlah / RSp!QTY, 0) * curQty)
'                        If CurhargaSatuan Mod 50 > 0 Then CurhargaSatuan = CurhargaSatuan + (50 - (CurhargaSatuan Mod 50))
'                        curTotal = curTotal + CurhargaSatuan
'                          dbs.Execute "INSERT INTO MSalesD ( IDSales,IDInventor, Qty,HargaSat, Harga, HargaPokok," & _
'                                      "KodeInv, NamaInv, Satuan, Barcode, IDSatuan, Konversi ) " & _
'                                      "IN '" & NamaFileReset & "' " & _
'                                      "VALUES(" & _
'                                       NoId & "," & RSp!IdInventor & "," & curQty & "," & Replace(Format(IIf(curQty = 0, CurhargaSatuan, Abs(CurhargaSatuan / curQty)), "##0.00"), ",", ".") & "," & Replace(Format(CurhargaSatuan, "##0.00"), ",", ".") & ",0,'" & _
'                                      Replace(RSp!KodeInv, "'", "''") & "','" & Replace(RSp!NamaInv, "'", "''") & "','" & Replace(RSp!Satuan, "'", "''") & "','" & RSp!BARCODE & "'," & RSp!idSatuan & "," & RSp!Konversi & ")"
'
'                        RSp.MoveNext
'                    Loop
'                    RsReset.Edit
'
'                    RsReset!SubTotal = curTotal
'                    'RsReset!DiscNota = curTotal
'                    RsReset!Hargatotal = curTotal
'                    RsReset!Tunai = curTotal '- rs!DiscNotaa
'                    RsReset.Update
'                End If
'            Else
'              dbs.Execute "INSERT INTO MSalesD ( IDSales,IDInventor, Qty, Harga, HargaPokok, KodeInv, NamaInv, Satuan, Barcode, IDSatuan, Konversi ) IN '" & NamaFileReset & "'" & _
'                "SELECT " & NoId & " As IDSales, MSalesD.IDInventor, Sum(MSalesD.Qty) AS Qty, Sum(MSalesD.Qty*((MSalesD.HargaBruto-MSalesD.DiscInternRp)-iif((MSales.SubTotal+MSales.TotalDiscount+Msales.JumDiscInternRp)<>0,MSalesD.HargaBruto*(MSales.DiscIntern/(MSales.SubTotal+MSales.TotalDiscount+Msales.JumDiscInternRp)),0))) AS Jumlah, Sum((MSalesD.HargaPokok*MSalesD.Qty)) AS HPP, MSalesD.KodeInv, MSalesD.NamaInv, MSalesD.Satuan, MSalesD.Barcode, MSalesD.IDSatuan, MSalesD.Konversi " & _
'                "FROM MSalesD INNER JOIN MSales ON MSalesD.IDSales = MSales.NoID " & _
'                "Where day(MSales.Tanggal) = " & DayReset & " and Month(MSales.Tanggal) = " & MonthReset & " and Year(MSales.Tanggal) = " & YearReset & "  And MSales.ShifT = " & NamaShiftReset & " " & _
'                "GROUP BY MSales.IDUser,Msales.IDBank,MSalesD.IDInventor, MSalesD.KodeInv, MSalesD.NamaInv, MSalesD.Satuan, MSalesD.Barcode, MSalesD.IDSatuan, MSalesD.Konversi " & _
'                "HAVING (((Sum(MSalesD.Qty))<>0) AND MSales.IDUser=" & rs!IDUser & " AND MSales.IDBank=" & NullToNol(rs!IDBank) & ")"
'            DbsReset.Execute "UPDATE MSalesD SET MSalesD.HargaSat = [msalesd].[harga]/Abs([msalesd].[qty]) WHERE (((MSalesD.Qty)<>0)) AND IDSales=" & NoId
'            End If
'            DbsReset.Execute "Update MSALES  set Tanggal=#" & Mid(Text2.Text, 3, 2) & "/" & Mid(Text2.Text, 1, 2) & "/" & Mid(Text2.Text, 5, 4) & "#"
'    '            If rs!Banka = 0 Then
'    '                If KodeKasir = KodeUserDua Then
'    '                  DbsReset.Execute "UPDATE MSalesD SET MSalesD.Qty=(MSalesD.Qty*" & PersenLap & ")\100,MSalesD.[harga]=((MSalesD.Qty*" & PersenLap & ")\100)*MSalesD.HargaSat WHERE  IDSales=" & NoId
'    '                 End If
'    '            End If
'                NoId = NoId + 1
'        rs.MoveNext
'        Loop
'     End If
'
'
'    Set rs = Nothing
'    dbs.Close
'    Set RsReset = Nothing
'    DbsReset.Close
'    Set dbs = Nothing
'    Set DbsReset = Nothing
'    Exit Sub
'pesan:
''   MsgBox "Ada kesalahan : " & Err.Description & vbCrLf & "Silahkan Tekan OK", vbCritical + vbOKOnly, "Pesan!"
'   If Err.Number = 52 Then
'    If cekFile = "Backup" Then
'        NamaFileBackup = NMBACKUPGAGAL & "\" & "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb"
'        Resume
'    ElseIf cekFile = "Reset" Then
'        NamaFileBackup = NMRESETGAGAL & "\" & "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb"
'        Resume
'    End If
'   Else
'    Resume Next
'    End If
''    Unload Me
On Error GoTo pesan

    Dim dbs As Database
    Dim rs As Recordset
    Dim DbsReset As Database
    Dim RsReset As Recordset
    Dim psn As String
  Dim i As Integer
  Dim NoID As Integer
  Dim idkasirreset As Integer
  Dim NamaFileReset As String
  Dim NamaFileBackup As String
  Dim TunaiPajak As Long
  Dim cTunai As Long
  Dim cJumlahNota As Long
  Dim cbank As Double
  Dim cSubtotal As Double
  Dim cUangMuka As Double
  Dim cTotal As Double
  Dim cDiskonNota As Double
  Dim cDiskonBrg As Double
  Dim cVoucher As Double
  Dim dbsClear As Database
  Dim strqry As String
  Dim MaxTotal As Double
  Dim curTotal As Double
  Dim curQty As Long
  Dim CurhargaSatuan As Double
  Dim cekFile As String
  Dim JmlSalah As Integer
  Dim Bulat As Long
    lblNama(9).Caption = "Biaya Credit Card"
    Set dbs = OpenDatabase(DIRPOS & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
    dbs.Execute "Update Msales Set IDbank =2 where isnull(idbank)"
    'JumlahSubTotal adalah bruto (sebelum diskon)
    strqry = "SELECT MSales.IDUser,Count(MSales.Kode) AS JumlahNota, sum(MSales.TotalDiscount)+ Sum(MSales.SubTotal) AS JumlahSubTotal,Sum(MSales.HargaTotal) AS JumlahTotal, sum(MSales.DiscNota) as JumDiskon, sum(MSales.TotalDiscount) as DiskonBrg, Sum(MSales.UangMuka-MSales.Bank-MSales.Voucher) AS JumUangMuka, Sum(MSales.Bank) AS JumBank , Sum(MSales.Voucher) AS jumVoucher, " & _
            "sum(MSales.JumDiscInternRp) as DiscBarangIntern,sum(MSales.DiscIntern) as DiscNotaIntern, " & _
            "Sum(MSales.Pembulatan) as Bulat " & _
            "From MSales " & _
            "WHERE Day(Tanggal)=" & DayReset & " AND Month(Tanggal)=" & MonthReset & " AND Year(Tanggal)=" & YearReset & " and SubTotal<>0 AND Shift= " & NamaShiftReset & " " & _
            "GROUP BY MSales.IDUser"
    Set rs = dbs.OpenRecordset(strqry)
    If rs.EOF And rs.BOF Then
      Exit Sub
    Else
    rs.MoveFirst
       PrinBigChar Chr(13) & Chr(10) & Chr(13) & "   KASSA : " & Str(namaMesinReset)
       psn = Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "----------------RESET----------------" & Chr(13) & Chr(10) & Chr(13) & Chr(10)
       Prin psn
       psn = ""
    Do While Not rs.EOF
      idkasirreset = rs!IDUser
        cariKodeNama KodeKasirReset, namakasirReset, rs!IDUser
        If KodeKasir = KodeUserDua Then
            TunaiPajak = ((PersenLap / 100) * rs!JumUangMuka) - (((PersenLap / 100) * rs!JumUangMuka) Mod 100)
            cTunai = rs!JumUangMuka
            cJumlahNota = rs!JumlahNota
            cbank = rs!Jumbank
            'cVoucher = rs!JumVoucher
            cVoucher = 0
            cDiskonNota = rs!JumDIskon
            cSubtotal = cTunai + cbank + cDiskonNota + cVoucher + rs!DIskonBrg + rs!Bulat
            lbl(0) = Format(DayReset, "0#") & "-" & Format(MonthReset, "0#") & "-" & Format(YearReset, "####") & ", " & IIf(NamaShiftReset = 1, "PAGI", "SORE") 'Format(Now, "dd-MM-yyyy, hh:mm:ss")
            lbl(1) = namaMesinReset & " - " & namakasirReset
            lbl(2) = Format(rs!JumlahNota, "###,##0")
            'lbl(3) = Format(rs!Jumbank + TunaiPajak + cVoucher + rs!JumDIskon + rs!DIskonBrg + rs!DiscBarangIntern + rs!DiscNotaIntern, "###,###,##0")
            lbl(3) = Format(rs!Jumbank + TunaiPajak + cVoucher + rs!JumDIskon + rs!DIskonBrg, "###,###,##0")
            lbl(4) = Format(rs!JumDIskon + rs!DIskonBrg, "###,###,##0") & " + " & vbCrLf & _
                     Format(rs!DiscBarangIntern + rs!DiscNotaIntern, "###,###,##0") & " = " & _
                     Format(rs!JumDIskon + rs!DIskonBrg + rs!DiscBarangIntern + rs!DiscNotaIntern, "###,###,##0")
                     
            lbl(5) = Format(rs!Jumbank + TunaiPajak, "###,###,##0")
            lbl(6) = Format(TunaiPajak, "###,###,##0")
            lbl(7) = Format(rs!Jumbank, "###,###,##0")
            lbl(8) = Format(cVoucher, "###,###,##0")
            'lbl(9) = (rs!Jumbank + TunaiPajak)
            Bulat = Format(rs!Bulat, "###,###,##0")
            
        Else
'            strqry = "SELECT MSales.IDUser,Count(MSales.Kode) AS JumlahNota, sum(MSales.JumDiscInternRp)+ sum(MSales.TotalDiscount)+ Sum(MSales.SubTotal) AS JumlahSubTotal,Sum(MSales.HargaTotal) AS JumlahTotal, sum(MSales.DiscNota) as JumDiskon, sum(MSales.TotalDiscount) as DiskonBrg, Sum(MSales.UangMuka-MSales.Bank-MSales.Voucher) AS JumUangMuka, Sum(MSales.Bank) AS JumBank , Sum(MSales.Voucher) AS jumVoucher, " & _
'            "sum(MSales.JumDiscInternRp) as DiscBarangIntern,sum(MSales.DiscIntern) as DiscNotaIntern, " & _
'            "Sum(MSales.Pembulatan) as Bulat " & _
'            "From MSales " & _
'            "WHERE Day(Tanggal)=" & DayReset & " AND Month(Tanggal)=" & MonthReset & " AND Year(Tanggal)=" & YearReset & " and SubTotal<>0 AND Shift= " & NamaShiftReset & " " & _
'            "GROUP BY MSales.IDUser"
            'lbl(0) = Format(Now, "dd-MM-yyyy, hh:mm:ss")
            lbl(0) = Format(DayReset, "0#") & "-" & Format(MonthReset, "0#") & "-" & Format(YearReset, "####") & ", " & IIf(NamaShiftReset = 1, "PAGI", "SORE") 'Format(Now, "dd-MM-yyyy, hh:mm:ss")
            lbl(1) = namaMesinReset & " - " & namakasirReset
            lbl(2) = Format(rs!JumlahNota, "###,##0")
            'Diskon intern dan diskon extern di anggap dari supplier
            'lbl(3) = Format(rs!JumlahSubTotal + rs!DIskonBrg + rs!DiscBarangIntern + rs!DiscNotaIntern, "###,###,##0")
             ''Subtotal sudah dipotong disko intern DARI BRUTO DAN DITAMBAH DISKON EXTERN
             lbl(3) = Format(rs!JumlahSubTotal, "###,###,##0") '- rs!DiscBarangIntern - rs!DiscNotaIntern
             lbl(4) = Format(rs!JumDIskon + rs!DIskonBrg, "###,###,##0") & " + " & vbCrLf & _
                     Format(rs!DiscBarangIntern + rs!DiscNotaIntern, "###,###,##0") & " = " & _
                     Format(rs!JumDIskon + rs!DIskonBrg + rs!DiscBarangIntern + rs!DiscNotaIntern, "###,###,##0")
            lbl(5) = Format(rs!JumlahTotal + rs!JumDIskon + rs!DIskonBrg, "###,###,##0")
            lbl(6) = Format(rs!JumUangMuka, "###,###,##0")
            lbl(7) = Format(rs!Jumbank, "###,###,##0")
            lbl(8) = Format(rs!JumVoucher, "###,###,##0")
            lbl(9) = Format(rs!JumlahTotal - (rs!JumlahSubTotal - rs!JumDIskon) + rs!DIskonBrg + rs!DiscNotaIntern + rs!Bulat, "###,###,##0")
            TunaiPajak = ((PersenLap / 100) * rs!JumUangMuka) - (((PersenLap / 100) * rs!JumUangMuka) Mod 100)
            cVoucher = rs!JumVoucher
            cTunai = rs!JumUangMuka
            cJumlahNota = rs!JumlahNota
            cbank = rs!Jumbank
            cDiskonNota = rs!JumDIskon
            cDiskonBrg = rs!DIskonBrg
            cSubtotal = cTunai + cbank + cDiskonNota + cVoucher + cDiskonBrg + rs!Bulat
            Bulat = Format(rs!Bulat, "###,###,##0")
        End If
            If KodeKasir = KodeUserDua Then
              For i = 0 To 7
                 psn = psn & lblNama(i).Caption & Space(18 - Len(Trim(lblNama(i).Caption))) & ": " & lbl(i).Caption & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                  'If i = 4 Then
                  '    psn = psn & "Pembulatan" & Space(18 - Len("Pembulatan")) & ": " & Format(Bulat, "###,##0") & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                  'End If
              Next
            Else
              For i = 0 To 9
               psn = psn & lblNama(i).Caption & Space(18 - Len(Trim(lblNama(i).Caption))) & ": " & lbl(i).Caption & Chr(13) & Chr(10) & Chr(13) & Chr(10)
               If i = 4 Then
                      psn = psn & "Pembulatan" & Space(18 - Len("Pembulatan")) & ": " & Format(Bulat, "###,##0") & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                  End If
              Next
              End If
              Prin psn
              psn = ""
        rs.MoveNext
    Loop
End If
    psn = psn & "----------------------------------------" & Chr(13) & Chr(10)
    psn = psn & CekKassaShift1 & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10)
    
    Prin psn
'    PrinBigChar Space(15 - Min(15, Len(NamaToko))) & UCase(NamaToko)
'    papercut
'    Prin Judulstruk & Chr(13) & Chr(10) & "----------------------------------------"
    'MODIFIKASI 20-03-2007
    'Dijadikan 2 , server dan lokal
    '
    If KodeKasir = KodeUserDua Then
        NamaFileReset = Replace(getRegistry("Reset2", "Data") & "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb", "*", Text3.Text)
        NamaFileBackup = Replace(getRegistry("Backup2", "Data") & "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb", "*", Text3.Text)
    Else
        NamaFileReset = Replace(getRegistry("Reset1", "Data") & "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb", "*", Text3.Text)
    'Text2.Text = getRegistry("Reset2", "Data")
        NamaFileBackup = Replace(getRegistry("Backup", "Data") & "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb", "*", Text3.Text)
    End If
    
'    NamaFileReset = DIRPOS & "\Reset\K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb"
'    NamaFileBackup = DIRPOS & "\Backup\K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb"
    'RESET
    cekFile = "Reset"
    If Dir(NamaFileReset) <> "" Then
     Set dbsClear = OpenDatabase(NamaFileReset)
     dbsClear.Execute "Delete * FROM MSALES"
     dbsClear.Execute "Delete * FROM MSALESD"
     dbsClear.Close
    Else
     If KodeKasir = KodeUserDua Then
        Create_DatabaseResetDua NamaFileReset
     Else
        Create_DatabaseReset NamaFileReset
      End If
    End If
    cekFile = "Backup"
    'If KodeKasir <> KodeUserDua Then
        If Dir(NamaFileBackup) <> "" Then
         On Error Resume Next
         Kill (NamaFileBackup)
         Create_DatabaseBackup NamaFileBackup
         Set dbsClear = OpenDatabase(NamaFileBackup)
         dbsClear.Execute "Delete * FROM MSALES"
         dbsClear.Execute "Delete * FROM MSALESD"
         dbsClear.Close
         On Error GoTo 0
        Else
          Create_DatabaseBackup NamaFileBackup
        End If
               
        'BackUp
        dbs.Execute "INSERT INTO MSales ( Shift, IDUser,IDBank, Tanggal, NoID, Kode, TotalPajak, TotalDiscount, SubTotal, DiscIntern,DiscNota,JumDiscInternRp, HargaTotal, IDPayment, UangMuka, IsSend, Bank, ISPending,Voucher,IsUpload,IDMember,BarangKSB,SisaKSB,Pembulatan ) IN '" & NamaFileBackup & "' " & _
                  "SELECT MSales.Shift, MSales.IDUser,MSales.IDBank, MSales.Tanggal, MSales.NoID, MSales.Kode, MSales.TotalPajak, MSales.TotalDiscount, MSales.SubTotal,MSales.DiscIntern,MSales.JumDiscInternRp, MSales.DiscNota, MSales.HargaTotal, MSales.IDPayment, MSales.UangMuka, MSales.IsSend, MSales.Bank, MSales.ISPending , MSales.Voucher,MSales.IsUpload,MSales.IDMember,MSales.BarangKSB,MSales.SisaKSB,MSales.Pembulatan " & _
                  "From MSales " & _
                  "WHERE (((MSales.Shift)=" & NamaShiftReset & ") AND  (Day(MSales.Tanggal)=" & DayReset & " and Month(MSales.Tanggal)=" & MonthReset & " and Year(MSales.Tanggal)=" & YearReset & "))"
        
        dbs.Execute "INSERT INTO MSalesD IN '" & NamaFileBackup & "' " & _
                  "SELECT MSalesD.*,MsalesD.Qty*MSalesD.Harga as Jumlah " & _
                  "FROM MSales INNER JOIN MSalesD ON MSales.NoID = MSalesD.IDSales " & _
                  "WHERE (((MSales.Shift)=" & NamaShiftReset & ") AND (Day(MSales.Tanggal)=" & DayReset & " and Month(MSales.Tanggal)=" & MonthReset & " and Year(MSales.Tanggal)=" & YearReset & "))"
         dbs.Close
    'End If
      
  'RESET
    Set DbsReset = OpenDatabase(NamaFileReset)
    Set RsReset = DbsReset.OpenRecordset("MSales")
    Set dbs = OpenDatabase(DIRPOS & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
    Set rs = dbs.OpenRecordset("SELECT   sum(SubTotal) as SubTotala,sum(DiscNota) as DiscNotaa,sum(TotalDiscount) as DiscBrga, " & _
            "sum(HargaTotal) as HargaTotala, sum(UangMuka)-sum(Bank) as Tunaia, sum(Voucher) as Vouchera," & _
            "sum(Bank) as Banka,IDBank,IDUser,count(NoID) as JumlahNotaa , Max(HargaTotal+DiscNota+TotalDiscount) as NotaMax,Min(HargaTotal+DiscNota+TotalDiscount) as NotaMin " & _
            "From MSales " & _
              "WHERE (((MSales.Shift)=" & NamaShiftReset & ") AND  (Day(MSales.Tanggal)=" & DayReset & " and Month(MSales.Tanggal)=" & MonthReset & " and Year(MSales.Tanggal)=" & YearReset & ") AND (SubTotal<>0)) " & _
              "Group By day(Tanggal),month(Tanggal),year(Tanggal),IDBank,IDUser HAVING sum(HargaTotal)<>0")
     If rs.EOF And rs.BOF Then
    
     Else
        rs.MoveFirst
        NoID = 1
        Do While Not rs.EOF
            RsReset.AddNew
            RsReset!kode = "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & Format(NoID, "0##0")
            RsReset!NoID = NoID
            RsReset!TANGGAL = Mid(Text2.Text, 3, 2) & "/" & Mid(Text2.Text, 1, 2) & "/" & Mid(Text2.Text, 5, 4)
            RsReset!SubTotal = rs!HargaTotala + rs!DiscNotaa + rs!DiscBrga 'rs!SubTotala + rs!DiscBrga
            RsReset!DiscNota = 0 'rs!DiscNotaa diskon dipaksa 0
            RsReset!Hargatotal = rs!HargaTotala + rs!DiscNotaa + rs!DiscBrga
            RsReset!Tunai = rs!Tunaia + rs!DiscBrga + rs!DiscNotaa
            RsReset!Voucher = rs!Vouchera
            RsReset!Bank = rs!Banka
            RsReset!IDBank = rs!IDBank
            RsReset!IDUser = rs!IDUser
            RsReset!NamaUser = namakasirReset
            RsReset!Shift = NamaShiftReset
            RsReset!IdPengawas = IDUser
            RsReset!NamaPengawas = NamaKasir
            If KodeKasir = KodeUserDua Then
            MaxTotal = PersenLap * (rs!Tunaia + rs!DiscBrga + rs!DiscNotaa) / 100
            Else
                RsReset!JumlahNota = rs!JumlahNotaa
                RsReset!PajakPersen = PersenLap
                RsReset!NotaMin = rs!NotaMin
                RsReset!NotaMax = rs!NotaMax
                RsReset!TunaiPajak = PersenLap * (rs!Tunaia + rs!DiscBrga + rs!DiscNotaa) / 100
                MaxTotal = (rs!Tunaia + rs!DiscBrga + rs!DiscNotaa) '100 persen
            End If
            
            
            RsReset.Update
            RsReset.Bookmark = RsReset.LastModified
            
            If rs!Banka = 0 And KodeKasir = KodeUserDua Then 'Tunai dan kodeuserdua
                Dim RSp As Recordset
                Set RSp = dbs.OpenRecordset("SELECT " & NoID & " As IDSales, MSalesD.IDInventor, Sum(MSalesD.Qty) AS Qty, Sum(MSalesD.Qty*((MSalesD.HargaBruto-MSalesD.DiscInternRp)-iif((MSales.SubTotal+MSales.TotalDiscount+Msales.JumDiscInternRp)<>0,MSalesD.HargaBruto*(MSales.DiscIntern/(MSales.SubTotal+MSales.TotalDiscount+Msales.JumDiscInternRp)),0))) AS Jumlah, Sum((MSalesD.HargaPokok*MSalesD.Qty)) AS HPP, MSalesD.KodeInv, MSalesD.NamaInv, MSalesD.Satuan, MSalesD.Barcode, MSalesD.IDSatuan, MSalesD.Konversi " & _
                "FROM MSalesD INNER JOIN MSales ON MSalesD.IDSales = MSales.NoID " & _
                "Where day(MSales.Tanggal) = " & DayReset & " and Month(MSales.Tanggal) = " & MonthReset & " and Year(MSales.Tanggal) = " & YearReset & "  And MSales.ShifT = " & NamaShiftReset & " " & _
                "GROUP BY MSales.IDUser,Msales.IDBank,MSalesD.IDInventor, MSalesD.KodeInv, MSalesD.NamaInv, MSalesD.Satuan, MSalesD.Barcode, MSalesD.IDSatuan, MSalesD.Konversi " & _
                "HAVING (((Sum(MSalesD.Qty))<>0) AND MSales.IDUser=" & rs!IDUser & " AND MSales.IDBank=" & NullToNol(rs!IDBank) & ") " & _
                "ORDER BY Max(MSales.Tanggal) DESC ")
                If RSp.EOF And RSp.BOF Then
                Else
                    RSp.MoveFirst
                    curTotal = 0
                    Do While Not RSp.EOF
                        If MaxTotal <= curTotal Then Exit Do
                        curQty = ((RSp!QTY * PersenLap - RSp!QTY * PersenLap Mod 100) \ 100) + IIf(RSp!QTY * PersenLap Mod 100 > 1, 1, 0)
                        CurhargaSatuan = CLng(IIf(RSp!QTY <> 0, RSp!Jumlah / RSp!QTY, 0) * curQty)
                        If CurhargaSatuan Mod 50 > 0 Then CurhargaSatuan = CurhargaSatuan + (50 - (CurhargaSatuan Mod 50))
                        curTotal = curTotal + CurhargaSatuan
                          dbs.Execute "INSERT INTO MSalesD ( IDSales,IDInventor, Qty,HargaSat, Harga, HargaPokok," & _
                                      "KodeInv, NamaInv, Satuan, Barcode, IDSatuan, Konversi ) " & _
                                      "IN '" & NamaFileReset & "' " & _
                                      "VALUES(" & _
                                       NoID & "," & RSp!IdInventor & "," & curQty & "," & Replace(Format(IIf(curQty = 0, CurhargaSatuan, Abs(CurhargaSatuan / curQty)), "##0.00"), ",", ".") & "," & Replace(Format(CurhargaSatuan, "##0.00"), ",", ".") & ",0,'" & _
                                      Replace(RSp!KodeInv, "'", "''") & "','" & Replace(RSp!NamaInv, "'", "''") & "','" & Replace(RSp!Satuan, "'", "''") & "','" & RSp!Barcode & "'," & RSp!idSatuan & "," & RSp!Konversi & ")"
    
                        RSp.MoveNext
                    Loop
                    RsReset.Edit
                    
                    RsReset!SubTotal = curTotal
                    'RsReset!DiscNota = curTotal
                    RsReset!Hargatotal = curTotal
                    RsReset!Tunai = curTotal '- rs!DiscNotaa
                    RsReset.Update
                End If
            Else
              dbs.Execute "INSERT INTO MSalesD ( IDSales,IDInventor, Qty, Harga, HargaPokok, KodeInv, NamaInv, Satuan, Barcode, IDSatuan, Konversi ) IN '" & NamaFileReset & "'" & _
                "SELECT " & NoID & " As IDSales, MSalesD.IDInventor, Sum(MSalesD.Qty) AS Qty, Sum(MSalesD.Qty*((MSalesD.HargaBruto-MSalesD.DiscInternRp)-iif((MSales.SubTotal+MSales.TotalDiscount+Msales.JumDiscInternRp)<>0,MSalesD.HargaBruto*(MSales.DiscIntern/(MSales.SubTotal+MSales.TotalDiscount+Msales.JumDiscInternRp)),0))) AS Jumlah, Sum((MSalesD.HargaPokok*MSalesD.Qty)) AS HPP, MSalesD.KodeInv, MSalesD.NamaInv, MSalesD.Satuan, MSalesD.Barcode, MSalesD.IDSatuan, MSalesD.Konversi " & _
                "FROM MSalesD INNER JOIN MSales ON MSalesD.IDSales = MSales.NoID " & _
                "Where day(MSales.Tanggal) = " & DayReset & " and Month(MSales.Tanggal) = " & MonthReset & " and Year(MSales.Tanggal) = " & YearReset & "  And MSales.ShifT = " & NamaShiftReset & " " & _
                "GROUP BY MSales.IDUser,Msales.IDBank,MSalesD.IDInventor, MSalesD.KodeInv, MSalesD.NamaInv, MSalesD.Satuan, MSalesD.Barcode, MSalesD.IDSatuan, MSalesD.Konversi " & _
                "HAVING (((Sum(MSalesD.Qty))<>0) AND MSales.IDUser=" & rs!IDUser & " AND MSales.IDBank=" & NullToNol(rs!IDBank) & ")"
            DbsReset.Execute "UPDATE MSalesD SET MSalesD.HargaSat = [msalesd].[harga]/Abs([msalesd].[qty]) WHERE (((MSalesD.Qty)<>0)) AND IDSales=" & NoID
            End If
            DbsReset.Execute "Update MSALES  set Tanggal=#" & Mid(Text2.Text, 3, 2) & "/" & Mid(Text2.Text, 1, 2) & "/" & Mid(Text2.Text, 5, 4) & "#"
    '            If rs!Banka = 0 Then
    '                If KodeKasir = KodeUserDua Then
    '                  DbsReset.Execute "UPDATE MSalesD SET MSalesD.Qty=(MSalesD.Qty*" & PersenLap & ")\100,MSalesD.[harga]=((MSalesD.Qty*" & PersenLap & ")\100)*MSalesD.HargaSat WHERE  IDSales=" & NoId
    '                 End If
    '            End If
                NoID = NoID + 1
        rs.MoveNext
        Loop
     End If
 
    
    Set rs = Nothing
    dbs.Close
    Set RsReset = Nothing
    DbsReset.Close
    Set dbs = Nothing
    Set DbsReset = Nothing
    Exit Sub
pesan:
'   MsgBox "Ada kesalahan : " & Err.Description & vbCrLf & "Silahkan Tekan OK", vbCritical + vbOKOnly, "Pesan!"
   If Err.Number = 52 Then
    If cekFile = "Backup" Then
        NamaFileBackup = NMBACKUPGAGAL & "\" & "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb"
        Resume
    ElseIf cekFile = "Reset" Then
        NamaFileBackup = NMRESETGAGAL & "\" & "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb"
        Resume
    End If
   Else
    Resume Next
    End If
'    Unload Me
End Sub


Sub ResetNewsebAdadiscountbarang()
On Error GoTo pesan
    Dim dbs As Database
    Dim rs As Recordset
    Dim DbsReset As Database
    Dim RsReset As Recordset
    Dim psn As String
  Dim i As Integer
  Dim NoID As Integer
  Dim idkasirreset As Integer
  Dim NamaFileReset As String
  Dim NamaFileBackup As String
  Dim TunaiPajak As Long
  Dim cTunai As Long
  Dim cJumlahNota As Long
  Dim cbank As Double
  Dim cSubtotal As Double
  Dim cUangMuka As Double
  Dim cTotal As Double
  Dim cDiskonNota As Double
  Dim cVoucher As Double
  Dim dbsClear As Database
  Dim strqry As String
  Dim MaxTotal As Double
  Dim curTotal As Double
  Dim curQty As Long
  Dim CurhargaSatuan As Double
  Dim cekFile As String
  Dim JmlSalah As Integer
    lblNama(9).Caption = "Biaya Credit Card:"
    Set dbs = OpenDatabase(DIRPOS & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
    dbs.Execute "Update Msales Set IDbank =2 where isnull(idbank)"
    strqry = "SELECT MSales.IDUser,Count(MSales.Kode) AS JumlahNota, Sum(MSales.SubTotal) AS JumlahSubTotal, Sum(MSales.HargaTotal) AS JumlahTotal, sum(MSales.DiscNota) as JumDiskon, Sum(MSales.UangMuka-MSales.Bank-MSales.Voucher) AS JumUangMuka, Sum(MSales.Bank) AS JumBank , Sum(MSales.Voucher) AS jumVoucher " & _
            "From MSales " & _
            "WHERE Day(Tanggal)=" & DayReset & " AND Month(Tanggal)=" & MonthReset & " AND Year(Tanggal)=" & YearReset & " and SubTotal<>0 AND Shift= " & NamaShiftReset & " " & _
            "GROUP BY MSales.IDUser"
    Set rs = dbs.OpenRecordset(strqry)
    If rs.EOF And rs.BOF Then
      Exit Sub
    Else
    rs.MoveFirst
       PrinBigChar Chr(13) & Chr(10) & Chr(13) & "   KASSA : " & Str(namaMesinReset)
       psn = Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "----------------RESET----------------" & Chr(13) & Chr(10) & Chr(13) & Chr(10)
       Prin psn
       psn = ""
    Do While Not rs.EOF
      idkasirreset = rs!IDUser
        cariKodeNama KodeKasirReset, namakasirReset, rs!IDUser
        If KodeKasir = KodeUserDua Then
            TunaiPajak = ((PersenLap / 100) * rs!JumUangMuka) - (((PersenLap / 100) * rs!JumUangMuka) Mod 100)
            cTunai = rs!JumUangMuka
            cJumlahNota = rs!JumlahNota
            cbank = rs!Jumbank
            'cVoucher = rs!JumVoucher
            cVoucher = 0
            cDiskonNota = rs!JumDIskon
            cSubtotal = cTunai + cbank + cDiskonNota + cVoucher
            lbl(0) = Format(DayReset, "0#") & "-" & Format(MonthReset, "0#") & "-" & Format(YearReset, "####") & ", " & IIf(NamaShiftReset = 1, "PAGI", "SORE") 'Format(Now, "dd-MM-yyyy, hh:mm:ss")
            lbl(1) = namaMesinReset & " - " & namakasirReset
            lbl(2) = Format(rs!JumlahNota, "###,##0")
            lbl(3) = Format(rs!Jumbank + TunaiPajak + cVoucher + rs!JumDIskon, "###,###,##0")
            lbl(4) = Format(rs!JumDIskon, "###,###,##0")
            lbl(5) = Format(rs!Jumbank + TunaiPajak, "###,###,##0")
            lbl(6) = Format(TunaiPajak, "###,###,##0")
            lbl(7) = Format(rs!Jumbank, "###,###,##0")
            lbl(8) = Format(cVoucher, "###,###,##0")
            'lbl(9) = (rs!Jumbank + TunaiPajak)
        Else
            'lbl(0) = Format(Now, "dd-MM-yyyy, hh:mm:ss")
            lbl(0) = Format(DayReset, "0#") & "-" & Format(MonthReset, "0#") & "-" & Format(YearReset, "####") & ", " & IIf(NamaShiftReset = 1, "PAGI", "SORE") 'Format(Now, "dd-MM-yyyy, hh:mm:ss")
            lbl(1) = namaMesinReset & " - " & namakasirReset
            lbl(2) = Format(rs!JumlahNota, "###,##0")
            lbl(3) = Format(rs!JumlahSubTotal, "###,###,##0")
            lbl(4) = Format(rs!JumDIskon, "###,###,##0")
            lbl(5) = Format(rs!JumlahTotal, "###,###,##0")
            lbl(6) = Format(rs!JumUangMuka, "###,###,##0")
            lbl(7) = Format(rs!Jumbank, "###,###,##0")
            lbl(8) = Format(rs!JumVoucher, "###,###,##0")
            lbl(9) = Format(rs!JumlahTotal - rs!JumlahSubTotal, "###,###,##0")
            TunaiPajak = ((PersenLap / 100) * rs!JumUangMuka) - (((PersenLap / 100) * rs!JumUangMuka) Mod 100)
            cVoucher = rs!JumVoucher
            cTunai = rs!JumUangMuka
            cJumlahNota = rs!JumlahNota
            cbank = rs!Jumbank
            cDiskonNota = rs!JumDIskon
            cSubtotal = cTunai + cbank + cDiskonNota + cVoucher
        End If
            If KodeKasir = KodeUserDua Then
            For i = 0 To 7
               psn = psn & lblNama(i).Caption & Space(18 - Len(Trim(lblNama(i).Caption))) & ": " & lbl(i).Caption & Chr(13) & Chr(10) & Chr(13) & Chr(10)
              Next
            Else
              For i = 0 To 9
               psn = psn & lblNama(i).Caption & Space(18 - Len(Trim(lblNama(i).Caption))) & ": " & lbl(i).Caption & Chr(13) & Chr(10) & Chr(13) & Chr(10)
              Next
              End If
              Prin psn
              psn = ""
        rs.MoveNext
    Loop
End If
    psn = psn & "----------------------------------------" & Chr(13) & Chr(10)
    psn = psn & CekKassaShift1 & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10)
    
    Prin psn
    PrinBigChar Space(15 - Min(15, Len(NamaToko))) & UCase(NamaToko)
    papercut
    Prin Judulstruk & Chr(13) & Chr(10) & "----------------------------------------"
    'MODIFIKASI 20-03-2007
    'Dijadikan 2 , server dan lokal
    '
    If KodeKasir = KodeUserDua Then
        NamaFileReset = Replace(getRegistry("Reset2", "Data") & "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb", "*", Text3.Text)
        NamaFileBackup = Replace(getRegistry("Backup2", "Data") & "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb", "*", Text3.Text)
    Else
        NamaFileReset = Replace(getRegistry("Reset1", "Data") & "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb", "*", Text3.Text)
    'Text2.Text = getRegistry("Reset2", "Data")
        NamaFileBackup = Replace(getRegistry("Backup", "Data") & "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb", "*", Text3.Text)
    End If
    
'    NamaFileReset = DIRPOS & "\Reset\K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb"
'    NamaFileBackup = DIRPOS & "\Backup\K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb"
    'RESET
    cekFile = "Reset"
    If Dir(NamaFileReset) <> "" Then
     Set dbsClear = OpenDatabase(NamaFileReset)
     dbsClear.Execute "Delete * FROM MSALES"
     dbsClear.Execute "Delete * FROM MSALESD"
     dbsClear.Close
    Else
     If KodeKasir = KodeUserDua Then
        Create_DatabaseResetDua NamaFileReset
     Else
        Create_DatabaseReset NamaFileReset
      End If
    End If
    cekFile = "Backup"
    'If KodeKasir <> KodeUserDua Then
        If Dir(NamaFileBackup) <> "" Then
         Set dbsClear = OpenDatabase(NamaFileBackup)
         dbsClear.Execute "Delete * FROM MSALES"
         dbsClear.Execute "Delete * FROM MSALESD"
         dbsClear.Close
        Else
          Create_DatabaseBackup NamaFileBackup
        End If
               
        'BackUp
        dbs.Execute "INSERT INTO MSales ( Shift, IDUser,IDBank, Tanggal, NoID, Kode, TotalPajak, TotalDiscount, SubTotal, DiscNota, HargaTotal, IDPayment, UangMuka, IsSend, Bank, ISPending,Voucher ) IN '" & NamaFileBackup & "' " & _
                  "SELECT MSales.Shift, MSales.IDUser,MSales.IDBank, MSales.Tanggal, MSales.NoID, MSales.Kode, MSales.TotalPajak, MSales.TotalDiscount, MSales.SubTotal, MSales.DiscNota, MSales.HargaTotal, MSales.IDPayment, MSales.UangMuka, MSales.IsSend, MSales.Bank, MSales.ISPending , MSales.Voucher " & _
                  "From MSales " & _
                  "WHERE (((MSales.Shift)=" & NamaShiftReset & ") AND  (Day(MSales.Tanggal)=" & DayReset & " and Month(MSales.Tanggal)=" & MonthReset & " and Year(MSales.Tanggal)=" & YearReset & "))"
        
        dbs.Execute "INSERT INTO MSalesD IN '" & NamaFileBackup & "' " & _
                  "SELECT MSalesD.*,MsalesD.Qty*MSalesD.Harga as Jumlah " & _
                  "FROM MSales INNER JOIN MSalesD ON MSales.NoID = MSalesD.IDSales " & _
                  "WHERE (((MSales.Shift)=" & NamaShiftReset & ") AND (Day(MSales.Tanggal)=" & DayReset & " and Month(MSales.Tanggal)=" & MonthReset & " and Year(MSales.Tanggal)=" & YearReset & "))"
         dbs.Close
    'End If
      
  'RESET
    Set DbsReset = OpenDatabase(NamaFileReset)
    Set RsReset = DbsReset.OpenRecordset("MSales")
    Set dbs = OpenDatabase(DIRPOS & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
    Set rs = dbs.OpenRecordset("SELECT   sum(SubTotal) as SubTotala,sum(DiscNota) as DiscNotaa, " & _
            "sum(HargaTotal) as HargaTotala, sum(UangMuka)-sum(Bank) as Tunaia, sum(Voucher) as Vouchera," & _
            "sum(Bank) as Banka,IDBank,IDUser,count(NoID) as JumlahNotaa , Max(SubTotal) as NotaMax,Min(SubTotal) as NotaMin " & _
            "From MSales " & _
              "WHERE (((MSales.Shift)=" & NamaShiftReset & ") AND  (Day(MSales.Tanggal)=" & DayReset & " and Month(MSales.Tanggal)=" & MonthReset & " and Year(MSales.Tanggal)=" & YearReset & ") AND (SubTotal<>0)) " & _
              "Group By day(Tanggal),month(Tanggal),year(Tanggal),IDBank,IDUser HAVING sum(HargaTotal)<>0")
     If rs.EOF And rs.BOF Then
    
     Else
        rs.MoveFirst
        NoID = 1
        Do While Not rs.EOF
            RsReset.AddNew
            RsReset!kode = "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & Format(NoID, "0##0")
            RsReset!NoID = NoID
            RsReset!TANGGAL = Mid(Text2.Text, 3, 2) & "/" & Mid(Text2.Text, 1, 2) & "/" & Mid(Text2.Text, 5, 4)
            RsReset!SubTotal = rs!SubTotala
            RsReset!DiscNota = rs!DiscNotaa
            RsReset!Hargatotal = rs!HargaTotala
            RsReset!Tunai = rs!Tunaia
            RsReset!Voucher = rs!Vouchera
            RsReset!Bank = rs!Banka
            RsReset!IDBank = rs!IDBank
            RsReset!IDUser = rs!IDUser
            RsReset!NamaUser = namakasirReset
            RsReset!Shift = NamaShiftReset
            RsReset!IdPengawas = IDUser
            RsReset!NamaPengawas = NamaKasir
            If KodeKasir = KodeUserDua Then
            MaxTotal = PersenLap * rs!Tunaia / 100
            Else
                RsReset!JumlahNota = rs!JumlahNotaa
                RsReset!PajakPersen = PersenLap
                RsReset!NotaMin = rs!NotaMin
                RsReset!NotaMax = rs!NotaMax
                RsReset!TunaiPajak = PersenLap * rs!Tunaia / 100
                MaxTotal = rs!Tunaia '100 persen
            End If
            
            
            RsReset.Update
            RsReset.Bookmark = RsReset.LastModified
            
            If rs!Banka = 0 And KodeKasir = KodeUserDua Then 'Tunai dan kodeuserdua
                Dim RSp As Recordset
                Set RSp = dbs.OpenRecordset("SELECT " & NoID & " As IDSales, MSalesD.IDInventor, Sum(MSalesD.Qty) AS Qty, Sum((MSalesD.Harga*MSalesD.Qty)) AS Jumlah, Sum((MSalesD.HargaPokok*MSalesD.Qty)) AS HPP, MSalesD.KodeInv, MSalesD.NamaInv, MSalesD.Satuan, MSalesD.Barcode, MSalesD.IDSatuan, MSalesD.Konversi " & _
                "FROM MSalesD INNER JOIN MSales ON MSalesD.IDSales = MSales.NoID " & _
                "Where day(MSales.Tanggal) = " & DayReset & " and Month(MSales.Tanggal) = " & MonthReset & " and Year(MSales.Tanggal) = " & YearReset & "  And MSales.ShifT = " & NamaShiftReset & " " & _
                "GROUP BY MSales.IDUser,Msales.IDBank,MSalesD.IDInventor, MSalesD.KodeInv, MSalesD.NamaInv, MSalesD.Satuan, MSalesD.Barcode, MSalesD.IDSatuan, MSalesD.Konversi " & _
                "HAVING (((Sum(MSalesD.Qty))<>0) AND MSales.IDUser=" & rs!IDUser & " AND MSales.IDBank=" & NullToNol(rs!IDBank) & ") " & _
                "ORDER BY Max(MSales.Tanggal) DESC ")
                RSp.MoveFirst
                curTotal = 0
                Do While Not RSp.EOF
                    If MaxTotal <= curTotal Then Exit Do
                    curQty = ((RSp!QTY * PersenLap - RSp!QTY * PersenLap Mod 100) \ 100) + IIf(RSp!QTY * PersenLap Mod 100 > 1, 1, 0)
                    CurhargaSatuan = CLng(IIf(RSp!QTY <> 0, RSp!Jumlah / RSp!QTY, 0) * curQty)
                    If CurhargaSatuan Mod 50 > 0 Then CurhargaSatuan = CurhargaSatuan + (50 - (CurhargaSatuan Mod 50))
                    curTotal = curTotal + CurhargaSatuan
                      dbs.Execute "INSERT INTO MSalesD ( IDSales,IDInventor, Qty,HargaSat, Harga, HargaPokok," & _
                                  "KodeInv, NamaInv, Satuan, Barcode, IDSatuan, Konversi ) " & _
                                  "IN '" & NamaFileReset & "' " & _
                                  "VALUES(" & _
                                   NoID & "," & RSp!IdInventor & "," & curQty & "," & Replace(Format(IIf(curQty = 0, CurhargaSatuan, Abs(CurhargaSatuan / curQty)), "##0.00"), ",", ".") & "," & Replace(Format(CurhargaSatuan, "##0.00"), ",", ".") & ",0,'" & _
                                  Replace(RSp!KodeInv, "'", "''") & "','" & Replace(RSp!NamaInv, "'", "''") & "','" & Replace(RSp!Satuan, "'", "''") & "','" & RSp!Barcode & "'," & RSp!idSatuan & "," & RSp!Konversi & ")"

                    RSp.MoveNext
                Loop
                RsReset.Edit
                
                RsReset!SubTotal = curTotal
                'RsReset!DiscNota = curTotal
                RsReset!Hargatotal = curTotal
                RsReset!Tunai = curTotal - rs!DiscNotaa
                RsReset.Update
            Else
              dbs.Execute "INSERT INTO MSalesD ( IDSales,IDInventor, Qty, Harga, HargaPokok, KodeInv, NamaInv, Satuan, Barcode, IDSatuan, Konversi ) IN '" & NamaFileReset & "'" & _
                "SELECT " & NoID & " As IDSales, MSalesD.IDInventor, Sum(MSalesD.Qty) AS Qty, Sum((MSalesD.Harga*MSalesD.Qty)) AS Jumlah, Sum((MSalesD.HargaPokok*MSalesD.Qty)) AS HPP, MSalesD.KodeInv, MSalesD.NamaInv, MSalesD.Satuan, MSalesD.Barcode, MSalesD.IDSatuan, MSalesD.Konversi " & _
                "FROM MSalesD INNER JOIN MSales ON MSalesD.IDSales = MSales.NoID " & _
                "Where day(MSales.Tanggal) = " & DayReset & " and Month(MSales.Tanggal) = " & MonthReset & " and Year(MSales.Tanggal) = " & YearReset & "  And MSales.ShifT = " & NamaShiftReset & " " & _
                "GROUP BY MSales.IDUser,Msales.IDBank,MSalesD.IDInventor, MSalesD.KodeInv, MSalesD.NamaInv, MSalesD.Satuan, MSalesD.Barcode, MSalesD.IDSatuan, MSalesD.Konversi " & _
                "HAVING (((Sum(MSalesD.Qty))<>0) AND MSales.IDUser=" & rs!IDUser & " AND MSales.IDBank=" & NullToNol(rs!IDBank) & ")"
            DbsReset.Execute "UPDATE MSalesD SET MSalesD.HargaSat = [msalesd].[harga]/Abs([msalesd].[qty]) WHERE (((MSalesD.Qty)<>0)) AND IDSales=" & NoID
            End If
            DbsReset.Execute "Update MSALES  set Tanggal=#" & Mid(Text2.Text, 3, 2) & "/" & Mid(Text2.Text, 1, 2) & "/" & Mid(Text2.Text, 5, 4) & "#"
    '            If rs!Banka = 0 Then
    '                If KodeKasir = KodeUserDua Then
    '                  DbsReset.Execute "UPDATE MSalesD SET MSalesD.Qty=(MSalesD.Qty*" & PersenLap & ")\100,MSalesD.[harga]=((MSalesD.Qty*" & PersenLap & ")\100)*MSalesD.HargaSat WHERE  IDSales=" & NoId
    '                 End If
    '            End If
                NoID = NoID + 1
        rs.MoveNext
        Loop
     End If
 
    
    Set rs = Nothing
    dbs.Close
    Set RsReset = Nothing
    DbsReset.Close
    Set dbs = Nothing
    Set DbsReset = Nothing
    Exit Sub
pesan:
'   MsgBox "Ada kesalahan : " & Err.Description & vbCrLf & "Silahkan Tekan OK", vbCritical + vbOKOnly, "Pesan!"
   If Err.Number = 52 Then
    If cekFile = "Backup" Then
        NamaFileBackup = NMBACKUPGAGAL & "\" & "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb"
        Resume
    ElseIf cekFile = "Reset" Then
        NamaFileBackup = NMRESETGAGAL & "\" & "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb"
        Resume
    End If
   Else
    Resume Next
    End If
'    Unload Me
End Sub

Sub ResetOLDNew()
    Dim dbs As Database
    Dim rs As Recordset
    Dim DbsReset As Database
    Dim RsReset As Recordset
    Dim psn As String
  Dim i As Integer
  Dim NoID As Integer
  Dim idkasirreset As Integer
  Dim NamaFileReset As String
  Dim NamaFileBackup As String
  Dim TunaiPajak As Long
  Dim cTunai As Long
  Dim cJumlahNota As Long
  Dim cbank As Double
  Dim cSubtotal As Double
  Dim cUangMuka As Double
  Dim cTotal As Double
  Dim cDiskonNota As Double
  Dim cVoucher As Double
  Dim dbsClear As Database
  Dim strqry As String
  Dim MaxTotal As Double
  Dim curTotal As Double
  Dim curQty As Long
  Dim CurhargaSatuan As Double
    lblNama(9).Caption = "Biaya Credit Card:"
    Set dbs = OpenDatabase(DIRPOS & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
    dbs.Execute "Update Msales Set IDbank =2 where isnull(idbank)"
    strqry = "SELECT MSales.IDUser,Count(MSales.Kode) AS JumlahNota, Sum(MSales.SubTotal) AS JumlahSubTotal, Sum(MSales.HargaTotal) AS JumlahTotal, sum(MSales.DiscNota) as JumDiskon, Sum(MSales.UangMuka-MSales.Bank-MSales.Voucher) AS JumUangMuka, Sum(MSales.Bank) AS JumBank , Sum(MSales.Voucher) AS jumVoucher " & _
            "From MSales " & _
            "WHERE Day(Tanggal)=" & DayReset & " AND Month(Tanggal)=" & MonthReset & " AND Year(Tanggal)=" & YearReset & " and SubTotal<>0 AND Shift= " & NamaShiftReset & " " & _
            "GROUP BY MSales.IDUser"
    Set rs = dbs.OpenRecordset(strqry)
    If rs.EOF And rs.BOF Then
      Exit Sub
    Else
    rs.MoveFirst
       PrinBigChar Chr(13) & Chr(10) & Chr(13) & "   KASSA : " & Str(namaMesinReset)
       psn = Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "----------------RESET----------------" & Chr(13) & Chr(10) & Chr(13) & Chr(10)
       Prin psn
       psn = ""
    Do While Not rs.EOF
      idkasirreset = rs!IDUser
        cariKodeNama KodeKasirReset, namakasirReset, rs!IDUser
        If KodeKasir = KodeUserDua Then
            TunaiPajak = ((PersenLap / 100) * rs!JumUangMuka) - (((PersenLap / 100) * rs!JumUangMuka) Mod 100)
            cTunai = rs!JumUangMuka
            cJumlahNota = rs!JumlahNota
            cbank = rs!Jumbank
            'cVoucher = rs!JumVoucher
            cVoucher = 0
            cDiskonNota = rs!JumDIskon
            cSubtotal = cTunai + cbank + cDiskonNota + cVoucher
            lbl(0) = Format(DayReset, "0#") & "-" & Format(MonthReset, "0#") & "-" & Format(YearReset, "####") & ", " & IIf(NamaShiftReset = 1, "PAGI", "SORE") 'Format(Now, "dd-MM-yyyy, hh:mm:ss")
            lbl(1) = namaMesinReset & " - " & namakasirReset
            lbl(2) = Format(rs!JumlahNota, "###,##0")
            lbl(3) = Format(rs!Jumbank + TunaiPajak + cVoucher + rs!JumDIskon, "###,###,##0")
            lbl(4) = Format(rs!JumDIskon, "###,###,##0")
            lbl(5) = Format(rs!Jumbank + TunaiPajak, "###,###,##0")
            lbl(6) = Format(TunaiPajak, "###,###,##0")
            lbl(7) = Format(rs!Jumbank, "###,###,##0")
            lbl(8) = Format(cVoucher, "###,###,##0")
            'lbl(9) = (rs!Jumbank + TunaiPajak)
        Else
            'lbl(0) = Format(Now, "dd-MM-yyyy, hh:mm:ss")
            lbl(0) = Format(DayReset, "0#") & "-" & Format(MonthReset, "0#") & "-" & Format(YearReset, "####") & ", " & IIf(NamaShiftReset = 1, "PAGI", "SORE") 'Format(Now, "dd-MM-yyyy, hh:mm:ss")
            lbl(1) = namaMesinReset & " - " & namakasirReset
            lbl(2) = Format(rs!JumlahNota, "###,##0")
            lbl(3) = Format(rs!JumlahSubTotal, "###,###,##0")
            lbl(4) = Format(rs!JumDIskon, "###,###,##0")
            lbl(5) = Format(rs!JumlahTotal, "###,###,##0")
            lbl(6) = Format(rs!JumUangMuka, "###,###,##0")
            lbl(7) = Format(rs!Jumbank, "###,###,##0")
            lbl(8) = Format(rs!JumVoucher, "###,###,##0")
            lbl(9) = Format(rs!JumlahTotal - rs!JumlahSubTotal, "###,###,##0")
            TunaiPajak = ((PersenLap / 100) * rs!JumUangMuka) - (((PersenLap / 100) * rs!JumUangMuka) Mod 100)
            cVoucher = rs!JumVoucher
            cTunai = rs!JumUangMuka
            cJumlahNota = rs!JumlahNota
            cbank = rs!Jumbank
            cDiskonNota = rs!JumDIskon
            cSubtotal = cTunai + cbank + cDiskonNota + cVoucher
        End If
            If KodeKasir = KodeUserDua Then
            For i = 0 To 7
               psn = psn & lblNama(i).Caption & Space(18 - Len(Trim(lblNama(i).Caption))) & ": " & lbl(i).Caption & Chr(13) & Chr(10) & Chr(13) & Chr(10)
              Next
            Else
              For i = 0 To 9
               psn = psn & lblNama(i).Caption & Space(18 - Len(Trim(lblNama(i).Caption))) & ": " & lbl(i).Caption & Chr(13) & Chr(10) & Chr(13) & Chr(10)
              Next
              End If
              Prin psn
              psn = ""
        rs.MoveNext
    Loop
End If
    psn = psn & "----------------------------------------" & Chr(13) & Chr(10)
    psn = psn & CekKassaShift1 & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10)
    
    Prin psn
    PrinBigChar Space(15 - Min(15, Len(NamaToko))) & UCase(NamaToko)
    papercut
    Prin Judulstruk & Chr(13) & Chr(10) & "----------------------------------------"
    'MODIFIKASI 20-03-2007
    'Dijadikan 2 , server dan lokal
    '
    If KodeKasir = KodeUserDua Then
        NamaFileReset = getRegistry("Reset2", "Data") & "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb"
        NamaFileBackup = getRegistry("Backup2", "Data") & "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb"
    Else
        NamaFileReset = getRegistry("Reset1", "Data") & "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb"
    'Text2.Text = getRegistry("Reset2", "Data")
        NamaFileBackup = getRegistry("Backup", "Data") & "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb"
    End If
    
'    NamaFileReset = DIRPOS & "\Reset\K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb"
'    NamaFileBackup = DIRPOS & "\Backup\K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & ".mdb"
    'RESET
    If Dir(NamaFileReset) <> "" Then
     Set dbsClear = OpenDatabase(NamaFileReset)
     dbsClear.Execute "Delete * FROM MSALES"
     dbsClear.Execute "Delete * FROM MSALESD"
     dbsClear.Close
    Else
     If KodeKasir = KodeUserDua Then
        Create_DatabaseResetDua NamaFileReset
     Else
        Create_DatabaseReset NamaFileReset
      End If
    End If
    
    'If KodeKasir <> KodeUserDua Then
        If Dir(NamaFileBackup) <> "" Then
         Set dbsClear = OpenDatabase(NamaFileBackup)
         dbsClear.Execute "Delete * FROM MSALES"
         dbsClear.Execute "Delete * FROM MSALESD"
         dbsClear.Close
        Else
          Create_DatabaseBackup NamaFileBackup
        End If
               
        'BackUp
        dbs.Execute "INSERT INTO MSales ( Shift, IDUser,IDBank, Tanggal, NoID, Kode, TotalPajak, TotalDiscount, SubTotal, DiscNota, HargaTotal, IDPayment, UangMuka, IsSend, Bank, ISPending,Voucher ) IN '" & NamaFileBackup & "' " & _
                  "SELECT MSales.Shift, MSales.IDUser,MSales.IDBank, MSales.Tanggal, MSales.NoID, MSales.Kode, MSales.TotalPajak, MSales.TotalDiscount, MSales.SubTotal, MSales.DiscNota, MSales.HargaTotal, MSales.IDPayment, MSales.UangMuka, MSales.IsSend, MSales.Bank, MSales.ISPending , MSales.Voucher " & _
                  "From MSales " & _
                  "WHERE (((MSales.Shift)=" & NamaShiftReset & ") AND  (Day(MSales.Tanggal)=" & DayReset & " and Month(MSales.Tanggal)=" & MonthReset & " and Year(MSales.Tanggal)=" & YearReset & "))"
        
        dbs.Execute "INSERT INTO MSalesD IN '" & NamaFileBackup & "' " & _
                  "SELECT MSalesD.* " & _
                  "FROM MSales INNER JOIN MSalesD ON MSales.NoID = MSalesD.IDSales " & _
                  "WHERE (((MSales.Shift)=" & NamaShiftReset & ") AND (Day(MSales.Tanggal)=" & DayReset & " and Month(MSales.Tanggal)=" & MonthReset & " and Year(MSales.Tanggal)=" & YearReset & "))"
         dbs.Close
    'End If
      
  'RESET
    Set DbsReset = OpenDatabase(NamaFileReset)
    Set RsReset = DbsReset.OpenRecordset("MSales")
    Set dbs = OpenDatabase(DIRPOS & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
    Set rs = dbs.OpenRecordset("SELECT   sum(SubTotal) as SubTotala,sum(DiscNota) as DiscNotaa, " & _
            "sum(HargaTotal) as HargaTotala, sum(UangMuka)-sum(Bank) as Tunaia, sum(Voucher) as Vouchera," & _
            "sum(Bank) as Banka,IDBank,IDUser,count(NoID) as JumlahNotaa , Max(SubTotal) as NotaMax,Min(SubTotal) as NotaMin " & _
            "From MSales " & _
              "WHERE (((MSales.Shift)=" & NamaShiftReset & ") AND  (Day(MSales.Tanggal)=" & DayReset & " and Month(MSales.Tanggal)=" & MonthReset & " and Year(MSales.Tanggal)=" & YearReset & ") AND (SubTotal<>0)) " & _
              "Group By day(Tanggal),month(Tanggal),year(Tanggal),IDBank,IDUser HAVING sum(HargaTotal)<>0")
     If rs.EOF And rs.BOF Then
    
     Else
        rs.MoveFirst
        NoID = 1
        Do While Not rs.EOF
            RsReset.AddNew
            RsReset!kode = "K" & Trim(namaMesinReset) & Mid(Text2.Text, 5, 4) & Mid(Text2.Text, 3, 2) & Mid(Text2.Text, 1, 2) & Format(NamaShiftReset, "0#") & Format(NoID, "0##0")
            RsReset!NoID = NoID
            RsReset!TANGGAL = Mid(Text2.Text, 3, 2) & "/" & Mid(Text2.Text, 1, 2) & "/" & Mid(Text2.Text, 5, 4)
            RsReset!SubTotal = rs!SubTotala
            RsReset!DiscNota = rs!DiscNotaa
            RsReset!Hargatotal = rs!HargaTotala
            RsReset!Tunai = rs!Tunaia
            RsReset!Voucher = rs!Vouchera
            RsReset!Bank = rs!Banka
            RsReset!IDBank = rs!IDBank
            RsReset!IDUser = rs!IDUser
            RsReset!NamaUser = namakasirReset
            RsReset!Shift = NamaShiftReset
            RsReset!IdPengawas = IDUser
            RsReset!NamaPengawas = NamaKasir
            If KodeKasir = KodeUserDua Then
            Else
            RsReset!JumlahNota = rs!JumlahNotaa
            RsReset!PajakPersen = PersenLap
            RsReset!NotaMin = rs!NotaMin
            RsReset!NotaMax = rs!NotaMax
            RsReset!TunaiPajak = PersenLap * rs!Tunaia / 100
           
            End If
            MaxTotal = PersenLap * rs!Tunaia / 100
            RsReset.Update
            RsReset.Bookmark = RsReset.LastModified
            
            If rs!Banka = 0 Then 'Tunai
                Dim RSp As Recordset
                Set RSp = dbs.OpenRecordset("SELECT " & NoID & " As IDSales, MSalesD.IDInventor, Sum(MSalesD.Qty) AS Qty, Sum((MSalesD.Harga*MSalesD.Qty)) AS Jumlah, Sum((MSalesD.HargaPokok*MSalesD.Qty)) AS HPP, MSalesD.KodeInv, MSalesD.NamaInv, MSalesD.Satuan, MSalesD.Barcode, MSalesD.IDSatuan, MSalesD.Konversi " & _
                "FROM MSalesD INNER JOIN MSales ON MSalesD.IDSales = MSales.NoID " & _
                "Where day(MSales.Tanggal) = " & DayReset & " and Month(MSales.Tanggal) = " & MonthReset & " and Year(MSales.Tanggal) = " & YearReset & "  And MSales.ShifT = " & NamaShiftReset & " " & _
                "GROUP BY MSales.IDUser,Msales.IDBank,MSalesD.IDInventor, MSalesD.KodeInv, MSalesD.NamaInv, MSalesD.Satuan, MSalesD.Barcode, MSalesD.IDSatuan, MSalesD.Konversi " & _
                "HAVING (((Sum(MSalesD.Qty))<>0) AND MSales.IDUser=" & rs!IDUser & " AND MSales.IDBank=" & NullToNol(rs!IDBank) & ") " & _
                "ORDER BY Max(MSales.Tanggal) DESC ")
                RSp.MoveFirst
                curTotal = 0
                Do While Not RSp.EOF
                    If MaxTotal <= curTotal Then Exit Do
                    curQty = ((RSp!QTY * PersenLap - RSp!QTY * PersenLap Mod 100) \ 100) + IIf(RSp!QTY * PersenLap Mod 100 > 1, 1, 0)
                    CurhargaSatuan = CLng(IIf(RSp!QTY <> 0, RSp!Jumlah / RSp!QTY, 0) * curQty)
                    If CurhargaSatuan Mod 50 > 0 Then CurhargaSatuan = CurhargaSatuan + (50 - (CurhargaSatuan Mod 50))
                    curTotal = curTotal + CurhargaSatuan
                      dbs.Execute "INSERT INTO MSalesD ( IDSales,IDInventor, Qty, Harga, HargaPokok," & _
                                  "KodeInv, NamaInv, Satuan, Barcode, IDSatuan, Konversi ) " & _
                                  "IN '" & NamaFileReset & "' " & _
                                  "VALUES(" & _
                                   NoID & "," & RSp!IdInventor & "," & curQty & "," & Replace(Format(CurhargaSatuan, "##0.00"), ",", ".") & ",0,'" & _
                                  Replace(RSp!KodeInv, "'", "''") & "','" & Replace(RSp!NamaInv, "'", "''") & "','" & Replace(RSp!Satuan, "'", "''") & "','" & RSp!Barcode & "'," & RSp!idSatuan & "," & RSp!Konversi & ")"

                    RSp.MoveNext
                Loop
                RsReset.Edit
                
                RsReset!SubTotal = curTotal
                'RsReset!DiscNota = curTotal
                RsReset!Hargatotal = curTotal
                RsReset!Tunai = curTotal - rs!DiscNotaa
                RsReset.Update
            Else
              dbs.Execute "INSERT INTO MSalesD ( IDSales,IDInventor, Qty, Harga, HargaPokok, KodeInv, NamaInv, Satuan, Barcode, IDSatuan, Konversi ) IN '" & NamaFileReset & "'" & _
                "SELECT " & NoID & " As IDSales, MSalesD.IDInventor, Sum(MSalesD.Qty) AS Qty, Sum((MSalesD.Harga*MSalesD.Qty)) AS Jumlah, Sum((MSalesD.HargaPokok*MSalesD.Qty)) AS HPP, MSalesD.KodeInv, MSalesD.NamaInv, MSalesD.Satuan, MSalesD.Barcode, MSalesD.IDSatuan, MSalesD.Konversi " & _
                "FROM MSalesD INNER JOIN MSales ON MSalesD.IDSales = MSales.NoID " & _
                "Where day(MSales.Tanggal) = " & DayReset & " and Month(MSales.Tanggal) = " & MonthReset & " and Year(MSales.Tanggal) = " & YearReset & "  And MSales.ShifT = " & NamaShiftReset & " " & _
                "GROUP BY MSales.IDUser,Msales.IDBank,MSalesD.IDInventor, MSalesD.KodeInv, MSalesD.NamaInv, MSalesD.Satuan, MSalesD.Barcode, MSalesD.IDSatuan, MSalesD.Konversi " & _
                "HAVING (((Sum(MSalesD.Qty))<>0) AND MSales.IDUser=" & rs!IDUser & " AND MSales.IDBank=" & NullToNol(rs!IDBank) & ")"
            DbsReset.Execute "UPDATE MSalesD SET MSalesD.HargaSat = [msalesd].[harga]/Abs([msalesd].[qty]) WHERE (((MSalesD.Qty)<>0)) AND IDSales=" & NoID
            End If
            DbsReset.Execute "Update MSALES  set Tanggal=#" & Mid(Text2.Text, 3, 2) & "/" & Mid(Text2.Text, 1, 2) & "/" & Mid(Text2.Text, 5, 4) & "#"
    '            If rs!Banka = 0 Then
    '                If KodeKasir = KodeUserDua Then
    '                  DbsReset.Execute "UPDATE MSalesD SET MSalesD.Qty=(MSalesD.Qty*" & PersenLap & ")\100,MSalesD.[harga]=((MSalesD.Qty*" & PersenLap & ")\100)*MSalesD.HargaSat WHERE  IDSales=" & NoId
    '                 End If
    '            End If
                NoID = NoID + 1
        rs.MoveNext
        Loop
     End If
 
    
    Set rs = Nothing
    dbs.Close
    Set RsReset = Nothing
    DbsReset.Close
    Set dbs = Nothing
    Set DbsReset = Nothing
'    Unload Me
End Sub
Sub AmbilPersenPerMesin()
Dim dbs As Database
Dim rs As Recordset
Set dbs = OpenDatabase(DIRPOS & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
  Set rs = dbs.OpenRecordset("Umum")
  If rs.EOF And rs.BOF Then
'    rs.AddNew
'    rs!Kassa = "01"
'    rs!IsNotaSelesai = True
'    rs!IDSalesAkhir = 1
'    rs.Update
'    lbKassa.Caption = ": 01"
'    NamaMesin = "01"
'    NamaToko = ""
'    NoPortPrinter = 1
'    NoPortDisplay = 2
'    NoPortBarcode = 3
  Else
'    NamaMesin = rs!Kassa
'    lbKassa.Caption = ": " & NamaMesin
'    IsNotaSelesai = rs!IsNotaSelesai
'    IDNotaTerakhir = rs!IDSalesAkhir
'    Judulstruk = rs!Judul
'    NamaToko = Trim(rs!Perusahaan)
'    NoPortPrinter = rs!NamaPrinter
'    NoPortBarcode = rs!Namabarcode
'    NoPortDisplay = rs!NamaCustomerDisplay
'    KodeUserDua = rs!kode
    PersenLap = NullToNol(getRegistry("Prosen", "Pengawas"))
'    lbStatus = ": " & GetStatusNetwork
'    'NamaToko =  'Trim(Mid(Judulstruk, 1, InStr(1, Judulstruk, Chr(13)) - 1))
'    Label3.Caption = NamaToko
  End If
  dbs.Close
End Sub
Sub cariKodeNama(ByRef KodeUser, ByRef NamaUser, ByVal IDUser)
Dim dbs As Database
Dim rs As Recordset
  Set dbs = OpenDatabase(DIRPOS & "\Database\dbMaster.mdb")
  Set rs = dbs.OpenRecordset("SELECT NoID,Kode,Nama From MEmp Where NoID=" & IDUser)
  If rs.BOF And rs.BOF Then
    KodeUser = "-"
    NamaUser = "-"
  Else
    KodeUser = rs!kode
    NamaUser = rs!Nama
  End If
rs.Close
Set rs = Nothing
dbs.Close
Set dbs = Nothing
End Sub

Sub CetakReset()
 'wis tak del
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub



Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim hasil As String
hasil = Trim(SendByCode(KeyCode))
KeyCode = 0
Select Case hasil
Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
  Text2.Text = Text2.Text & hasil
  Text2.SelStart = Len(Text2.Text)
 Case "SPC"
  Text2.Text = Text2.Text & " "
  Text2.SelStart = Len(Text2.Text)
Case "BKS"
  If Len(Text2.Text) > 0 Then Text2.Text = Left(Text2.Text, Len(Text2.Text) - 1)
   Text2.SelStart = Len(Text2.Text)
Case "CLR"
    Text2.Text = ""
Case "UP"
    Text1.SetFocus

Case "ENT"

    If Len(Text2.Text) <> 8 Then Exit Sub
    If Val(Mid(Text2.Text, 1, 2)) < 1 Or Val(Mid(Text2.Text, 1, 2)) > 31 Then Exit Sub
    If Val(Mid(Text2.Text, 3, 2)) < 1 Or Val(Mid(Text2.Text, 3, 2)) > 12 Then Exit Sub
    If Val(Mid(Text2.Text, 5, 4)) < 2004 Or Val(Mid(Text2.Text, 5, 4)) > 2500 Then Exit Sub
    TglReset = Text2.Text
    Text3.SetFocus
Case "ESC"
    Text1.SetFocus
End Select
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub




Private Sub Text3_DblClick()
Text3.Locked = False
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim hasil As String
  hasil = Trim(SendByCode(KeyCode))
  KeyCode = 0
  Select Case hasil
  Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
    If Left(KodeKasir, 2) = "ST" Then 'Khusus STAF dg awalan ST
      'Text3.Enabled = True
     Else
        hasil = ""
      End If
    Text3.Text = Text3.Text & hasil
    Text3.SelStart = Len(Text3.Text)
   Case "SPC"
    Text3.Text = Text3.Text & " "
    Text3.SelStart = Len(Text3.Text)
  Case "BKS"
    If Len(Text3.Text) > 0 Then Text3.Text = Left(Text3.Text, Len(Text3.Text) - 1)
     Text3.SelStart = Len(Text3.Text)
  Case "CLR"
    If KodeKasir = "ST03" Then 'Khusus AMiruddin
      Text3.Text = ""
    Else
        
    End If
      
Case "UP"
    Text2.SetFocus
  Case "ENT"
      If Val(Text3.Text) < 1 Or Val(Text3.Text) > 99 Then Exit Sub
      If Len(Text3.Text) = 1 Then
        Text3.Text = "0" & Text3.Text
      End If
      namaMesinReset = Text3.Text

      Text4.Text = "\\KASSA" & Format(Val(namaMesinReset), "0#") & "\pos"
      Text4.SetFocus
  Case "ESC"
      Text2.SetFocus
  End Select
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub Text4_DblClick()
Text4.Text = App.Path
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo pesan
Dim hasil As String
hasil = Trim(SendByCode(KeyCode))
KeyCode = 0
Select Case hasil
Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "\", ".", ";", ":"
  Text4.Text = Text4.Text & hasil
  Text4.SelStart = Len(Text4.Text)
 Case "SPC"
  Text4.Text = Text4.Text & " "
  Text4.SelStart = Len(Text4.Text)
Case "BKS"
  If Len(Text4.Text) > 0 Then Text4.Text = Left(Text4.Text, Len(Text4.Text) - 1)
   Text4.SelStart = Len(Text4.Text)
Case "CLR"
    If Text4.Text = "" Then
      Text4.Text = App.Path
    Else
      Text4.Text = ""
    End If
Case "UP"
    Text3.SetFocus
Case "ENT"
namaMesinReset = Text3.Text
    Text4.Locked = True
    If Text4.Text = "" Then Exit Sub
    If Dir(Text4.Text & "\database\TempDB" & Format(DateValue(Text2.Text), "_yyyyMM") & ".mdb") = "" Then
      frmPesan.lbPesan = "MESIN TIDAK ONLINE!!"
      frmPesan.Show 1
      Text4.Locked = False
      Exit Sub
    End If
    DIRPOS = Text4.Text
    DayReset = Val(Mid(Text2.Text, 1, 2))
    MonthReset = Val(Mid(Text2.Text, 3, 2))
    YearReset = Val(Mid(Text2.Text, 5, 4))
    AmbilPersenPerMesin
'    If Text3.Text = "09" Or Text3.Text = "10" Or Text3.Text = "20" Or Text3.Text = "99" Then
        CekTransaksiBermasalahBusana
'    Else
'        CekTransaksiBermasalah
'    End If
    If isDataOK Then
        Screen.MousePointer = vbHourglass
        If Text3.Text = "09" Or Text3.Text = "10" Or Text3.Text = "20" Or Text3.Text = "99" Then
            ResetBusana
        Else
            ResetRetail
        End If
        CetakVoid
        ResendServerOnline
      
      If (UCase(Trim(getRegistry("AutoDelete", "Data"))) = "Y") And (KodeKasir = KodeUserDua) Then
        HAPUSTRANSAKSI
      End If
      Screen.MousePointer = vbDefault
      
      frmPesan.lbPesan = "Selesai !!!!"
      frmPesan.Show 1
      Unload Me
    Else
      frmPesan.lbPesan = "Ada Transaksi Pending!!!!"
      frmPesan.Show 1
      Unload Me
    End If
    Text4.Locked = False
Case "ESC"
    Unload Me
End Select
Exit Sub
pesan:
    MsgBox "Ada kesalahan : " & Err.Description & vbCrLf & "Silahkan Tekan OK", vbCritical + vbOKOnly, "Pesan!"
    Resume Next
End Sub

Sub CetakVoid()
Dim dbc As Database
Dim rsc As Recordset

Set dbc = OpenDatabase(DIRPOS & "\Database\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
Set rsc = dbc.OpenRecordset("select MSales.Kode,MSales.Tanggal,MSalesD.KodeInv,MSalesD.NamaInv,MSalesD.Qty,MSalesD.Harga FROM MSalesD Inner Join MSales ON MSalesD.IDSales=MSales.NoID " & _
"where Day(Tanggal)=" & DayReset & " AND Month(Tanggal)=" & MonthReset & " AND Year(Tanggal)=" & YearReset & " AND Shift= " & NamaShiftReset & " AND (Transaksi='VOD' or Transaksi='CRC')")
If rsc.EOF And rsc.BOF Then
Else
Prin Chr(13) & Chr(10)
Prin Chr(13) & Chr(10)
Prin "---------------------------------------"
Prin "            DAFTAR ITEM VOID"
Prin "Tgl:" & DayReset & "-" & MonthReset & "-" & YearReset & ",Shift:" & NamaShiftReset & ", Ksr:" & namakasirReset
Prin "---------------------------------------"
  rsc.MoveFirst
  Do While Not rsc.EOF
  Prin "#" & rsc!kode & ", Jam :" & Format(rsc!TANGGAL, "HH:mm:nn")
      cetakdetil rsc!KodeInv, rsc!NamaInv, Format(rsc!QTY, "##0"), Format(rsc!harga, "###,###,##0"), Format(rsc!QTY * rsc!harga, "###,###,##0")
       
    rsc.MoveNext
  Loop
End If
rsc.Close
Set rsc = Nothing
dbc.Close
Set dbc = Nothing
Prin "---------------------------------------"

Prin Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10)
PrinBigChar Space(15 - Min(15, Len(NamaToko))) & UCase(NamaToko)
    papercut
Prin Judulstruk & Chr(13) & Chr(10) & "----------------------------------------"
End Sub
Sub HAPUSTRANSAKSI()
    Dim dbs As Database
    Set dbs = OpenDatabase(DIRPOS & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
    dbs.Execute "DELETE MSalesD.* FROM MSalesD INNER JOIN MSales ON MSalesD.IDSales=MSales.NoID WHERE Year(Tanggal)=" & Right(Text2.Text, 4) & " AND Month(Tanggal)=" & Mid(Text2.Text, 3, 2) & " AND Day(Tanggal)=" & Left(Text2.Text, 2) & " AND Shift=" & Text1.Text
    dbs.Execute "DELETE MSales.* FROM MSales WHERE Year(Tanggal)=" & Right(Text2.Text, 4) & " AND Month(Tanggal)=" & Mid(Text2.Text, 3, 2) & " AND Day(Tanggal)=" & Left(Text2.Text, 2) & " AND Shift=" & Text1.Text
    dbs.Close
    Set dbs = Nothing
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Public Sub ResendServerOnline()
  Dim ispending As Boolean
Dim IDSales As Long
  ispending = False
    If isRemcomendedOnline = True Then
      Dim kassa As String
      Dim nmfile As String
      Dim SQL As String
      Dim jumrec As Long
      Dim TANGGAL As String
      nmfile = App.Path & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb"

      DataLokal.DatabaseName = nmfile
       DataLokalDTL.DatabaseName = nmfile
      DataLokal.RecordSource = "Select * from MSales " & _
       "WHERE (IsUpload=0 AND ((MSales.Shift)=" & NamaShiftReset & ") AND  (Day(MSales.Tanggal)=" & DayReset & " and Month(MSales.Tanggal)=" & MonthReset & " and Year(MSales.Tanggal)=" & YearReset & "))"
        
      DataLokal.Refresh

      With DataLokal.Recordset
        If .EOF Or .BOF Then
        Else
        .MoveFirst
        Do While Not .EOF
          IDSales = !NoID
          TANGGAL = "convert(datetime,'" & Format(!TANGGAL, "MM/dd/yyyy") & "',101)"
  
          SQL = "INSERT INTO MSALESPOS(IDPos,NoIDSales,Kode,Tanggal,Tgl,Jam,TotalPajak,TotalDiscount," & _
          "SubTotal,DiscNota,HargaTotal,IDPayment,UangMuka,IDUser,IsSend,Bank,ISPending," & _
          "Shift,Voucher,IDBank,IDCustomer,IDVoucher,KodeMember,IDMember,BarangKSB,SisaKSB) VALUES(" & _
           IDPOSDef & "," & !NoID & ",'" & !kode & "',convert(datetime,'" & _
            Format(!TANGGAL, "MM/dd/yyyy hh:nn:ss") & "',101) ,convert(datetime,'" & _
            Format(!TANGGAL, "MM/dd/yyyy") & "',101) ,convert(datetime,'" & _
            Format(!TANGGAL, "hh:nn:ss") & "',14) ," & FixKoma(!TotalPajak) & "," & _
            FixKoma(!TotalDiscount) & "," & FixKoma(!SubTotal) & "," & FixKoma(!DiscNota) & "," & _
            FixKoma(!Hargatotal) & "," & !IDPayment & "," & FixKoma(!UangMuka) & "," & _
            !IDUser & "," & BoolToInt(!IsSend) & "," & FixKoma(!Bank) & "," & _
            BoolToInt(!ispending) & "," & !Shift & "," & FixKoma(!Voucher) & "," & _
            NullToNol(!IDBank) & "," & NullToNol(!idcustomer) & "," & _
            NullToNol(!IDVoucher) & ",'" & NullToStr(!KodeMember) & "'," & NullToNol(!IDMember) & "," & FixKoma(!BarangKSB) & "," & FixKoma(!SisaKSB) & ")"
            '999999: cara lama masih ada kemungkinan record sales dengan detil kosong kekirim yang berakibat fatal diserver
            If BoolToInt(!ispending) = 1 Then
              ispending = True
            Else
              ispending = False
            End If
  '            lbStatus.Caption = "Status : " & ExecuteSQL(sql)
  '          End If
              If ispending = False Then
                'Delete header ada kemungkinan keisi tapi tidak lengkap
                ExecuteSQL "Delete From MSalesPos where Shift=" & NamaShiftReset & " AND Tgl=" & TANGGAL & " AND NoIDSales=" & IDSales & " AND IDPos=" & IDPOSDef
                
                ExecuteSQL "Delete From MSalesPosD where Shift=" & NamaShiftReset & " AND Tanggal=" & TANGGAL & " AND NoIDSalesPos=" & IDSales & " AND IDPos=" & IDPOSDef
               DataLokalDTL.RecordSource = "Select * from MSalesd where IDSales=" & IDSales
                DataLokalDTL.Refresh
                  With DataLokalDTL.Recordset
                    If .EOF Or .BOF Then
                    Else
                    '999999: Pindah disini
                       ExecuteSQL (SQL) 'lbStatus.Caption = "Status : " &
                      .MoveFirst
                      Do While Not .EOF
                      
                      SQL = "INSERT INTO MSALESPOSD(NoIDSalesD,NoIDSalesPos,IDPos,IDGudang,IDInvsat," & _
                      "Qty,Harga,IDSatuan,Konversi,Transaksi,HargaPokok,HargaBruto,DiscRp," & _
                      "DiscProsen,IsDiscSupplier,Tanggal,IsMember) VALUES(" & _
                        !NoID & "," & !IDSales & "," & IDPOSDef & "," & IDGudangDef & "," & !IdInventor & "," & _
                        FixKoma(!QTY) & "," & FixKoma(!harga) & "," & !idSatuan & "," & FixKoma(!Konversi) & ",'" & _
                        !Transaksi & "'," & FixKoma(!HargaPokok) & "," & FixKoma(!HargaBruto) & "," & FixKoma(!DiscRp) & "," & _
                        FixKoma(!DiscProsen) & "," & 0 & "," & TANGGAL & "," & BoolToInt(!IsMember) & ")"
                      ExecuteSQL (SQL)
                        DoEvents
                      .MoveNext
                      Loop
                      DoEvents
                    End If
                  End With
                .Edit
                !IsUpload = 1
                .Update
                .Bookmark = .LastModified
            End If
            DoEvents
            If NullToNol(!IDMember) > 0 Then 'And NullToNol(!BarangKSB) >= 100000
                ExecuteSQL "Insert Into MCustomerPoint(IDCustomer,Kode,Tanggal,Kassa,Netto,Debet) Values(" & _
                NullToNol(!IDMember) & ",'" & NullToStr(!kode) & "'," & TANGGAL & ",''," & FixKoma(NullToNol(!BarangKSB)) & "," & NullToNol(!BarangKSB) \ 100000 & ")"
           End If
            .MoveNext
        Loop
        End If
        End With

        
   End If
    
End Sub
