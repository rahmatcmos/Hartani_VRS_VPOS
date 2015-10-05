VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmPenerbitVoucher 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\AKTIF\BILKA AKHIR 2004\POS\Database\Bank.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "MBank"
      Top             =   210
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   180
      TabIndex        =   0
      Top             =   4275
      Width           =   7440
   End
   Begin TrueDBGrid60.TDBGrid DBGrid1 
      Bindings        =   "frmPenerbitVoucher.frx":0000
      Height          =   3285
      Left            =   180
      OleObjectBlob   =   "frmPenerbitVoucher.frx":0014
      TabIndex        =   21
      Top             =   630
      Width           =   7485
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   585
      TabIndex        =   23
      Top             =   4725
      Width           =   195
   End
   Begin VB.Label lbQty 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   180
      TabIndex        =   22
      Top             =   4725
      Width           =   165
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOMOR VOUCHER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   180
      TabIndex        =   20
      Top             =   4005
      Width           =   2385
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "PILIHVOUCHER, ENTER-SETUJU /  ESC-BATAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   150
      TabIndex        =   19
      Top             =   330
      Width           =   7035
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Transaksi Pending"
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
      Height          =   345
      Index           =   8
      Left            =   11220
      TabIndex        =   18
      Top             =   3390
      Width           =   3405
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Bank"
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
      Height          =   345
      Index           =   7
      Left            =   11220
      TabIndex        =   17
      Top             =   3000
      Width           =   3405
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Tunai"
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
      Height          =   345
      Index           =   6
      Left            =   11220
      TabIndex        =   16
      Top             =   2640
      Width           =   3405
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Total"
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
      Height          =   345
      Index           =   5
      Left            =   11220
      TabIndex        =   15
      Top             =   2280
      Width           =   3405
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Diskon"
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
      Height          =   345
      Index           =   4
      Left            =   11220
      TabIndex        =   14
      Top             =   1920
      Width           =   3405
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah SubTotal"
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
      Height          =   345
      Index           =   3
      Left            =   11220
      TabIndex        =   13
      Top             =   1560
      Width           =   3405
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Nota"
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
      Height          =   345
      Index           =   2
      Left            =   11220
      TabIndex        =   12
      Top             =   1200
      Width           =   3405
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Mesin-Kasir"
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
      Height          =   345
      Index           =   1
      Left            =   11220
      TabIndex        =   11
      Top             =   840
      Width           =   3405
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
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
      Height          =   345
      Index           =   0
      Left            =   11220
      TabIndex        =   10
      Top             =   480
      Width           =   3405
   End
   Begin VB.Label lblNama 
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher"
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
      Height          =   345
      Index           =   8
      Left            =   9240
      TabIndex        =   9
      Top             =   3360
      Width           =   2205
   End
   Begin VB.Label lblNama 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Bank"
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
      Height          =   345
      Index           =   7
      Left            =   9240
      TabIndex        =   8
      Top             =   3000
      Width           =   2205
   End
   Begin VB.Label lblNama 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Tunai"
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
      Height          =   345
      Index           =   6
      Left            =   9240
      TabIndex        =   7
      Top             =   2640
      Width           =   2205
   End
   Begin VB.Label lblNama 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Total"
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
      Height          =   345
      Index           =   5
      Left            =   9240
      TabIndex        =   6
      Top             =   2280
      Width           =   2205
   End
   Begin VB.Label lblNama 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Diskon"
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
      Height          =   345
      Index           =   4
      Left            =   9240
      TabIndex        =   5
      Top             =   2340
      Width           =   2205
   End
   Begin VB.Label lblNama 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah SubTotal"
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
      Height          =   345
      Index           =   3
      Left            =   9240
      TabIndex        =   4
      Top             =   1560
      Width           =   2205
   End
   Begin VB.Label lblNama 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Nota"
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
      Height          =   345
      Index           =   2
      Left            =   9240
      TabIndex        =   3
      Top             =   1200
      Width           =   2205
   End
   Begin VB.Label lblNama 
      BackStyle       =   0  'Transparent
      Caption         =   "Mesin-Kasir"
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
      Height          =   345
      Index           =   1
      Left            =   9240
      TabIndex        =   2
      Top             =   840
      Width           =   2205
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   5025
      Left            =   120
      Top             =   180
      Width           =   7680
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   5220
      Left            =   90
      Top             =   60
      Width           =   7800
   End
   Begin VB.Label lblNama 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
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
      Height          =   345
      Index           =   0
      Left            =   9240
      TabIndex        =   1
      Top             =   480
      Width           =   2205
   End
End
Attribute VB_Name = "frmPenerbitVoucher"
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
Dim GNominal As Long
Dim IDPenerbitVoucher_ As Integer
Dim KodePenerbitVoucher_ As String
Dim Qty_ As Integer
Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Data1.Recordset.BOF Or Data1.Recordset.EOF Then
    GNominal = 0
Else
    GNominal = NullToNol(Data1.Recordset!Nominal)
End If
End Sub

Private Sub Form_Activate()
  Text1.SetFocus
End Sub

Private Sub Form_Load()
  isHasilKonversi = False
 ' Label1.Caption = "JUMLAH BELUM DIBAYAR : " & Format(GNominal, "###,###,##0")
  Data1.DatabaseName = App.path & "\database\dbmaster.mdb"
  Data1.RecordSource = "SELECT * FROM MPenerbitVoucher Order by Nominal"
  Data1.Refresh
 If Data1.Recordset.BOF Or Data1.Recordset.EOF Then
    GNominal = 0
Else
    GNominal = NullToNol(Data1.Recordset!Nominal)
End If
lbQTY.Caption = "1"
Qty_ = 1
End Sub
 

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Data1.Database.Close
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim hasil As String
Dim SelBks
hasil = Trim(SendByCode(KeyCode))
KeyCode = 0
Select Case hasil
  Case "0", "00", "1", "2", "3", "4", "5", "6", "7", "8", "9", "`", "!", "@", "#", "$", "%", "^", "&", "<", "(", ")", "-", "+", "{", "}", "[", "]", "/", "?", ".", ","
    Text1.Text = Text1.Text & hasil
    Text1.SelStart = Len(Text1.Text)
  Case "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"
    Text1.Text = Text1.Text & hasil
    Text1.SelStart = Len(Text1.Text)
  Case "CLR"
    Text1.Text = ""
  Case "*"
    lbQTY.Caption = Text1.Text
    Text1.Text = ""
  Case "BKS"
    If Len(Text1.Text) > 0 Then
      Text1.Text = Left(Text1.Text, Len(Text1.Text) - 1)
      Text1.SelStart = Len(Text1.Text)
    End If
  Case "UP"
    Data1.Recordset.MovePrevious
    If Data1.Recordset.BOF Then
      Data1.Recordset.MoveNext
    End If
    Set SelBks = DBGrid1.SelBookmarks
    While SelBks.Count <> 0
        SelBks.Remove 0
    Wend
    DBGrid1.SelBookmarks.Add DBGrid1.Bookmark
  Case "DN"
    Data1.Recordset.MoveNext
    If Data1.Recordset.EOF Then
      Data1.Recordset.MovePrevious
    End If
    Set SelBks = DBGrid1.SelBookmarks
    While SelBks.Count <> 0
        SelBks.Remove 0
    Wend
    DBGrid1.SelBookmarks.Add DBGrid1.Bookmark
  Case "ENT"
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then Exit Sub
'    If Trim(Text1.Text) = "" Then
'      frmPesan.lbPesan = "ISILAH NOMOR KARTU...!"
'      frmPesan.Show 1
'      Exit Sub
'    End If
    IDPenerbitVoucher_ = Data1.Recordset("NoID")
    If IsNumeric(lbQTY.Caption) Then
        Qty_ = lbQTY.Caption
    Else
        Qty_ = 1
    End If
    KodePenerbitVoucher_ = Data1.Recordset("Kode")
    jawab = True
    Unload Me
  Case "ESC"
    IDPenerbitVoucher_ = -1
    KodePenerbitVoucher_ = ""
    jawab = False
    Unload Me
End Select
End Sub
Sub View()
On Error Resume Next
Dim SelBks
  Set SelBks = DBGrid1.SelBookmarks
  While SelBks.Count <> 0
      SelBks.Remove 0
  Wend
  DBGrid1.SelBookmarks.Add DBGrid1.Bookmark
End Sub
Sub Tampil(ByRef jawaban As Boolean, ByRef IDPenerbitVoucher As Integer, ByRef KodePenerbitVoucher As String, ByRef Nominal As Double, ByRef Qty As Integer)
''  KeyOk = Key
'  GNominal = Nominal
'  IsFilterCC_ = IsFilterCC
'  IsMemberCharge = IsMemberCharge_
  Me.Show 1
  jawaban = jawab
  IDPenerbitVoucher = IDPenerbitVoucher_
  KodePenerbitVoucher = KodePenerbitVoucher_
  Nominal = GNominal
  Qty = Qty_
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
