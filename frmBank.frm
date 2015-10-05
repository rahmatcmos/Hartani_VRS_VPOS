VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmBank 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6300
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\AKTIF\BILKA AKHIR 2004\POS\Database\Bank.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "MBank"
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
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
      Top             =   5670
      Width           =   5415
   End
   Begin TrueDBGrid60.TDBGrid DBGrid1 
      Bindings        =   "frmBank.frx":0000
      Height          =   1080
      Left            =   180
      OleObjectBlob   =   "frmBank.frx":0014
      TabIndex        =   27
      Top             =   1890
      Width           =   5415
   End
   Begin TrueDBGrid60.TDBGrid DBGrid2 
      Bindings        =   "frmBank.frx":4155
      Height          =   2340
      Left            =   180
      OleObjectBlob   =   "frmBank.frx":4169
      TabIndex        =   28
      Top             =   3015
      Width           =   5415
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOMOR KARTU DEBIT / KREDIT"
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
      TabIndex        =   26
      Top             =   5370
      Width           =   4020
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   3360
      X2              =   5490
      Y1              =   1140
      Y2              =   1140
   End
   Begin VB.Label lbTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
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
      Left            =   4965
      TabIndex        =   25
      Top             =   1140
      Width           =   540
   End
   Begin VB.Label LbBiayaCC 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
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
      Left            =   4965
      TabIndex        =   24
      Top             =   810
      Width           =   540
   End
   Begin VB.Label LbBelumbayar 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
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
      Left            =   4965
      TabIndex        =   23
      Top             =   480
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JUMLAH TOTAL"
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
      Top             =   1140
      Width           =   1980
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BIAYA KARTU KREDIT"
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
      TabIndex        =   21
      Top             =   810
      Width           =   2835
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BELUM DIBAYAR "
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
      Top             =   480
      Width           =   2280
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "PILIH BANK , ENTER-SETUJU /  ESC-BATAL"
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
      Top             =   1590
      Width           =   5505
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
      Height          =   6015
      Left            =   120
      Top             =   150
      Width           =   5565
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   6165
      Left            =   60
      Top             =   60
      Width           =   5685
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
Attribute VB_Name = "frmBank"
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
Dim BiayaCC As Long
Dim totalBayarCC As Long
Dim IDBankPilih As Integer
Dim IDBankServerPilih As Integer
Dim IDJenisKartu_ As Integer
Dim NoAcc As String
Dim KodeBank As String
Dim NamaJenisKartu_ As String
Dim IsMemberCharge As Boolean
Dim NamaBank As String
Dim ChargeBank As Double
Dim IsFilterCC_ As Boolean

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'If Data1.Recordset.BOF Or Data1.Recordset.EOF Then
'    totalBayarCC = GNominal
'    BiayaCC = totalBayarCC - GNominal
'Else
'    BiayaCC = GNominal * NullToNol(Data1.Recordset!ProsenBiaya) / 100 '/  ((100 - NullToNol(Data1.Recordset!ProsenBiaya)) / 100)
'    totalBayarCC = GNominal + BiayaCC
'End If
'LbBiayaCC = Format(BiayaCC, "###,##0")
'lbTotal = Format(totalBayarCC, "###,##0")
End Sub

Private Sub DBGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Data2.Recordset.BOF Or Data2.Recordset.EOF Then
    totalBayarCC = GNominal
    BiayaCC = totalBayarCC - GNominal
Else
    If IsFilterCC_ And IsMemberCharge Then 'Member dan dapat diskon kena charge 2% khusus Kartu kredit
      BiayaCC = GNominal * NullToNol(2#) / 100
  Else
    BiayaCC = GNominal * NullToNol(Data2.Recordset!Charge) / 100 '/  ((100 - NullToNol(Data1.Recordset!ProsenBiaya)) / 100)
  End If
  totalBayarCC = GNominal + BiayaCC
End If
LbBiayaCC = Format(BiayaCC, "###,##0")
lbTotal = Format(totalBayarCC, "###,##0")
End Sub

Private Sub Form_Activate()
  Text1.SetFocus
End Sub

Private Sub Form_Load()
  isHasilKonversi = False
 ' Label1.Caption = "JUMLAH BELUM DIBAYAR : " & Format(GNominal, "###,###,##0")
  Data1.DatabaseName = App.path & "\database\dbmaster.mdb"
If IsFilterCC_ Then
  Data1.RecordSource = "SELECT * FROM MBank where IsKartuKredit=True  Order by Kode"
Else
Data1.RecordSource = "SELECT * FROM MBank where IsKartuDebet=True Order by Kode"
End If
  Data1.Refresh

  Data2.DatabaseName = App.path & "\database\dbmaster.mdb"
If IsFilterCC_ Then
  Data2.RecordSource = "SELECT * FROM MJenisKartu where IsKartuKredit=True Order by Kode"
Else
  Data2.RecordSource = "SELECT * FROM MJenisKartu where IsKartuDebet=True Order by Kode"
End If
  Data2.Refresh
If Data2.Recordset.BOF Or Data2.Recordset.EOF Then
  totalBayarCC = GNominal
  BiayaCC = totalBayarCC - GNominal
Else
  If IsFilterCC_ And IsMemberCharge Then 'Member dan dapat diskon kena charge 2% khusus Kartu kredit
      BiayaCC = GNominal * NullToNol(2#) / 100
  Else
    BiayaCC = GNominal * NullToNol(Data2.Recordset!Charge) / 100 '/  ((100 - NullToNol(Data1.Recordset!ProsenBiaya)) / 100)
  End If
  totalBayarCC = GNominal + BiayaCC
End If

  LbBelumbayar = Format(GNominal, "###,##0")
  LbBiayaCC = Format(BiayaCC, "###,##0")
  lbTotal = Format(totalBayarCC, "###,##0")
End Sub

Private Sub gridJenis_Click()

End Sub

Private Sub gridJenis_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Data1.Database.Close
Data2.Database.Close

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Hasil As String
Dim SelBks
Hasil = Trim(SendByCode(KeyCode))
KeyCode = 0
Select Case Hasil
  Case "0", "00", "1", "2", "3", "4", "5", "6", "7", "8", "9", "*", "`", "!", "@", "#", "$", "%", "^", "&", "<", "(", ")", "-", "+", "{", "}", "[", "]", "/", "?", ".", ","
    Text1.Text = Text1.Text & Hasil
    Text1.SelStart = Len(Text1.Text)
  Case "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"
    Text1.Text = Text1.Text & Hasil
    Text1.SelStart = Len(Text1.Text)
  Case "CLR"
    Text1.Text = ""
  Case "BKS"
    If Len(Text1.Text) > 0 Then
      Text1.Text = Left(Text1.Text, Len(Text1.Text) - 1)
      Text1.SelStart = Len(Text1.Text)
    End If
  Case "UP"
    Data2.Recordset.MovePrevious
    If Data2.Recordset.BOF Then
      Data2.Recordset.MoveNext
    End If
    Set SelBks = DBGrid2.SelBookmarks
    While SelBks.Count <> 0
        SelBks.Remove 0
    Wend
    DBGrid2.SelBookmarks.Add DBGrid2.Bookmark
  Case "DN"
    Data2.Recordset.MoveNext
    If Data2.Recordset.EOF Then
      Data2.Recordset.MovePrevious
    End If
    Set SelBks = DBGrid2.SelBookmarks
    While SelBks.Count <> 0
        SelBks.Remove 0
    Wend
    DBGrid2.SelBookmarks.Add DBGrid2.Bookmark
  Case "ENT"
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then Exit Sub
    If Data2.Recordset.EOF Or Data2.Recordset.BOF Then Exit Sub
    If Trim(Text1.Text) = "" Then
      frmPesan.lbPesan = "ISILAH NOMOR KARTU...!"
      frmPesan.Show 1
      Exit Sub
    End If
    IDBankPilih = Data1.Recordset("NoID")
    IDBankServerPilih = Data1.Recordset("NoId")
    KodeBank = Data1.Recordset("Kode")
    NamaBank = Data1.Recordset("Nama")
    NamaJenisKartu_ = NullToStr(Data2.Recordset("Nama"))
    IDJenisKartu_ = Data2.Recordset("NoID")
    ChargeBank = NullToNol(Data2.Recordset("Charge"))
    NoAcc = Text1.Text
    
    jawab = True
    Unload Me
  Case "ESC"
    IDBankPilih = 0
    IDBankServerPilih = 0
    BiayaCC = 0
    totalBayarCC = 0
    IDBankPilih = 0
    IDBankServerPilih = 0
    IDJenisKartu_ = 0
    NamaJenisKartu_ = ""
    NoAcc = ""
    KodeBank = ""
    NamaBank = ""
    ChargeBank = 0
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
Sub Tampil(ByRef jawaban As Boolean, IDBank As Integer, IDBankServer As Integer, NoAccCust As String, KdBank As String, NmBank As String, CgBank As Double, Key As Integer, ByVal Nominal As Long, ByRef Biaya As Double, ByRef totalBayar As Long, ByVal IsFilterCC As Boolean, ByRef IDJenisKartu As Integer, ByRef NamaJenisKartu, ByVal IsMemberCharge_ As Boolean)
  KeyOk = Key
  GNominal = Nominal
  IsFilterCC_ = IsFilterCC
  IsMemberCharge = IsMemberCharge_
  Me.Show 1
  Biaya = BiayaCC
  totalBayar = totalBayarCC
  IDBank = IDBankPilih
  IDBankServer = IDBankServerPilih
  IDJenisKartu = IDJenisKartu_
  NamaJenisKartu = NamaJenisKartu_
  NoAccCust = NoAcc
  KdBank = KodeBank
  NmBank = NamaBank
  CgBank = ChargeBank
  jawaban = jawab
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
