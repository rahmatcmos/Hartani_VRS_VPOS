VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmMember 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1740
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   540
      IMEMode         =   3  'DISABLE
      Left            =   3120
      TabIndex        =   0
      Top             =   570
      Width           =   2475
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      NullDiscard     =   -1  'True
      RThreshold      =   1
      RTSEnable       =   -1  'True
      SThreshold      =   1
      InputMode       =   1
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Esc Untuk Non Member"
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
      Left            =   225
      TabIndex        =   21
      Top             =   810
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "F1 : SCK     F2 : SCN"
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
      Height          =   195
      Left            =   150
      TabIndex        =   20
      Top             =   1350
      Width           =   3735
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Masukan Kd/Nm Cust."
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
      Left            =   255
      TabIndex        =   19
      Top             =   510
      Width           =   3735
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
      Left            =   7830
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
      Left            =   7830
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
      Left            =   7830
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
      Left            =   7830
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
      Left            =   7830
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
      Left            =   7830
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
      Left            =   7830
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
      Left            =   7830
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
      Left            =   7830
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
      Left            =   5850
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
      Left            =   5850
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
      Left            =   5850
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
      Left            =   5850
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
      Left            =   5850
      TabIndex        =   5
      Top             =   1920
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
      Left            =   5850
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
      Left            =   5850
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
      Left            =   5850
      TabIndex        =   2
      Top             =   840
      Width           =   2205
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   1065
      Left            =   120
      Top             =   270
      Width           =   5595
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   1455
      Left            =   60
      Top             =   180
      Width           =   5745
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
      Left            =   5850
      TabIndex        =   1
      Top             =   480
      Width           =   2205
   End
End
Attribute VB_Name = "frmMember"
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
Dim BarcodeIn As String
Public IsAllow As Boolean

Private Sub Form_Activate()
BukaCommBarcode
End Sub

Private Sub Form_DeActivate()
TutupCommBarcode
End Sub
Sub BukaCommBarcode()
On Error Resume Next
MSComm1.CommPort = NoPortBarcode
MSComm1.PortOpen = True
End Sub

Sub TutupCommBarcode()
On Error Resume Next
MSComm1.PortOpen = False
End Sub

Private Sub Form_Load()
  isHasilKonversi = False
End Sub

Private Sub MSComm1_OnComm()
Dim kode As String
Select Case MSComm1.CommEvent
        ' Event messages.
        Case comEvReceive
            Dim buffer As Variant
            Dim pos As Integer
            buffer = MSComm1.Input
            'Debug.Print "Receive - " & StrConv(Buffer, vbUnicode)
            BarcodeIn = BarcodeIn & StrConv(buffer, vbUnicode)
            pos = InStr(1, BarcodeIn, Chr(13))
            If pos Then
                kode = Left(BarcodeIn, pos - 1)
                BarcodeIn = ""
                Text1.Text = kode
                SendKeys "{ENTER}", False
'
'              Unload Me
              Exit Sub
            End If
            'ShowData txtTerm, (StrConv(Buffer, vbUnicode))
End Select
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Hasil As String, IsAmbil As Boolean, kdMember As String
Hasil = Trim(SendByCode(KeyCode))
Select Case Hasil
Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "00", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "'", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "-", "+", "=", "|", "\", ".", ":", ";", "?", "[", "]"
  Text1.Text = Text1.Text & Hasil
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
    GetIDCUSTOMER
   ' CekTransaksiBermasalah
    If IDMember < 1 Then
'      CetakReset
    Else
    '  frmPesan.lbPesan = "Ada Transaksi Pending!!!!"
      'frmPesan.Show 1
      Unload Me
    End If
Case "SCK"
    frmCustomer.NamaField = "Kode"
    frmCustomer.lbCari = "CARI KODE"
    frmCustomer.Tampil IsAmbil, kdMember
    If IsAmbil And kdMember = "" Then
        Text1.Text = ""
    Else
        Text1.Text = kdMember
        GetIDCUSTOMER
        If IDMember < 1 Then
        Else
          Unload Me
        End If
    End If
Case "SCN"
    frmCustomer.NamaField = "Nama"
    frmCustomer.lbCari = "CARI NAMA"
    frmCustomer.Tampil IsAmbil, kdMember
    If IsAmbil And kdMember = "" Then
        Text1.Text = ""
    Else
        Text1.Text = kdMember
        GetIDCUSTOMER
        If IDMember < 1 Then
        Else
          Unload Me
        End If
    End If
Case "ESC"
    IDMember = -1
    NamaMember = ""
    kdMember = ""
    LimitHutang = 0
    Unload Me
End Select
End Sub
Public Sub GetIDCUSTOMER()
'On Error GoTo akhir
    If isOnline Then
      GetIDCUSTOMERONLINE
    Else
   GetIDCUSTOMERONLINE
        'GetStatusNetwork
      'GetIDCUSTOMERLOKAL
    End If
    
End Sub

Public Sub GetIDCUSTOMERONLINE()
GetIDCUSTOMERLOKAL
'   Dim sqlcon As New ADODB.Connection
'   Set sqlcon = New ADODB.Connection
'   Dim rs As New ADODB.Recordset
'   sqlcon.ConnectionString = Cnstr
'   sqlcon.Open
'   Set rs = sqlcon.Execute("Select NoID,Kode,Nama,IsActive from MCustomer where Kode='" & Text1.Text & "'")
'   If rs.EOF And rs.BOF Then
'        frmPesan.lbPesan = "Member tidak ada..."
'        frmPesan.Show 1
'
'        NamaMember = ""
'        KodeMember = ""
'        IDMember = -1
'   Else
'        If rs!Isactive Then
'            NamaMember = NullToStr(rs!Nama)
'            KodeMember = NullToStr(rs!kode)
'            IDMember = rs!NoId
'        Else
'            frmPesan.lbPesan = "No KSB Sudah Tidak Aktif..."
'            frmPesan.Show 1
'            NamaMember = ""
'            KodeMember = ""
'            IDMember = -1
'        End If
'    End If
'    sqlcon.Close
'    Set sqlcon = Nothing
End Sub

Public Sub GetIDCUSTOMERLOKAL()

   DataLokal.DatabaseName = App.path & "\database\dbmaster.mdb"
   DataLokal.RecordSource = "Select NoID,Kode,Nama,DiscountCustomer,SyaratMinimum,AllowInputDiscount from MCustomer where Kode='" & Replace(Text1.Text, "'", "''") & "' OR Barcode='" & Replace(Text1.Text, "'", "''") & "'"
   DataLokal.Refresh
    If DataLokal.Recordset.EOF And DataLokal.Recordset.BOF Then
        NamaMember = ""
        KodeMember = ""
        IDMember = -1
        LimitHutang = 0
  defDiscMember = 0
  defMinMemberdapatDisc = 0
  defDiscMemberBolehInput = False
    Else
     ' If IsAllow Then
        NamaMember = NullToStr(DataLokal.Recordset!Nama)
        KodeMember = NullToStr(DataLokal.Recordset!kode)
      '  LimitHutang = NullToNol(DataLokal.Recordset!LimitHutang)
        IDMember = DataLokal.Recordset!NoID
       ' TipeHargaJual = NullToNol(DataLokal.Recordset!TipeHargaJual)
 defDiscMember = NullToNol(DataLokal.Recordset!DiscountCustomer)
  defMinMemberdapatDisc = NullToNol(DataLokal.Recordset!SyaratMinimum)
  defDiscMemberBolehInput = NullToStr(DataLokal.Recordset!AllowInputDiscount)
'      Else
'        If KodeMember <> "" And TipeHargaJual <> NullToNol(DataLokal.Recordset!TipeHargaJual) Then
'          frmPesan.lbPesan.Caption = "Customer sudah ada, dan Tipe Harga Berbeda dengan sebelumnya."
'          frmPesan.Show 1
'          Exit Sub
'        Else
'          NamaMember = NullToStr(DataLokal.Recordset!Nama)
'          KodeMember = NullToStr(DataLokal.Recordset!kode)
'          LimitHutang = NullToNol(DataLokal.Recordset!LimitHutang)
'          IDMember = DataLokal.Recordset!NoID
'          TipeHargaJual = NullToNol(DataLokal.Recordset!TipeHargaJual)
'  '        Dim JumBayar As Double
'  '        Dim JumBelanja As Double
'  '        JumBelanja = NullToNol(ExecuteSkalarSQL("Select Sum(Total-UangMuka) as Hasil From Mjual where IDSupplier=" & IDMember))
'  '        JumBayar = NullToNol(ExecuteSkalarSQL("Select Sum(MBayarHutangD.Bayar+MBayarHutangD.Retur) as Hasil From MBayarHutangD LEFT join Mjual On MBayarHutangD.IDBeli=MJual.ID where MJual.IDSupplier=" & IDMember))
'  '        LimitHutang = LimitHutang - (JumBelanja - JumBayar)
'        End If
'      End If
    End If
    DataLokal.Database.Close
End Sub

Sub CekTransaksiBermasalah()
Dim dbc As Database
Dim rsc As Recordset
Set dbc = OpenDatabase(DirDatabase & "\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
Set rsc = dbc.OpenRecordset("SELECT MSales.NoID, Sum([Qty]*[Harga]) AS QSubTotal, MSales.SubTotal, MSales.DiscNota, MSales.HargaTotal, MSales.UangMuka, MSales.ISPending " & _
    "FROM MSales INNER JOIN MSalesD ON MSales.NoID = MSalesD.IDSales " & _
    "GROUP BY MSales.NoID, MSales.SubTotal, MSales.DiscNota, MSales.HargaTotal, MSales.UangMuka, MSales.ISPending " & _
    "HAVING (((Sum([Qty]*[Harga]))<>[SubTotal])) OR (((Sum([Qty]*[Harga]))<>[HargaTotal]+[DiscNota])) OR (((Sum([Qty]*[Harga]))<>[UangMuka]+[DiscNota]))")
If rsc.EOF And rsc.BOF Then
Else
  rsc.MoveFirst
  Do While Not rsc.EOF
    dbc.Execute "Update MSales set SubTotal=" & CCur(IIf(IsNull(rsc!QSubTotal), 0, rsc!QSubTotal)) & ", HargaTotal=" & CCur(IIf(IsNull(rsc!QSubTotal), 0, rsc!QSubTotal)) - CCur(IIf(IsNull(rsc!DiscNota), 0, rsc!DiscNota)) & ", ISpending=" & IIf((rsc!QSubTotal = (rsc!UangMuka + rsc!DiscNota)), False, True) & " Where NoId=" & rsc!NoID
  rsc.MoveNext
  Loop
End If
Set rsc = dbc.OpenRecordset("Select * From MSales Where IsPending=TRUE")
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

Sub Tampil(ByVal Bolehkan As Boolean)
  IsAllow = Bolehkan
  Me.Show 1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
