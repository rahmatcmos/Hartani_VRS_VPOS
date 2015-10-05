VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmReset1Saja 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3060
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   5760
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      Top             =   2190
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
      Top             =   1470
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
      Top             =   930
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
      Top             =   390
      Width           =   1275
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
      Top             =   3720
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
      Top             =   3360
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
      Top             =   1890
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
      Top             =   1530
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
      Top             =   990
      Width           =   3735
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Reset X atau Z (1=X atau 2=Z)                Cetak <ENTER>, Keluar <ESC>"
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
      Top             =   390
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
      Left            =   9210
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
      Left            =   9210
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
      Left            =   9210
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
      Left            =   9210
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
      Left            =   9210
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
      Left            =   9210
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
      Left            =   9210
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
      Left            =   9210
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
      Left            =   9210
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
      Left            =   7230
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
      Left            =   7230
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
      Left            =   7230
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
      Left            =   7230
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
      Left            =   7230
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
      Left            =   7230
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
      Left            =   7230
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
      Left            =   7230
      TabIndex        =   2
      Top             =   840
      Width           =   2205
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   2595
      Left            =   150
      Top             =   270
      Width           =   6795
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   2745
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
      Left            =   7230
      TabIndex        =   1
      Top             =   480
      Width           =   2205
   End
End
Attribute VB_Name = "frmReset1Saja"
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
Sub CetakVoid()
If DefCetakVoidSaatReset = 0 Then 'Tidak perlu cetak void
    Prin "---------------------------------------"
    
    Prin Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10)
    PrinBigChar Space(15 - Min(15, Len(NamaToko))) & UCase(NamaToko)
        papercut
    Prin Judulstruk & Chr(13) & Chr(10) & "----------------------------------------"

Else
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
End If
End Sub
Private Sub Form_Load()
  isHasilKonversi = False
  Text1.Text = 1
'  If Format(Time, "HHnnss") > "150000" Then
'    Text1.Text = 2
'  Else
'    Text1.Text = 1
'  End If
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
    Dim DbsReset As Database
    Dim dbs As Database
    Dim rs As Recordset
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
        ',MsalesD.Qty*MSalesD.Harga as Jumlah
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

      Text4.Text = App.path '"\\KASSA" & Format(Val(namaMesinReset), "0#") & "\pos"
      Text4.SetFocus
  Case "ESC"
      Text2.SetFocus
  End Select
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub Text4_DblClick()
  Text4.Text = App.path
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo pesan
Dim hasil As String
hasil = Trim(SendByCode(KeyCode))
KeyCode = 0
Select Case hasil
Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "\", ".", ":", ";", "-", "_"
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
      Text4.Text = App.path
    Else
      Text4.Text = ""
    End If
Case "UP"
    Text3.SetFocus
Case "ENT"
'    If BolehReset(Text1.Text, Text2.Text) = False Then
'      frmPesan.lbPesan = "MAAF BELUM BOLEH RESET!!"
'      frmPesan.Show 1
'      Text4.Locked = False
'      Exit Sub
'    End If
'
'    namaMesinReset = Text3.Text
'    Text4.Locked = True
'    If Text4.Text = "" Then Exit Sub
'    If Dir(Text4.Text & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb") = "" Then
'      frmPesan.lbPesan = "MESIN TIDAK ONLINE!!"
'      frmPesan.Show 1
'      Text4.Locked = False
'      Exit Sub
'    End If
    DIRPOS = Text4.Text
    DayReset = Val(Mid(Text2.Text, 1, 2))
    MonthReset = Val(Mid(Text2.Text, 3, 2))
    YearReset = Val(Mid(Text2.Text, 5, 4))
'    AmbilPersenPerMesin
'    'If Text3.Text = "09" Or Text3.Text = "10" Or Text3.Text = "20" Or Text3.Text = "99" Then
'
'        'Hartani 28/02/2012
'        CekTransaksiBermasalahBusana
'
'   ' Else
'   '     CekTransaksiBermasalah
'    'End If
'
'    If isDataOK Then
'        Screen.MousePointer = vbHourglass
'        Text4.Enabled = False
''        If Text3.Text = "09" Or Text3.Text = "10" Or Text3.Text = "20" Or Text3.Text = "99" Then
''            ResetBusana
''        Else
'            ResetRetail
''        End If
'        CetakVoid
'        ResendServerOnline
'
'      If (UCase(Trim(getRegistry("AutoDelete", "Data"))) = "Y") And (KodeKasir = KodeUserDua) Then
'        HAPUSTRANSAKSI
'      End If
'      Text4.Enabled = True
'      Screen.MousePointer = vbDefault
'      PRINTRESET
'      frmPesan.lbPesan = "Selesai !!!!"
'      frmPesan.Show 1
'      Unload Me
'    Else
'      frmPesan.lbPesan = "Ada Transaksi Pending!!!!"
'      frmPesan.Show 1
'      Unload Me
'    End If
'    Text4.Locked = False
  Text4.Enabled = False
  DoEvents
  ResetHartani
  DoEvents
  frmPesan.lbPesan = "Kirim Data Ke Server....."
  frmPesan.Show 1
  ResendServerOnline
  frmPesan.lbPesan = "Selesai !!!!"
  frmPesan.Show 1
  Unload Me
Case "ESC"
    Unload Me
End Select
Exit Sub
pesan:
    MsgBox "Ada kesalahan : " & Err.Description & vbCrLf & "Silahkan Tekan OK", vbCritical + vbOKOnly, "Pesan!"

    Resume Next
End Sub

Private Sub ResetHartani()
Dim dbs As Database
Dim RsReset As Recordset
Dim strqry As String
Dim rs As Recordset
    '========MReset
      Set dbs = OpenDatabase(DIRPOS & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
      Set RsReset = dbs.OpenRecordset("SELECT * FROM MReset WHERE Day(Tanggal)=" & DayReset & " AND MONTH(Tanggal)=" & MonthReset & " AND YEAR(Tanggal)=" & YearReset)
      strqry = "SELECT SUM(Subtotal) AS Subtotal1, SUM(DiscIntern) AS DiscNota1, SUM(Pembulatan) AS Pembulatan1, SUM(Subtotal-DiscIntern-Pembulatan) AS Total1, SUM(Voucher) AS Voucher1," & vbCrLf & _
               " SUM(HargaTotal-Bank-Voucher) AS Tunai1, " & vbCrLf & _
               " (SELECT SUM(BANK) FROM MSales WHERE IDBank=1 AND Day(Tanggal)=" & DayReset & " AND Month(Tanggal)=" & MonthReset & " AND Year(Tanggal)=" & YearReset & " AND Shift= " & NamaShiftReset & ") AS Debet, " & vbCrLf & _
               " (SELECT SUM(MSales.BANK-(MSales.HargaTotal-(MSales.SubTotal-MSales.DiscIntern-MSales.Pembulatan))) FROM MSales WHERE IDBank=2 AND Day(Tanggal)=" & DayReset & " AND Month(Tanggal)=" & MonthReset & " AND Year(Tanggal)=" & YearReset & " AND Shift= " & NamaShiftReset & ") AS Kredit, " & vbCrLf & _
               " (SELECT SUM(MSales.HargaTotal-(MSales.SubTotal-MSales.DiscIntern-MSales.Pembulatan)) FROM MSales WHERE Day(Tanggal)=" & DayReset & " AND Month(Tanggal)=" & MonthReset & " AND Year(Tanggal)=" & YearReset & " AND Shift= " & NamaShiftReset & ") AS Charge1, " & vbCrLf & _
               " (SELECT COUNT(NoID) FROM MSales WHERE Subtotal>0 AND Day(Tanggal)=" & DayReset & " AND Month(Tanggal)=" & MonthReset & " AND Year(Tanggal)=" & YearReset & " AND Shift= " & NamaShiftReset & ") AS JumNota, " & vbCrLf & _
               " (SELECT SUM(BANK) FROM MSales WHERE IDJenisKartu=5 AND Day(Tanggal)=" & DayReset & " AND Month(Tanggal)=" & MonthReset & " AND Year(Tanggal)=" & YearReset & " AND Shift= " & NamaShiftReset & ") AS DebetMandiri, " & vbCrLf & _
               " (SELECT Count(BANK) FROM MSales WHERE IDJenisKartu=5 AND Day(Tanggal)=" & DayReset & " AND Month(Tanggal)=" & MonthReset & " AND Year(Tanggal)=" & YearReset & " AND Shift= " & NamaShiftReset & ") AS CountDebetMandiri, " & vbCrLf & _
               " (SELECT SUM(BANK) FROM MSales WHERE IDJenisKartu=3 AND Day(Tanggal)=" & DayReset & " AND Month(Tanggal)=" & MonthReset & " AND Year(Tanggal)=" & YearReset & " AND Shift= " & NamaShiftReset & ") AS DebetBCA, " & vbCrLf & _
               " (SELECT COUNT(BANK) FROM MSales WHERE IDJenisKartu=3 AND Day(Tanggal)=" & DayReset & " AND Month(Tanggal)=" & MonthReset & " AND Year(Tanggal)=" & YearReset & " AND Shift= " & NamaShiftReset & ") AS CountDebetBCA, " & vbCrLf & _
               " (SELECT SUM(BANK) FROM MSales WHERE IDJenisKartu=6 AND Day(Tanggal)=" & DayReset & " AND Month(Tanggal)=" & MonthReset & " AND Year(Tanggal)=" & YearReset & " AND Shift= " & NamaShiftReset & ") AS DebetLain, " & vbCrLf & _
               " (SELECT Count(BANK) FROM MSales WHERE IDJenisKartu=6 AND Day(Tanggal)=" & DayReset & " AND Month(Tanggal)=" & MonthReset & " AND Year(Tanggal)=" & YearReset & " AND Shift= " & NamaShiftReset & ") AS CountDebetLain, " & vbCrLf & _
               " (SELECT SUM(BANK-(CHARGE/100*BANK)) FROM MSales WHERE IDJenisKartu=1 AND Day(Tanggal)=" & DayReset & " AND Month(Tanggal)=" & MonthReset & " AND Year(Tanggal)=" & YearReset & " AND Shift= " & NamaShiftReset & ") AS CreditVisa, " & vbCrLf & _
               " (SELECT COUNT(BANK) FROM MSales WHERE IDJenisKartu=1 AND Day(Tanggal)=" & DayReset & " AND Month(Tanggal)=" & MonthReset & " AND Year(Tanggal)=" & YearReset & " AND Shift= " & NamaShiftReset & ") AS CountCreditVisa, " & vbCrLf & _
               " (SELECT SUM(BANK-(CHARGE/100*BANK)) FROM MSales WHERE IDJenisKartu=2 AND Day(Tanggal)=" & DayReset & " AND Month(Tanggal)=" & MonthReset & " AND Year(Tanggal)=" & YearReset & " AND Shift= " & NamaShiftReset & ") AS CreditMaster, " & vbCrLf & _
               " (SELECT Count(BANK) FROM MSales WHERE IDJenisKartu=2 AND Day(Tanggal)=" & DayReset & " AND Month(Tanggal)=" & MonthReset & " AND Year(Tanggal)=" & YearReset & " AND Shift= " & NamaShiftReset & ") AS CountCreditMaster, " & vbCrLf & _
               " (SELECT SUM(BANK-(CHARGE/100*BANK)) FROM MSales WHERE IDJenisKartu=4 AND Day(Tanggal)=" & DayReset & " AND Month(Tanggal)=" & MonthReset & " AND Year(Tanggal)=" & YearReset & " AND Shift= " & NamaShiftReset & ") AS CreditBCA, " & vbCrLf & _
               " (SELECT Count(BANK) FROM MSales WHERE IDJenisKartu=4 AND Day(Tanggal)=" & DayReset & " AND Month(Tanggal)=" & MonthReset & " AND Year(Tanggal)=" & YearReset & " AND Shift= " & NamaShiftReset & ") AS CountCreditBCA, " & vbCrLf & _
               " (SELECT SUM(BANK-(CHARGE/100*BANK)) FROM MSales WHERE IDJenisKartu=7 AND Day(Tanggal)=" & DayReset & " AND Month(Tanggal)=" & MonthReset & " AND Year(Tanggal)=" & YearReset & " AND Shift= " & NamaShiftReset & ") AS CreditLain, " & vbCrLf & _
               " (SELECT Count(BANK) FROM MSales WHERE IDJenisKartu=7 AND Day(Tanggal)=" & DayReset & " AND Month(Tanggal)=" & MonthReset & " AND Year(Tanggal)=" & YearReset & " AND Shift= " & NamaShiftReset & ") AS CountCreditLain, " & vbCrLf & _
               " (SELECT SUM(MSalesD.Jumlah) FROM MSales INNER JOIN MSalesD ON MSalesD.IDSales=MSales.NoID WHERE UCASE(MSalesD.Transaksi)='RTN' AND Day(MSales.Tanggal)=" & DayReset & " AND Month(MSales.Tanggal)=" & MonthReset & " AND Year(MSales.Tanggal)=" & YearReset & " AND MSales.Shift= " & NamaShiftReset & ") AS Retur " & vbCrLf & _
               " FROM MSales  " & vbCrLf & _
               " WHERE Day(Tanggal)=" & DayReset & " AND Month(Tanggal)=" & MonthReset & " AND Year(Tanggal)=" & YearReset
      Set rs = dbs.OpenRecordset(strqry)
      If Not (rs.EOF And rs.BOF) Then
        RsReset.MoveFirst
        RsReset.Edit
        If Text1.Text = 2 Then
          If Not NullToBool(RsReset!ResetZ) Then
            RsReset!ResetZ = True
            RsReset!SubTotal = NullToNol(rs!SubTotal1)
            RsReset!DiscMember = NullToNol(rs!DiscNota1)
            RsReset!Pembulatan = NullToNol(rs!Pembulatan1)
            RsReset!Total = NullToNol(rs!Total1)
            RsReset!Voucher = NullToNol(rs!Voucher1)
            RsReset!Tunai = NullToNol(rs!Tunai1)
            
            RsReset!Debet = NullToNol(rs!Debet)
            RsReset!Credit = NullToNol(rs!Kredit)
            RsReset!JumlahNota = NullToNol(rs!JumNota)
            RsReset!DebetMandiri = NullToNol(rs!DebetMandiri)
            RsReset!DebetBCA = NullToNol(rs!DebetBCA)
            RsReset!DebetLain = NullToNol(rs!DebetLain)
            RsReset!CreditVisa = NullToNol(rs!CreditVisa)
            RsReset!CreditBCA = NullToNol(rs!CreditBCA)
            RsReset!CreditMaster = NullToNol(rs!CreditMaster)
            RsReset!CreditLain = NullToNol(rs!CreditLain)
            RsReset!CountDebetMandiri = NullToNol(rs!CountDebetMandiri)
            RsReset!CountDebetBCA = NullToNol(rs!CountDebetBCA)
            RsReset!CountDebetLain = NullToNol(rs!CountDebetLain)
            RsReset!CountCreditVisa = NullToNol(rs!CountCreditVisa)
            RsReset!CountCreditBCA = NullToNol(rs!CountCreditBCA)
            RsReset!CountCreditMaster = NullToNol(rs!CountCreditMaster)
            RsReset!CountCreditLain = NullToNol(rs!CountCreditLain)
            
            RsReset!Charge = NullToNol(rs!Charge1)
            RsReset!Retur = NullToNol(rs!Retur)
            
            RsReset!IDPengawasZ = IDUser
            RsReset!KodePengawasZ = KodeKasir
            RsReset!NamaPengawasZ = NamaKasir
            RsReset.Update
            PrintResetHartani NullToNol(RsReset!NoID)
          End If
        Else
          If Not NullToBool(RsReset!ResetZ) Then
            RsReset!ResetX = True
            RsReset!SubTotal = NullToNol(rs!SubTotal1)
            RsReset!DiscMember = NullToNol(rs!DiscNota1)
            RsReset!Pembulatan = NullToNol(rs!Pembulatan1)
            RsReset!Total = NullToNol(rs!Total1)
            RsReset!Voucher = NullToNol(rs!Voucher1)
            RsReset!Tunai = NullToNol(rs!Tunai1)
            
            RsReset!Debet = NullToNol(rs!Debet)
            RsReset!Credit = NullToNol(rs!Kredit)
            RsReset!JumlahNota = NullToNol(rs!JumNota)
            RsReset!DebetMandiri = NullToNol(rs!DebetMandiri)
            RsReset!DebetBCA = NullToNol(rs!DebetBCA)
            RsReset!DebetLain = NullToNol(rs!DebetLain)
            RsReset!CreditVisa = NullToNol(rs!CreditVisa)
            RsReset!CreditBCA = NullToNol(rs!CreditBCA)
            RsReset!CreditMaster = NullToNol(rs!CreditMaster)
            RsReset!CreditLain = NullToNol(rs!CreditLain)
            RsReset!CountDebetMandiri = NullToNol(rs!CountDebetMandiri)
            RsReset!CountDebetBCA = NullToNol(rs!CountDebetBCA)
            RsReset!CountDebetLain = NullToNol(rs!CountDebetLain)
            RsReset!CountCreditVisa = NullToNol(rs!CountCreditVisa)
            RsReset!CountCreditBCA = NullToNol(rs!CountCreditBCA)
            RsReset!CountCreditMaster = NullToNol(rs!CountCreditMaster)
            RsReset!CountCreditLain = NullToNol(rs!CountCreditLain)
            RsReset!Charge = NullToNol(rs!Charge1)
            RsReset!Retur = NullToNol(rs!Retur)

            RsReset!IDPengawasX = IDUser
            RsReset!KodePengawasX = KodeKasir
            RsReset!NamaPengawasX = NamaKasir
            RsReset.Update
            PrintResetHartani NullToNol(RsReset!NoID)
          End If
        End If
      Else
      RsReset.AddNew
      If Text1.Text = 2 Then
          If Not NullToBool(RsReset!ResetZ) Then
            RsReset!ResetZ = True
            RsReset!SubTotal = NullToNol(rs!SubTotal1)
            RsReset!DiscMember = NullToNol(rs!DiscNota1)
            RsReset!Pembulatan = NullToNol(rs!Pembulatan1)
            RsReset!Total = NullToNol(rs!Total1)
            RsReset!Voucher = NullToNol(rs!Voucher1)
            RsReset!Tunai = NullToNol(rs!Tunai1)
            
            RsReset!Debet = NullToNol(rs!Debet)
            RsReset!Credit = NullToNol(rs!Kredit)
            RsReset!JumlahNota = NullToNol(rs!JumNota)
            RsReset!DebetMandiri = NullToNol(rs!DebetMandiri)
            RsReset!DebetBCA = NullToNol(rs!DebetBCA)
            RsReset!DebetLain = NullToNol(rs!DebetLain)
            RsReset!CreditVisa = NullToNol(rs!CreditVisa)
            RsReset!CreditBCA = NullToNol(rs!CreditBCA)
            RsReset!CreditMaster = NullToNol(rs!CreditMaster)
            RsReset!CreditLain = NullToNol(rs!CreditLain)
            RsReset!CountDebetMandiri = NullToNol(rs!CountDebetMandiri)
            RsReset!CountDebetBCA = NullToNol(rs!CountDebetBCA)
            RsReset!CountDebetLain = NullToNol(rs!CountDebetLain)
            RsReset!CountCreditVisa = NullToNol(rs!CountCreditVisa)
            RsReset!CountCreditBCA = NullToNol(rs!CountCreditBCA)
            RsReset!CountCreditMaster = NullToNol(rs!CountCreditMaster)
            RsReset!CountCreditLain = NullToNol(rs!CountCreditLain)
            RsReset!Charge = NullToNol(rs!Charge1)
            RsReset!Retur = NullToNol(rs!Retur)
            
            RsReset!IDPengawasZ = IDUser
            RsReset!KodePengawasZ = KodeKasir
            RsReset!NamaPengawasZ = NamaKasir
            RsReset.Update
            PrintResetHartani NullToNol(RsReset!NoID)
          End If
        Else
          If Not NullToBool(RsReset!ResetZ) Then
            RsReset!ResetX = True
            RsReset!SubTotal = NullToNol(rs!SubTotal1)
            RsReset!DiscMember = NullToNol(rs!DiscNota1)
            RsReset!Pembulatan = NullToNol(rs!Pembulatan1)
            RsReset!Total = NullToNol(rs!Total1)
            RsReset!Voucher = NullToNol(rs!Voucher1)
            RsReset!Tunai = NullToNol(rs!Tunai1)
            
            RsReset!Debet = NullToNol(rs!Debet)
            RsReset!Credit = NullToNol(rs!Kredit)
            RsReset!JumlahNota = NullToNol(rs!JumNota)
            RsReset!DebetMandiri = NullToNol(rs!DebetMandiri)
            RsReset!DebetBCA = NullToNol(rs!DebetBCA)
            RsReset!DebetLain = NullToNol(rs!DebetLain)
            RsReset!CreditVisa = NullToNol(rs!CreditVisa)
            RsReset!CreditBCA = NullToNol(rs!CreditBCA)
            RsReset!CreditMaster = NullToNol(rs!CreditMaster)
            RsReset!CreditLain = NullToNol(rs!CreditLain)
            RsReset!CountDebetMandiri = NullToNol(rs!CountDebetMandiri)
            RsReset!CountDebetBCA = NullToNol(rs!CountDebetBCA)
            RsReset!CountDebetLain = NullToNol(rs!CountDebetLain)
            RsReset!CountCreditVisa = NullToNol(rs!CountCreditVisa)
            RsReset!CountCreditBCA = NullToNol(rs!CountCreditBCA)
            RsReset!CountCreditMaster = NullToNol(rs!CountCreditMaster)
            RsReset!CountCreditLain = NullToNol(rs!CountCreditLain)
            RsReset!Charge = NullToNol(rs!Charge1)
            RsReset!Retur = NullToNol(rs!Retur)
            
            RsReset!IDPengawasX = IDUser
            RsReset!KodePengawasX = KodeKasir
            RsReset!NamaPengawasX = NamaKasir
            RsReset.Update
            PrintResetHartani NullToNol(RsReset!NoID)
          End If
        End If
        RsReset.Update
      End If
      
      RsReset.Close
      Set RsReset = Nothing
      dbs.Close
      Set dbs = Nothing
    '==============
End Sub
Private Sub ResendServerOnline()
Dim rstTabel As New ADODB.Recordset
Dim dsTabel As New ADODB.Connection
Dim rstServer As New ADODB.Recordset
Dim m_con As New ADODB.Connection
Dim i As Integer, SQL As String
On Error GoTo Trace
  If isRemcomendedOnline Then
    dsTabel.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DIRPOS & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
    dsTabel.Open
    SQL = "SELECT * FROM MSales WHERE Year(TANGGAL) = " & CInt(Right(Text2.Text, 4)) & " And Month(TANGGAL) = " & CInt(Mid(Text2.Text, 3, 2)) & " And Day(TANGGAL) = " & CInt(Left(Text2.Text, 2))
    rstTabel.CursorLocation = adUseClient
    rstTabel.Open SQL, dsTabel, adOpenDynamic, adLockOptimistic, adCmdText
    If Not (rstTabel.EOF Or rstTabel.BOF) Then
'      rstTabel.MoveFirst
      If m_con.State = adStateOpen Then
        m_con.Close
      End If
      m_con.ConnectionString = Cnstr
      m_con.Open
      
      For i = 0 To rstTabel.RecordCount - 1
        'CekServer
        SQL = "SELECT * FROM MJual WHERE NoIDPOS=" & NullToNol(rstTabel!NoID) & " AND IDPOS=" & CInt(IDPOSDef) & " AND Year(TANGGAL) = " & Right(Text2.Text, 4) & " And Month(TANGGAL) = " & Mid(Text2.Text, 3, 2) & " And Day(TANGGAL) = " & Left(Text2.Text, 2)
'        If NullToNol(rstTabel!NoID) = 230 Then
'          MsgBox "test"
'        End If
        If rstServer.State = adStateOpen Then
          rstServer.Close
        End If
        rstServer.CursorLocation = adUseClient
        rstServer.Open SQL, m_con, adOpenDynamic, adLockOptimistic, adCmdText
        If Not (rstServer.EOF Or rstServer.BOF) Then
'          Dim x As Integer
'          Do While Not rstServer.EOF
'            If Not CBool(NullToNol(rstServer!IsPosted)) Then
'              ExecuteSQL "DELETE FROM MJualD WHERE IDJual=" & NullToNol(rstServer!NoID)
'              ExecuteSQL "DELETE FROM MJual WHERE NoID=" & NullToNol(rstServer!NoID)
'            Else
'              UnPostingStokBarangPenjualan NullToNol(rstServer!NoID)
'              ExecuteSQL "DELETE FROM MJualD WHERE IDJual=" & NullToNol(rstServer!NoID)
'              ExecuteSQL "DELETE FROM MJual WHERE NoID=" & NullToNol(rstServer!NoID)
'            End If
'            rstServer.MoveNext
'          Loop
'          SQL = "SELECT COUNT(NoID) FROM MJual WHERE NoID=" & NullToNol(rstServer!NoID)
'          Dim Jumlah As Long
'          Jumlah = NullToNol(ExecuteSkalarSQL(SQL))
'          If Jumlah >= 1 Then
'          Else
'            KirimKeServer NullToNol(rstTabel!NoID)
'          End If
        Else
        DoEvents
          KirimKeServerBeginTrans NullToNol(rstTabel!NoID)
          DoEvents
        End If
        rstTabel.MoveNext
      Next
      rstServer.Close
      Set rstServer = Nothing
      m_con.Close
      Set m_con = Nothing
    End If
    rstTabel.Close
    Set rstTabel = Nothing
    dsTabel.Close
    Set dsTabel = Nothing
  End If
Trace:
  If Err.Number <> 0 Then
    MsgBox "Info : " & Err.Number & " " & Err.Description, vbCritical, "VPOS"
    Err.Clear
  End If
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

'Public Sub KirimKeServer(ByVal IDSales As Long)
'    Dim ispending As Boolean
'    ispending = False
'    Dim penambahanPointSebelumnya As Long
'    Dim SaldoKSBNotaIni As Long
'    Dim NamaTabelSales As String
'    Dim DiscNotaProsen As Double
'    Dim DiscPersen1 As Double
'    Dim IDWilayah As Long
'    Dim IDSalesAHS As Long
'    Dim i As Integer
'    If isRemcomendedOnline = True Then
'      Dim nmfile As String
'      Dim SQL As String
'      Dim jumrec As Long
'      Dim TANGGAL As String
'      nmfile = App.Path & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
'
'      DataLokal.DatabaseName = nmfile
'      DataLokal.RecordSource = "Select * from MSales where NoID=" & IDSales
'      DataLokal.Refresh
'
'      With DataLokal.Recordset
'        If .EOF Or .BOF Then
'        Else
'        TANGGAL = "convert(datetime,'" & Format(!TANGGAL, "MM/dd/yyyy") & "',101)"
'        'BO pakai AHS
'        NamaTabelSales = "MJual"
'        If NamaTabelSales = "MJual" Then
'                If (!SubTotal) <> 0 Then
'                    DiscNotaProsen = !DiscNota * 100 / !SubTotal
'                Else
'                    DiscNotaProsen = 0
'                End If
'                IDSalesAHS = NullToNol(ExecuteSkalarSQL("Select Max(NoID) as Hasil From MJual")) + 1
'                IDWilayah = NullToNol(ExecuteSkalarSQL("SELECT IDWilayah FROM MGudang WHERE NoID=" & IDGudangDef))
''                sql = "INSERT INTO " & NamaTabelSales & "(NoID,Kode,Jam,Tanggal,Shift," & _
''                "IsPOS,IDTransaksiKassa,IDKassa," & _
''                "SubTotal,DiskonNotaRp,DiskonNotaProsen,Biaya,Total,Bayar,IDUserEntry,IDUser, KodeSalesman) VALUES(" & _
''                IDSalesAHS & ",'" & !kode & "',convert(datetime,'" & Format(!TANGGAL, "HH:nn:ss") & "',101) ,convert(datetime,'" & _
''                Format(!TANGGAL, "MM/dd/yyyy") & "',101) ," & NamaShift & "," & _
''                "1," & !NoId & "," & IDPOSDef & "," & _
''                FixKoma(!SubTotal) & "," & FixKoma(!DiscNota) & _
''                "," & FixKoma(DiscNotaProsen) & "," & FixKoma(!UangMuka - !Hargatotal) & "," & FixKoma(!Hargatotal) & "," & _
''                FixKoma(!UangMuka) & "," & IDUser & "," & IDUser & ",'" & NamaKasir & "' )"
'
'                SQL = "INSERT INTO MJual (IDGudang,IDWilayah,IsPOS,NoID,Kode,KodeReff,Tanggal,TanggalStock,JatuhTempo,"
'                SQL = SQL & " IDCustomer,TanggalSJ,NoSJ,SubTotal,DiskonNotaProsen,DiskonNotaRp,DiskonNotaTotal,"
'                SQL = SQL & " Biaya, Total, Bayar, Sisa,IDAdmin,IDPacking,Shift,NamaKasir,Pembulatan,IDBank,NoAcc) VALUES (" & vbCrLf
'                SQL = SQL & IDGudangDef & "," & IDWilayah & ","
'                SQL = SQL & 1 & ","
'                SQL = SQL & IDSalesAHS & ","
'                SQL = SQL & "'" & Replace(!kode & "/" & Format(IDPOSDef, "00") & Format(Now, "yyMM"), "'", "''") & "',"
'                               SQL = SQL & "'" & Replace(!kode, "'", "''") & "',"
'                SQL = SQL & "'" & Format(!TANGGAL, "yyyy-MM-dd HH:nn:ss") & "',"
'                SQL = SQL & "'" & Format(!TANGGAL, "yyyy-MM-dd HH:nn:ss") & "',"
'                SQL = SQL & "'" & Format(!TANGGAL, "yyyy-MM-dd HH:nn:ss") & "',"
'                SQL = SQL & IDMember & ","
'                SQL = SQL & "'" & Format(!TANGGAL, "yyyy-MM-dd HH:nn:ss") & "',"
'                SQL = SQL & "'',"
'                SQL = SQL & FixKoma(!SubTotal) & ","
'                SQL = SQL & FixKoma(DiscNotaProsen) & ","
'                SQL = SQL & FixKoma(!DiscNota) & ","
'                SQL = SQL & FixKoma(!DiscNota) & ","
'                SQL = SQL & FixKoma(0) & ","
'                SQL = SQL & FixKoma(!Hargatotal) & ","
'                SQL = SQL & FixKoma(!UangMuka) & ","
'                SQL = SQL & FixKoma(!Hargatotal - !UangMuka) & ","
'                SQL = SQL & IDUser & ","
'                SQL = SQL & -1 & "," & Replace(NamaShift, "'", "''") & ", '" & Replace(NamaKasir, "'", "''") & "'," & FixKoma(NullToNol(!Pembulatan)) & "," & NullToNol(!IDBank) & ",'" & Replace(NullToStr(!NoAcc), "'", "''") & "')"
'       End If
'          '999999: cara lama masih ada kemungkinan record sales dengan detil kosong kekirim yang berakibat fatal diserver
'          If BoolToInt(!ispending) = 1 Then
'            ispending = True
'          Else
'            ispending = False
'          End If
''            lbStatus.Caption = "Status : " & ExecuteSQL(sql)
''          End If
'          DoEvents
'        End If
'        End With
'
'        If ispending = False Then
'        Dim IDJualD As Long
'              DataLokal.RecordSource = "Select * FROM MSalesd WHERE IDSales=" & IDSales & " ORDER BY NoID"
'              DataLokal.Refresh
'              With DataLokal.Recordset
'                  If .EOF Or .BOF Then
'                  Else
'                  '999999: Pindah disini
'
''                    lbStatus.Caption = "Status : " & ExecuteSQL(SQL)
'                    .MoveFirst
'                    i = 1
'                    Do While Not .EOF
'                         If NamaTabelSales = "MJual" Then
'                          IDJualD = NullToNol(ExecuteSkalarSQL("Select Max(NoID) as Hasil From MJualD")) + 1
'                          If !HargaBruto <> 0 Then
'                             DiscPersen1 = (!DiscRp + !DiscInternRp) * 100 / !HargaBruto
'                          Else
'                             DiscPersen1 = 0
'                          End If
'
'                          SQL = "INSERT INTO MJualD (NoID,IDJual,IDPackingD,NoUrut,Tgl,Jam,IDBarang,IDSatuan,Qty,QtyPcs,Harga,HargaPcs,CTN,DiscPersen1,DiscPersen2,DiscPersen3,Disc1,Disc2,Disc3,Jumlah,Catatan,IDGudang,Konversi) VALUES ("
'                          SQL = SQL & IDJualD & ","
'                          SQL = SQL & IDSalesAHS & ","
'                          SQL = SQL & -1 & ","
'                          SQL = SQL & i & ","
'                          SQL = SQL & "GetDate(),"
'                          SQL = SQL & "GetDate(),"
'                          SQL = SQL & !IdInventor & ","
'                          SQL = SQL & !idSatuan & ","
'                          SQL = SQL & FixKoma(!QTY) & ","
'                          SQL = SQL & FixKoma(!QTY * !Konversi) & ","
'                          SQL = SQL & FixKoma(!HargaBruto) & ","
'                          SQL = SQL & FixKoma(!Jumlah / IIf(!QTY = 0, 1, !QTY) / IIf(!Konversi = 0, 1, !Konversi)) & ","
'                          SQL = SQL & FixKoma(NullToNol(ExecuteSkalarSQL("Select " & !QTY & "/MBarang.Ctn_Pcs*" & !Konversi & " AS Hasil FROM MBarang WHERE NoID=" & !IdInventor))) & ","
'                          SQL = SQL & FixKoma(!DiscProsen1) & ","
'                          SQL = SQL & FixKoma(!DiscProsen2) & ","
'                          SQL = SQL & FixKoma(!DiscProsen3) & ","
'                          SQL = SQL & FixKoma(!Disc1) & ","
'                          SQL = SQL & FixKoma(!Disc2) & ","
'                          SQL = SQL & FixKoma(!Disc3) & ","
'                          SQL = SQL & FixKoma(!Jumlah) & ","
'                          If !Transaksi = "PLU" Then
'                            SQL = SQL & "'Penjualan POS',"
'                          Else
'                            SQL = SQL & "'Returan POS',"
'                          End If
'                          SQL = SQL & IDGudangDef & ","
'                          SQL = SQL & !Konversi & ""
'                          SQL = SQL & ")"
'
''                          sql = "INSERT INTO MJualD(NoID,IDJual,IDGudang,IDBarang," & _
''                             "Qty,Harga,HargaPokok,DiscPersen1,Disc1) VALUES(" & _
''                              IDJualD & "," & IDSalesAHS & "," & IDGudangDef & "," & !IDInventor & "," & _
''                              FixKoma(!QTY) & "," & FixKoma(!HargaBruto) & "," & _
''                              FixKoma(!HargaPokok) & "," & FixKoma(DiscPersen1) & "," & FixKoma(!DiscRp + !DiscInternRp) & ")"
'                        End If
'                    ExecuteSQL (SQL)
'                    i = i + 1
'                    DoEvents
'                    .MoveNext
'                    Loop
'                    DoEvents
'                  End If
'                End With
'                DataLokal.Database.Execute "Update MSales Set IsUpload=1, IsSelesai=1 where NoID=" & IDSales
'          End If
'    If IDMember > 0 Then
''        penambahanPointSebelumnya = ExecuteSkalarSQLMEMBER("select Sum(SaldoNotaIni) as hasil " & _
''        "from MCustomerPoint where IDCustomer=" & IDMember & " AND Tanggal=convert(datetime,'" & Format(Date, "MM/dd/yyyy") & "',101)")
''
''        SaldoKSBNotaIni = (BelanjaPoin - Voucher) - ((BelanjaPoin - Voucher + penambahanPointSebelumnya) \ 100000) * 100000
''
''        ExecuteSQLMEMBER "Insert Into MCustomerPoint(IDCustomer,KodeCustomer,Kode,Tanggal,Kassa,Bruto,Netto,Debet,IsKueBasah,PenambahanSebelumnya,SaldoNotaIni) Values(" & _
''                IDMember & ",'" & KodeMember & "','" & Format(IDSales, "0000000") & "'," & TANGGAL & ",'" & NamaMesin & "'," & FixKoma(Total) & "," & _
''                FixKoma(BelanjaPoin - Voucher) & "," & (BelanjaPoin - Voucher + penambahanPointSebelumnya) \ 100000 & "," & IIf(NamaMesin = "12" Or NamaMesin = "13", 1, 0) & "," & FixKoma(penambahanPointSebelumnya) & "," & FixKoma(SaldoKSBNotaIni) & ")"
'    End If
'   End If
'End Sub
'Public Sub ResendServerOnline()
'  Dim ispending As Boolean
'  Dim IDSales As Long
'  ispending = False
'    If isRemcomendedOnline = True Then
'      Dim kassa As String
'      Dim nmfile As String
'      Dim SQL As String
'      Dim jumrec As Long
'      Dim TANGGAL As String
'      nmfile = App.Path & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
'
'      DataLokal.DatabaseName = nmfile
'       DataLokalDTL.DatabaseName = nmfile
'      DataLokal.RecordSource = "Select * from MSales " & _
'       "WHERE (IsUpload=0 AND ((MSales.Shift)=" & NamaShiftReset & ") AND  (Day(MSales.Tanggal)=" & DayReset & " and Month(MSales.Tanggal)=" & MonthReset & " and Year(MSales.Tanggal)=" & YearReset & "))"
'
'      DataLokal.Refresh
'
'      With DataLokal.Recordset
'        If .EOF Or .BOF Then
'        Else
'        .MoveFirst
'        Do While Not .EOF
'          IDSales = !NoId
''          TANGGAL = "convert(datetime,'" & Format(!TANGGAL, "MM/dd/yyyy") & "',101)"
''
''          SQL = "INSERT INTO MSALESPOS(IDPos,NoIDSales,Kode,Tanggal,Tgl,Jam,TotalPajak,TotalDiscount," & _
''          "SubTotal,DiscNota,HargaTotal,IDPayment,UangMuka,IDUser,IsSend,Bank,ISPending," & _
''          "Shift,Voucher,IDBank,IDCustomer,IDVoucher,KodeMember,IDMember,BarangKSB,SisaKSB) VALUES(" & _
''           IDPOSDef & "," & !NoId & ",'" & !Kode & "',convert(datetime,'" & _
''            Format(!TANGGAL, "MM/dd/yyyy hh:nn:ss") & "',101) ,convert(datetime,'" & _
''            Format(!TANGGAL, "MM/dd/yyyy") & "',101) ,convert(datetime,'" & _
''            Format(!TANGGAL, "hh:nn:ss") & "',14) ," & FixKoma(!TotalPajak) & "," & _
''            FixKoma(!TotalDiscount) & "," & FixKoma(!SubTotal) & "," & FixKoma(!DiscNota) & "," & _
''            FixKoma(!Hargatotal) & "," & !IDPayment & "," & FixKoma(!UangMuka) & "," & _
''            !IDUser & "," & BoolToInt(!IsSend) & "," & FixKoma(!Bank) & "," & _
''            BoolToInt(!ispending) & "," & !Shift & "," & FixKoma(!Voucher) & "," & _
''            NullToNol(!IDBank) & "," & NullToNol(!idcustomer) & "," & _
''            NullToNol(!IDVoucher) & ",'" & NullToStr(!KodeMember) & "'," & NullToNol(!IDMember) & "," & FixKoma(!BarangKSB) & "," & FixKoma(!SisaKSB) & ")"
''            '999999: cara lama masih ada kemungkinan record sales dengan detil kosong kekirim yang berakibat fatal diserver
''            If BoolToInt(!ispending) = 1 Then
''              ispending = True
''            Else
''              ispending = False
''            End If
''  '            lbStatus.Caption = "Status : " & ExecuteSQL(sql)
''  '          End If
''              If ispending = False Then
''                'Delete header ada kemungkinan keisi tapi tidak lengkap
''                ExecuteSQL "Delete From MSalesPos where Shift=" & NamaShiftReset & " AND Tgl=" & TANGGAL & " AND NoIDSales=" & IDSales & " AND IDPos=" & IDPOSDef
''
''                ExecuteSQL "Delete From MSalesPosD where Shift=" & NamaShiftReset & " AND Tanggal=" & TANGGAL & " AND NoIDSalesPos=" & IDSales & " AND IDPos=" & IDPOSDef
''               DataLokalDTL.RecordSource = "Select * from MSalesd where IDSales=" & IDSales
''                DataLokalDTL.Refresh
''                  With DataLokalDTL.Recordset
''                    If .EOF Or .BOF Then
''                    Else
''                    '999999: Pindah disini
''                       ExecuteSQL (SQL) 'lbStatus.Caption = "Status : " &
''                      .MoveFirst
''                      Do While Not .EOF
''
''                      SQL = "INSERT INTO MSALESPOSD(NoIDSalesD,NoIDSalesPos,IDPos,IDGudang,IDInvsat," & _
''                      "Qty,Harga,IDSatuan,Konversi,Transaksi,HargaPokok,HargaBruto,DiscRp," & _
''                      "DiscProsen,IsDiscSupplier,Tanggal,IsMember) VALUES(" & _
''                        !NoId & "," & !IDSales & "," & IDPOSDef & "," & IDGudangDef & "," & !IdInventor & "," & _
''                        FixKoma(!QTY) & "," & FixKoma(!harga) & "," & !idSatuan & "," & FixKoma(!Konversi) & ",'" & _
''                        !Transaksi & "'," & FixKoma(!HargaPokok) & "," & FixKoma(!HargaBruto) & "," & FixKoma(!DiscRp) & "," & _
''                        FixKoma(!DiscProsen) & "," & 0 & "," & TANGGAL & "," & BoolToInt(!IsMember) & ")"
''                      ExecuteSQL (SQL)
''                        DoEvents
''                      .MoveNext
''                      Loop
''                      DoEvents
''                    End If
''                  End With
'                .Edit
'                !IsUpload = 1
'                .Update
'                .Bookmark = .LastModified
''            End If
'            DoEvents
''            If NullToNol(!IDMember) > 0 Then 'And NullToNol(!BarangKSB) >= 100000
''                ExecuteSQL "Insert Into MCustomerPoint(IDCustomer,Kode,Tanggal,Kassa,Netto,Debet) Values(" & _
''                NullToNol(!IDMember) & ",'" & NullToStr(!Kode) & "'," & TANGGAL & ",''," & FixKoma(NullToNol(!BarangKSB)) & "," & NullToNol(!BarangKSB) \ 100000 & ")"
''           End If
'            .MoveNext
'        Loop
'        End If
'        End With
'
'
'   End If
'
'End Sub
Sub PRINTRESET()
'     lbl(1) = namaMesinReset & " - " & namakasirReset
'     lbl(2) = Format(rs!JumlahNota, "###,##0")
'     lbl(3) = Format(rs!JumlahSubTotal + rs!DIskonBrg + rs!JumDIskon, "###,###,##0") '+ rs!DIskonBrg + rs!JumDIskon
'     lbl(4) = Format(rs!JumDIskon + rs!DIskonBrg, "###,###,##0")
'     lbl(5) = Format(rs!JumlahTotal + rs!JumDIskon + rs!DIskonBrg, "###,###,##0")
'     lbl(6) = Format(rs!JumUangMuka, "###,###,##0")
'     lbl(7) = Format(rs!Jumbank, "###,###,##0")
'     lbl(8) = Format(rs!JumVoucher, "###,###,##0")
'     lbl(9) = Format(rs!JumlahTotal - rs!JumlahSubTotal + rs!JumDIskon + rs!Bulat, "###,###,##0")
'     TunaiPajak = ((PersenLap / 100) * rs!JumUangMuka) - (((PersenLap / 100) * rs!JumUangMuka) Mod 100)
Dim dbsr As Database
Dim rst As Recordset
On Error GoTo Trace
  Set dbsr = OpenDatabase(DirDatabase & "\TempDB" & Format(CDate(Mid(Text2.Text, 5, 4) & "-" & Mid(Text2.Text, 3, 2) & "-" & Mid(Text2.Text, 1, 2)), "_yyyyMM") & ".mdb")
  CrystalReport1.ReportFileName = App.path & "\Report\STRUCKRESET.rpt"

  Set rst = dbsr.OpenRecordset("SELECT COUNT(MSales.NoID) AS JumlahNota, " & _
  " SUM(MSales.Subtotal) as TSubtotal, SUM(Msales.Hutang) AS TPiutang, " & _
  " SUM(Msales.Pembulatan+Msales.DiscNota) AS TPembulatan, 0 AS Retur, SUM(MSales.UangMuka) as TTunai, SUM(HargaTotal) AS TTotal, SUM(BANK) AS TDebet " & _
  " From MSales WHERE Subtotal<>0 AND MONTH(Tanggal)=" & Mid(Text2.Text, 3, 2) & " AND YEAR(Tanggal)=" & Right(Text2.Text, 4) & " AND DAY(Tanggal)=" & Left(Text2.Text, 2))
  If rst.BOF Or rst.EOF Then
  Else
    CrystalReport1.Formulas(1) = "JumlahNota=" & rst!JumlahNota
    CrystalReport1.Formulas(2) = "Subtotal=" & rst!TSubtotal
    CrystalReport1.Formulas(3) = "Pembulatan=" & rst!TPembulatan
    CrystalReport1.Formulas(4) = "Tunai=" & rst!TTunai - rst!TDebet
    CrystalReport1.Formulas(5) = "Total=" & rst!TTotal
    CrystalReport1.Formulas(6) = "Piutang=" & rst!TPiutang
    CrystalReport1.Formulas(7) = "Debet=" & rst!TDebet
  End If
  
'  CrystalReport1.Formulas(7) = "NamaMesinReset='" & Format(Text3.Text, "00") & "'"
  CrystalReport1.Formulas(8) = "NamaPerusahaan='" & Trim(NamaToko) & "'"
  CrystalReport1.Formulas(9) = "AlamatPerusahaan='" & Trim(Replace(Judulstruk, vbCrLf, "'+ Chr(13) +'")) & "'"
  CrystalReport1.WindowShowPrintSetupBtn = True
  CrystalReport1.WindowShowRefreshBtn = True
  CrystalReport1.WindowState = crptMaximized
'  If TipeCetakan = Preview Then
'    CrystalReport1.Destination = crptToWindow
'  Else
    CrystalReport1.Destination = crptToPrinter
'  End If
  CrystalReport1.Action = 1
  
'  Set rst = Nothing
'  dbsr.Close
'  Set dbsr = Nothing
Trace:
  If Err.Number <> 0 Then
'    MsgBox Err.Number & " " & Err.Description, vbCritical, App.Title
    Err.Clear
  End If
End Sub

Sub PrintResetHartani(ByVal NoID As Long)
Dim dbsr As Database
Dim rst As Recordset
On Error GoTo Trace
  Set dbsr = OpenDatabase(DirDatabase & "\TempDB" & Format(CDate(Mid(Text2.Text, 5, 4) & "-" & Mid(Text2.Text, 3, 2) & "-" & Mid(Text2.Text, 1, 2)), "_yyyyMM") & ".mdb")
  CrystalReport1.ReportFileName = App.path & "\Report\STRUCKRESET.rpt"
  CrystalReport1.DataFiles(0) = App.path & "\Database\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
  'cdate(2011,12,31)
  CrystalReport1.Formulas(1) = "NoID=" & NoID
  CrystalReport1.Formulas(2) = "NamaMesinReset='" & Text3.Text & "'"
  CrystalReport1.Formulas(3) = "Reset='" & IIf(Text1.Text = 2, "Z", "X") & "'"
  CrystalReport1.Formulas(4) = "NamaPerusahaan='" & Trim(NamaToko) & "'"
  CrystalReport1.Formulas(5) = "AlamatPerusahaan='" & Trim(Replace(Judulstruk, vbCrLf, "'+ Chr(13) +'")) & "'"
  CrystalReport1.WindowShowPrintSetupBtn = True
  CrystalReport1.WindowShowRefreshBtn = True
  CrystalReport1.WindowState = crptMaximized
  CrystalReport1.Destination = crptToPrinter
  CrystalReport1.Action = 1
Trace:
  If Err.Number <> 0 Then
    'MsgBox Err.Number & " " & Err.Description, vbCritical, App.Title
    Err.Clear
  End If
End Sub


