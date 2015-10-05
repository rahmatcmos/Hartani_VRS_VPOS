VERSION 5.00
Begin VB.Form frmEntriBarang 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entri Barang"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSaldoStok 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   3195
      Width           =   2745
   End
   Begin VB.TextBox txtCari 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2280
      TabIndex        =   18
      Top             =   45
      Width           =   4815
   End
   Begin VB.TextBox PvCurrency3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2280
      TabIndex        =   15
      Top             =   3765
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.TextBox PVCurrency2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2730
      Width           =   2745
   End
   Begin VB.TextBox PVCurrency1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "&Tutup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6060
      TabIndex        =   17
      Top             =   3630
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "&Simpan"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4770
      TabIndex        =   16
      Top             =   4440
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   9
      Top             =   4125
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   7
      Top             =   1830
      Width           =   4830
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2280
      TabIndex        =   5
      Top             =   1380
      Width           =   4845
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   3
      Top             =   930
      Width           =   4830
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2280
      TabIndex        =   1
      Top             =   480
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo Stok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   9
      Left            =   180
      TabIndex        =   21
      Top             =   3225
      Width           =   1965
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Barcode/Kode"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   8
      Left            =   180
      TabIndex        =   19
      Top             =   45
      Width           =   1965
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Harga Pokok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   7
      Left            =   180
      TabIndex        =   14
      Top             =   3795
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Harga Jual"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   180
      TabIndex        =   12
      Top             =   2760
      Width           =   1965
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Satuan Terkecil"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   180
      TabIndex        =   8
      Top             =   4185
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Isi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   180
      TabIndex        =   10
      Top             =   2295
      Width           =   1965
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Satuan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   180
      TabIndex        =   6
      Top             =   1842
      Width           =   1965
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Barcode"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   180
      TabIndex        =   4
      Top             =   1388
      Width           =   1965
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Barang"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Top             =   934
      Width           =   1965
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Barang"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   480
      Width           =   1965
   End
End
Attribute VB_Name = "frmEntriBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NoIdbrg As Long
Dim isNewBrg As Boolean
Dim isSimpanBrg As Boolean
Dim kodeBrg As String
Dim NamaBrg As String
Dim idSatuan As Long
'Dim dbs As Database
'Dim rs As Recordset

'Sub RecordHistory()
'Dim dbsh As Database
'Dim rsh As Recordset
'Set dbsh = OpenDatabase(DirDatabase & "\HistInv.mdb", , False)
'Set rsh = dbsh.OpenRecordset("HisHarga")
'  If Not isNewBrg Then
'    rsh.AddNew
'    rsh!TANGGAL = Date
'    rsh!Jam = Time
'    rsh!NoID = rs!NoID
'    rsh!IdInventor = rs!IdInventor
'    rsh!OldKode = rs!kode
'    rsh!OldBarcode = rs!Barcode
'    rsh!OldNama = rs!Nama
'    rsh!OldIDSatuan = rs!idSatuan
'    rsh!OldKodeSat = rs!KodeSat
'    rsh!OldKonversi = rs!Konversi
'    rsh!OldHargaJual = rs!HargaJual
'    rsh!OldHargaPokok = rs!HargaPokok
'    rsh!kode = Text1.Text
'    rsh!Barcode = Text3.Text
'    rsh!Nama = Text2.Text
'    rsh!idSatuan = rs!idSatuan
'    rsh!KodeSat = Text4.Text
'    rsh!Konversi = CCur(PVCurrency1.Text)
'    rsh!HargaJual = CCur(PVCurrency2.Text)
'    rsh!HargaPokok = CCur(PvCurrency3.Text)
'    rsh.Update
'  Else
'    rsh.AddNew
'    rsh!TANGGAL = Date
'    rsh!Jam = Time
'    rsh!NoID = NoIdbrg
'    rsh!IdInventor = NoIdbrg
'    rsh!OldKode = Text1.Text
'    rsh!OldBarcode = Text3.Text
'    rsh!OldNama = Text2.Text
'    rsh!OldIDSatuan = 1
'    rsh!OldKodeSat = Text4.Text
'    rsh!OldKonversi = CCur(PVCurrency1.Text)
'    rsh!OldHargaJual = CCur(PVCurrency2.Text)
'    rsh!OldHargaPokok = CCur(PvCurrency3.Text)
'    rsh!kode = Text1.Text
'    rsh!Barcode = Text3.Text
'    rsh!Nama = Text2.Text
'    rsh!idSatuan = 1
'    rsh!KodeSat = Text4.Text
'    rsh!Konversi = CCur(PVCurrency1.Text)
'    rsh!HargaJual = CCur(PVCurrency2.Text)
'    rsh!HargaPokok = CCur(PvCurrency3.Text)
'    rsh.Update
'  End If
'  rsh.Close
'Set rsh = Nothing
'dbsh.Close
'Set dbsh = Nothing
'End Sub
'Sub Simpan()
'  If Not EntriValid Then Exit Sub
'  CekSatuan
'  If Not isNewBrg Then
'    RecordHistory
'    rs.Edit
'  Else
'    NoIdbrg = GetNewIDMaster("Tinv")
'    RecordHistory
'    rs.AddNew
'    rs!NoID = NoIdbrg
'    rs!IdInventor = NoIdbrg
'  End If
'  rs!kode = Text1.Text
'  rs!Nama = Text2.Text
'  rs!Barcode = Text3.Text
'  rs!KodeSat = Text4.Text
''  rs!SatTerkecil = Text5.Text
'  rs!Konversi = IIf(IsNumeric(PVCurrency1.Text), PVCurrency1.Text, 0)
'  rs!HargaJual = IIf(IsNumeric(PVCurrency2.Text), PVCurrency2.Text, 0)
'  rs!HargaPokok = IIf(IsNumeric(PvCurrency3.Text), PvCurrency3.Text, 0)
'  rs!idSatuan = idSatuan
'  rs.Update
'
'  kodeBrg = Text1.Text
'  NamaBrg = Text2.Text
'  isSimpanBrg = True
'  dbs.Close
'  Set dbs = Nothing
'  Unload Me
'End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim hasil As String
  hasil = Trim(SendByCode(KeyCode))
  KeyCode = 0
  Select Case hasil
  Case "ENT"
'      Simpan
  Case "DN"
      Command2.SetFocus
  Case "RGT"
      Command2.SetFocus
  Case "UP"
      PvCurrency3.SetFocus
  Case "LFT"
      PvCurrency3.SetFocus
  End Select
End Sub



Sub Tampil(ByVal isNew As Boolean, ByRef NoID As Long, ByRef kode As String, ByRef Nama As String, ByRef issimpan As Boolean)
  isNewBrg = isNew
  NoIdbrg = NoID
  Me.Show 1
  NoID = NoIdbrg
  issimpan = isSimpanBrg
  kode = kodeBrg
  Nama = NamaBrg
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim hasil As String
hasil = Trim(SendByCode(KeyCode))
KeyCode = 0
Select Case hasil
Case "ENT"
'    dbs.Close
'    Set dbs = Nothing
    Unload Me
Case "DN"
    Text1.SetFocus
Case "RGT"
    Text1.SetFocus
Case "UP"
    Command1.SetFocus
Case "LFT"
    Command1.SetFocus
End Select
End Sub

Private Sub Form_Activate()
txtCari.SetFocus
End Sub

Private Sub Form_Load()
isSimpanBrg = False

'Set dbs = OpenDatabase(DirDatabase & "\dbMaster.mdb")
'If isNewBrg Then
'  Set rs = dbs.OpenRecordset("Select * From Tinv Where NoID=0")
'  Text1.Text = ""
'  Text2.Text = ""
'  Text3.Text = ""
'  Text4.Text = ""
''  Text5.Text = ""
'  PVCurrency1.Text = "1"
'  PVCurrency2.Text = "0"
'  PvCurrency3.Text = "0"
'Else
'  Set rs = dbs.OpenRecordset("Select * From Tinv Where NoID=" & NoIdbrg)
'  Text1.Text = NullToStr(rs!kode)
'  Text2.Text = NullToStr(rs!Nama)
'  Text3.Text = NullToStr(rs!Barcode)
'  Text4.Text = NullToStr(rs!KodeSat)
''  Text5.Text = NullToStr(rs!SatTerkecil)
'  PVCurrency1.Text = Trim(CStr(NullToNol(rs!Konversi)))
'  PVCurrency2.Text = Trim(CStr(NullToNol(rs!HargaJual)))
'  PvCurrency3.Text = Trim(CStr(NullToNol(rs!HargaPokok)))
'End If
End Sub
Public Function GetNewIDMaster(ByVal nmTabel As String) As Long
  Dim dbs As Database
  Dim rs As Recordset
  Set dbs = OpenDatabase(DirDatabase & "\dbMaster.mdb")
  Set rs = dbs.OpenRecordset("SELECT MAX(NoID) as ID FROM " & nmTabel)
  If rs.EOF And rs.BOF Then
    GetNewIDMaster = 1
  Else
    If IsNull(rs!ID) Then
      GetNewIDMaster = 1
    Else
      GetNewIDMaster = rs!ID + 1
    End If
  End If
  Set rs = Nothing
  dbs.Close
  Set dbs = Nothing
End Function

Function EntriValid() As Boolean
  EntriValid = True
  If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Then ' Or Text5.Text = "" Then
    MsgBox "Kode, Nama  dan satuan barang tidak boleh kosong", vbCritical
    EntriValid = False
  ElseIf PVCurrency1.Text = "" Or PVCurrency1.Text = "" Or PVCurrency1.Text = "" Then
    MsgBox "Harga harus diisi", vbCritical
    EntriValid = False
  ElseIf IsNumeric(PVCurrency1.Text) Then
    If Val(PVCurrency1.Text) < 1 Then
      MsgBox "Konversi satuan harus lebih besar 0", vbCritical
      EntriValid = False
    End If
  End If
End Function

Sub CekSatuan()
'  Dim rsSatuan As Recordset
'  Set rsSatuan = dbs.OpenRecordset("MSatuan", dbOpenDynaset)
'  If rsSatuan.EOF And rsSatuan.BOF Then
'    idSatuan = GetNewIDMaster("Msatuan")
'    rsSatuan.AddNew
'    rsSatuan!NoId = idSatuan
'    rsSatuan!Nama = Text4.Text
'    rsSatuan!Konversi = Val(PVCurrency1.Text)
'    rsSatuan!SatuanDasar = Text5.Text
'    rsSatuan.Update
'  Else
'    rsSatuan.FindFirst "Nama='" & Replace(Text4.Text, "'", "''") & "' AND SatuanDasar='" & Replace(Text5.Text, "'", "''") & "'"
'    If rsSatuan.NoMatch Then
'      idSatuan = GetNewIDMaster("Msatuan")
'      rsSatuan.AddNew
'      rsSatuan!NoId = idSatuan
'      rsSatuan!Nama = Text4.Text
'      rsSatuan!Konversi = Val(PVCurrency1.Text)
'      rsSatuan!SatuanDasar = Text5.Text
'      rsSatuan.Update
'    Else
'      idSatuan = rsSatuan!NoId
'      rsSatuan.Edit
'        rsSatuan!Konversi = Val(PVCurrency1.Text)
'      rsSatuan.Update
'    End If
'  End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
'    dbs.Close
'    Set dbs = Nothing
End Sub

Private Sub PVCurrency1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim hasil As String
hasil = Trim(SendByCode(KeyCode))
Select Case hasil
Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "'", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "-", "+", "=", "|", "\", ".", ":", ";", "?"
  PVCurrency1.Text = PVCurrency1.Text & hasil
  PVCurrency1.SelStart = Len(PVCurrency1.Text)
 Case "SPC"
  PVCurrency1.Text = PVCurrency1.Text & " "
  PVCurrency1.SelStart = Len(PVCurrency1.Text)
Case "BKS"
  If Len(PVCurrency1.Text) > 0 Then PVCurrency1.Text = Left(PVCurrency1.Text, Len(PVCurrency1.Text) - 1)
   PVCurrency1.SelStart = Len(PVCurrency1.Text)
Case "CLR"
  PVCurrency1.Text = ""
Case "ENT", "DN"
  PVCurrency2.SetFocus
Case "UP"
  Text5.SetFocus
End Select
End Sub

Private Sub PVCurrency1_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub PVCurrency2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim hasil As String
hasil = Trim(SendByCode(KeyCode))
Select Case hasil
Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "'", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "-", "+", "=", "|", "\", ".", ":", ";", "?"
  PVCurrency2.Text = PVCurrency2.Text & hasil
  PVCurrency2.SelStart = Len(PVCurrency2.Text)
 Case "SPC"
  PVCurrency2.Text = PVCurrency2.Text & " "
  PVCurrency2.SelStart = Len(PVCurrency2.Text)
Case "BKS"
  If Len(PVCurrency2.Text) > 0 Then PVCurrency2.Text = Left(PVCurrency2.Text, Len(PVCurrency2.Text) - 1)
   PVCurrency2.SelStart = Len(PVCurrency2.Text)
Case "CLR"
        PVCurrency2.Text = ""
Case "ENT", "DN"
  PvCurrency3.SetFocus
Case "UP"
  PVCurrency1.SetFocus
End Select
End Sub

Private Sub PVCurrency2_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub PvCurrency3_KeyDown(KeyCode As Integer, Shift As Integer)
Dim hasil As String
hasil = Trim(SendByCode(KeyCode))
Select Case hasil
Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "'", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "-", "+", "=", "|", "\", ".", ":", ";", "?"
  PvCurrency3.Text = PvCurrency3.Text & hasil
  PvCurrency3.SelStart = Len(PvCurrency3.Text)
 Case "SPC"
  PvCurrency3.Text = PvCurrency3.Text & " "
  PvCurrency3.SelStart = Len(PvCurrency3.Text)
Case "BKS"
  If Len(PvCurrency3.Text) > 0 Then PvCurrency3.Text = Left(PvCurrency3.Text, Len(PvCurrency3.Text) - 1)
   PvCurrency3.SelStart = Len(PvCurrency3.Text)
Case "CLR"
        PvCurrency3.Text = ""
Case "ENT", "DN"
  Command1.SetFocus
Case "UP"
  PVCurrency2.SetFocus
End Select
End Sub

Private Sub PvCurrency3_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim hasil As String
hasil = Trim(SendByCode(KeyCode))
Select Case hasil
Case "0", "00", "1", "2", "3", "4", "5", "6", "7", "8", "9", "*", "`", "!", "@", "#", "$", "%", "^", "&", "<", "(", ")", "-", "+", "{", "}", "[", "]", "/", "?", ".", ","
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
Case "ENT", "DN"
  Text2.SetFocus
Case "UP"
  Command2.SetFocus
End Select
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim hasil As String
hasil = Trim(SendByCode(KeyCode))
Select Case hasil
Case "0", "00", "1", "2", "3", "4", "5", "6", "7", "8", "9", "*", "`", "!", "@", "#", "$", "%", "^", "&", "<", "(", ")", "-", "+", "{", "}", "[", "]", "/", "?", ".", ","
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
Case "ENT", "DN"
  Text3.SetFocus
Case "UP"
  Text1.SetFocus
End Select
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
Dim hasil As String
hasil = Trim(SendByCode(KeyCode))
Select Case hasil
Case "0", "00", "1", "2", "3", "4", "5", "6", "7", "8", "9", "*", "`", "!", "@", "#", "$", "%", "^", "&", "<", "(", ")", "-", "+", "{", "}", "[", "]", "/", "?", ".", ","
  Text3.Text = Text3.Text & hasil
  Text3.SelStart = Len(Text3.Text)
 Case "SPC"
  Text3.Text = Text3.Text & " "
  Text3.SelStart = Len(Text3.Text)
Case "BKS"
  If Len(Text3.Text) > 0 Then Text3.Text = Left(Text3.Text, Len(Text3.Text) - 1)
   Text3.SelStart = Len(Text3.Text)
Case "CLR"
        Text3.Text = ""
Case "ENT", "DN"
  Text4.SetFocus
Case "UP"
  Text2.SetFocus
End Select
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
Dim hasil As String
hasil = Trim(SendByCode(KeyCode))
Select Case hasil
Case "0", "00", "1", "2", "3", "4", "5", "6", "7", "8", "9", "*", "`", "!", "@", "#", "$", "%", "^", "&", "<", "(", ")", "-", "+", "{", "}", "[", "]", "/", "?", ".", ","
  Text4.Text = Text4.Text & hasil
  Text4.SelStart = Len(Text4.Text)
 Case "SPC"
  Text4.Text = Text4.Text & " "
  Text4.SelStart = Len(Text4.Text)
Case "BKS"
  If Len(Text4.Text) > 0 Then Text4.Text = Left(Text4.Text, Len(Text4.Text) - 1)
   Text4.SelStart = Len(Text4.Text)
Case "CLR"
        Text4.Text = ""
Case "ENT", "DN"
  Text5.SetFocus
Case "UP"
  Text3.SetFocus
End Select
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
Dim hasil As String
hasil = Trim(SendByCode(KeyCode))
Select Case hasil
Case "0", "00", "1", "2", "3", "4", "5", "6", "7", "8", "9", "*", "`", "!", "@", "#", "$", "%", "^", "&", "<", "(", ")", "-", "+", "{", "}", "[", "]", "/", "?", ".", ","
  Text5.Text = Text5.Text & hasil
  Text5.SelStart = Len(Text5.Text)
 Case "SPC"
  Text5.Text = Text5.Text & " "
  Text5.SelStart = Len(Text5.Text)
Case "BKS"
  If Len(Text5.Text) > 0 Then Text5.Text = Left(Text5.Text, Len(Text5.Text) - 1)
   Text5.SelStart = Len(Text5.Text)
Case "CLR"
        Text5.Text = ""
Case "ENT", "DN"
  PVCurrency1.SetFocus
Case "UP"
  Text4.SetFocus
End Select
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub txtCari_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
    Dim dbs As Database
    Dim rs As Recordset
    Set dbs = OpenDatabase(DirDatabase & "\dbMaster.mdb")
    
      Set rs = dbs.OpenRecordset("Select * From Tinv Where Barcode='" & txtCari.Text & "'")
      If rs.EOF Or rs.BOF Then
          rs.Close
          Set rs = dbs.OpenRecordset("Select * From Tinv Where Kode='" & txtCari.Text & "'")
          If rs.EOF Or rs.BOF Then
            MsgBox "Barang dengan Barcode/Kode tersebut tidak ada tidak ada!", vbExclamation
          Else
              Text1.Text = NullToStr(rs!kode)
              Text2.Text = NullToStr(rs!Nama)
              Text3.Text = NullToStr(rs!Barcode)
              Text4.Text = NullToStr(rs!KodeSat)
            '  Text5.Text = NullToStr(rs!SatTerkecil)
              PVCurrency1.Text = Trim(CStr(NullToNol(rs!Konversi)))
              PVCurrency2.Text = Trim(CStr(NullToNol(rs!HargaJual)))
              PvCurrency3.Text = Trim(CStr(NullToNol(rs!HargaPokok)))
            If isRemcomendedOnline Then
              txtSaldoStok.Text = NullToNol(ExecuteSkalarSQL("Select Sum(ISNULL(Konversi,1)*(ISNULL(QtyMasuk,0)-ISNULL(QtyKeluar,0))) as Saldo from MKartuStok where MKartuStok.IDBarang=" & NullToNol(rs!IdInventor)))
            End If
          End If
        rs.Close
      Else
          Text1.Text = NullToStr(rs!kode)
          Text2.Text = NullToStr(rs!Nama)
          Text3.Text = NullToStr(rs!Barcode)
          Text4.Text = NullToStr(rs!KodeSat)
        '  Text5.Text = NullToStr(rs!SatTerkecil)
          PVCurrency1.Text = Trim(CStr(NullToNol(rs!Konversi)))
          PVCurrency2.Text = Trim(CStr(NullToNol(rs!HargaJual)))
          PvCurrency3.Text = Trim(CStr(NullToNol(rs!HargaPokok)))
         
        If isRemcomendedOnline Then
              txtSaldoStok.Text = NullToNol(ExecuteSkalarSQL("Select Sum(ISNULL(Konversi,1)*(ISNULL(QtyMasuk,0)-ISNULL(QtyKeluar,0))) as Saldo from MKartuStok where MKartuStok.IDBarang=" & NullToNol(rs!IdInventor)))
        End If
 rs.Close
      End If
    Set rs = Nothing
    dbs.Close
    Set dbs = Nothing
  ElseIf KeyCode = 27 Then
    If txtCari.Text <> "" Then
      txtCari.Text = ""
      Text1.Text = ""
      Text2.Text = ""
      Text3.Text = ""
      Text4.Text = ""
    '  Text5.Text = NullToStr(rs!SatTerkecil)
      PVCurrency1.Text = ""
      PVCurrency2.Text = ""
      PvCurrency3.Text = ""
txtSaldoStok.Text = ""
    Else
      Unload Me
    End If
  End If
End Sub
