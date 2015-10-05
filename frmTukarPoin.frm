VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmTukarPoin 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tukar Poin"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8715
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   8715
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   3600
      Top             =   3450
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      Left            =   2880
      TabIndex        =   14
      Top             =   45
      Width           =   4935
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
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2730
      Width           =   1575
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
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1830
      Width           =   1575
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
      Left            =   6660
      TabIndex        =   13
      Top             =   3630
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "&Simpan"
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
      Left            =   5430
      TabIndex        =   12
      Top             =   3630
      Width           =   1125
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
      Left            =   2880
      MaxLength       =   30
      TabIndex        =   7
      Top             =   2280
      Width           =   4950
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
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1380
      Width           =   1575
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
      Left            =   2880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   3
      Top             =   930
      Width           =   4950
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
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   4935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Barcode/Kode Member"
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
      Height          =   270
      Index           =   8
      Left            =   180
      TabIndex        =   15
      Top             =   45
      Width           =   2565
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sisa"
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
      TabIndex        =   10
      Top             =   2760
      Width           =   1965
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Catatan"
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
      TabIndex        =   8
      Top             =   2295
      Width           =   1965
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Poin yang ditukar"
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
      Caption         =   "Saldo Poin"
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
      Caption         =   "Alamat"
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
      Caption         =   "Nama"
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
Attribute VB_Name = "frmTukarPoin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NoIDMember As Long
Dim isNewBrg As Boolean
Dim isSimpanBrg As Boolean
Dim kodeBrg As String
Dim NamaBrg As String
Dim idSatuan As Long
'Dim dbs As Database
'Dim rs As Recordset

Private Sub Command1_Click()
 If EntriValid Then
    Simpan
 End If
End Sub

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
  Dim Hasil As String
  Hasil = Trim(SendByCode(KeyCode))
  KeyCode = 0
  Select Case Hasil
  Case "ENT"
  If EntriValid Then
      Simpan
    End If
  Case "DN"
      Command2.SetFocus
  Case "RGT"
      Command2.SetFocus
  Case "UP"
      Text4.SetFocus
  Case "LFT"
      Text4.SetFocus
  End Select
End Sub



 

Private Sub Command2_Click()
   Unload Me
End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Hasil As String
Hasil = Trim(SendByCode(KeyCode))
KeyCode = 0
Select Case Hasil
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

End Sub

Function EntriValid() As Boolean
  EntriValid = True
  If Text1.Text = "" Or Text2.Text = "" Then   ' Or Text5.Text = "" Then
    MsgBox "Kode, Nama   tidak boleh kosong", vbCritical
    EntriValid = False
  ElseIf PVCurrency1.Text = "" Then
    MsgBox "Poin yang ditukar harus diisi", vbCritical
    EntriValid = False
  ElseIf IsNumeric(PVCurrency1.Text) Then
    If Val(PVCurrency1.Text) < 1 Then
      MsgBox "Poin yang ditukar harus lebih besar 0", vbCritical
      EntriValid = False
    ElseIf IsNumeric(Text3.Text) Then
        If CCur(Text3.Text) < CCur(PVCurrency1.Text) Then
            MsgBox "Poin yang ditukar harus lebih kecil sama dengan Saldo", vbCritical
            EntriValid = False
        End If
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
Dim Hasil As String
Hasil = Trim(SendByCode(KeyCode))
Select Case Hasil
Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "'", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "-", "+", "=", "|", "\", ".", ":", ";", "?"
  PVCurrency1.Text = PVCurrency1.Text & Hasil
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
If IsNumeric(PVCurrency1.Text) Then
  PVCurrency2.Text = CCur(Text3.Text) - CCur(PVCurrency1.Text)
  Text4.SetFocus
  End If
Case "UP"
  txtCari.SetFocus
End Select
End Sub

Private Sub PVCurrency1_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub PVCurrency2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Hasil As String
Hasil = Trim(SendByCode(KeyCode))
Select Case Hasil
Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "'", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "-", "+", "=", "|", "\", ".", ":", ";", "?"
  PVCurrency2.Text = PVCurrency2.Text & Hasil
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
  Command1.SetFocus
Case "UP"
  PVCurrency1.SetFocus
End Select
End Sub

Private Sub PVCurrency2_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Hasil As String
Hasil = Trim(SendByCode(KeyCode))
Select Case Hasil
Case "0", "00", "1", "2", "3", "4", "5", "6", "7", "8", "9", "*", "`", "!", "@", "#", "$", "%", "^", "&", "<", "(", ")", "-", "+", "{", "}", "[", "]", "/", "?", ".", ","
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
Dim Hasil As String
Hasil = Trim(SendByCode(KeyCode))
Select Case Hasil
Case "0", "00", "1", "2", "3", "4", "5", "6", "7", "8", "9", "*", "`", "!", "@", "#", "$", "%", "^", "&", "<", "(", ")", "-", "+", "{", "}", "[", "]", "/", "?", ".", ","
  Text2.Text = Text2.Text & Hasil
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
Dim Hasil As String
Hasil = Trim(SendByCode(KeyCode))
Select Case Hasil
Case "0", "00", "1", "2", "3", "4", "5", "6", "7", "8", "9", "*", "`", "!", "@", "#", "$", "%", "^", "&", "<", "(", ")", "-", "+", "{", "}", "[", "]", "/", "?", ".", ","
  Text3.Text = Text3.Text & Hasil
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
Dim Hasil As String
Hasil = Trim(SendByCode(KeyCode))
Select Case Hasil
Case "0", "00", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "'", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "-", "+", "=", "|", "\", ".", ":", ";", "?"
  Text4.Text = Text4.Text & Hasil
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
  Command1.SetFocus
Case "UP"
  PVCurrency1.SetFocus
End Select
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub
  
Private Sub txtCari_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
     'On Error GoTo 0
    bacaSettingServer
    Dim isOnline As Boolean
    Dim sqlcon As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    'Dim rsPoin As New ADODB.Recordset
    Set sqlcon = New ADODB.Connection
    sqlcon.ConnectionString = Cnstr
    sqlcon.Open
    Set rs = sqlcon.Execute("Select * from Malamat where iscustomer=1 and kode='" & Replace(txtCari.Text, "'", "''") & "'")
    If rs.EOF And rs.BOF Then
    Else
          Text1.Text = NullToStr(rs!Nama) 'nama
          Text2.Text = NullToStr(rs!Alamat) 'alamat
          
          NoIDMember = NullToNol(rs!NoID) 'alamat
          'Set rsPoin = sqlcon.Execute("SELECT isnull((select SUM(isnull(mjual.nilaiPoin,0)) FROM MJUAL where mjual.idcustomer=" & NoIDMember & "),0)-Isnull((SELECT SUM(isnull(MTukarPoin.Kredit,0)) from  mtukarpoin  WHERE IDMEMBER=" & NoIDMember & "),0) AS SaldoPoin")
          'Set rsPoin = sqlcon.Execute("SELECT vSaldoPoin.SaldoPoin FROM vSaldoPoin WHERE vSaldoPoin.IDCustomer=" & NoIDMember)
          Text3.Text = GetNilaiPoinMember(NoIDMember)
          PVCurrency1.SetFocus
'          Text4.Text = NullToStr(rs!KodeSat) 'catatan
'          PVCurrency1.Text = Trim(CStr(NullToNol(rs!Konversi))) 'poin yang ditukar
'          PVCurrency2.Text = Trim(CStr(NullToNol(rs!HargaJual))) 'sisa
    End If
    rs.Close
    Set rs = Nothing
    sqlcon.Close
    Set sqlcon = Nothing
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
    Else
      Unload Me
    End If
  End If
End Sub
Sub Simpan()
    On Error GoTo pesan
    Command1.Enabled = False
    DoEvents
    bacaSettingServer
    Dim NoIDTukarPoin As Long
    Dim isOnline As Boolean
    Dim sqlcon As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim rsPoin As New ADODB.Recordset
    PVCurrency2.Text = CCur(Text3.Text) - CCur(PVCurrency1.Text)
 
    Set sqlcon = New ADODB.Connection
    sqlcon.ConnectionString = Cnstr
    sqlcon.Open
        Set rs = sqlcon.Execute("Select max(NoID) as IDMax from MTukarPoin")
        If rs.EOF And rs.BOF Then
            NoIDTukarPoin = 1
        Else
            NoIDTukarPoin = NullToNol(rs!IDMax) + 1
        End If
    sqlcon.Execute ("Insert Into MTukarPoin(NoID,Tanggal,Jam,IDMember,NoMember,IDKassa,IDkasir,JumlahPoin,Kredit,Saldo,Keterangan) " & _
                    "Values(" & NoIDTukarPoin & ",'" & Format(Date, "yyyy-MM-dd") & "',getdate()," & NoIDMember & ",'" & Replace(Text1.Text, "'", "''") & "'," & _
                    IDPOSDef & "," & IDUser & "," & FixKoma(Text3.Text) & "," & FixKoma(PVCurrency1.Text) & "," & FixKoma(PVCurrency2.Text) & ",'" & _
                    Replace(Text4.Text, "'", "''") & "')")
    rs.Close
    Set rs = Nothing
    sqlcon.Close
    Set sqlcon = Nothing
      CetakStruck NoIDTukarPoin, ""
      If MsgBox("Siap Cetak Reprint?", vbQuestion + vbYesNo + vbDefaultButton2, "VPOS") = vbYes Then
        CetakStruck NoIDTukarPoin, "COPY"
      End If
      Unload Me
    Exit Sub
pesan:
    MsgBox "Ada kesalahan : " & Err.Description & "(" & Err.Number & ")", vbExclamation
      Command1.Enabled = True
    DoEvents
End Sub

Private Sub CetakStruck(ByVal NoID As Long, ByVal Reprint As String)
'openDrawer

If TipeCetakan = None Then Exit Sub
If TipeCetakan = Optional_ Then
  If MsgBox("Mau Cetak Struk Tukar Poin?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
End If
Dim i As Integer
'On Error GoTo Err

'dbstmp.Execute "Delete * From MCetakSales"
'dbstmp.Execute "Insert into MCetakSales(NoID) Values(" & NoID & ")"
DoEvents
CrystalReport1.Reset
' If NullToNol(rst!TipeHargaJual) = 0 Then
      CrystalReport1.ReportFileName = App.path & "\Report\TukarPoin.rpt"
    'ElseIf NullToNol(rst!TipeHargaJual) = 2 Then
     ' CrystalReport1.ReportFileName = App.Path & "\Report\STRUCK1.rpt"
    'Else
     ' CrystalReport1.ReportFileName = App.Path & "\Report\STRUCK.rpt"
    'End If
    
    'CrystalReport1.re .DataFiles(2) = App.path & "\Database\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
'    CrystalReport1.DataFiles(2) = App.path & "\Database\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
 
'    CrystalReport1.SelectionFormula = "{MSales.NoID}=" & NoID
    CrystalReport1.Formulas(0) = "NoID=" & NoID
    CrystalReport1.Formulas(1) = "NamaKasir='" & NamaKasir & "'"
    CrystalReport1.Formulas(2) = "NamaPerusahaan='" & Trim(NamaToko) & "'"
    CrystalReport1.Formulas(3) = "AlamatPerusahaan='" & Trim(Replace(Judulstruk, vbCrLf, "'+ Chr(13) +'")) & "'"
    CrystalReport1.Formulas(4) = "Kassa='" & Format(IDPOSDef, "##") & "'"
    CrystalReport1.Formulas(5) = "NoMember='" & Trim(txtCari.Text) & "'"
    CrystalReport1.Formulas(6) = "NamaMember='" & Trim(Text1.Text) & "'"
    CrystalReport1.Formulas(7) = "AlamatMember='" & Trim(Text2.Text) & "'"
    CrystalReport1.Formulas(8) = "Poin='" & Trim(Text3.Text) & "'"
    CrystalReport1.Formulas(9) = "TukarPoin='" & Trim(PVCurrency1.Text) & "'"
    CrystalReport1.Formulas(10) = "SisaPoin='" & Trim(PVCurrency2.Text) & "'"
    CrystalReport1.Formulas(11) = "Catatan='" & Trim(Text4.Text) & "'"
    CrystalReport1.Formulas(12) = "Reprint='" & Trim(Reprint) & "'"
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowState = crptMaximized
    If TipeCetakan = Preview_ Then
      CrystalReport1.Destination = crptToWindow
    Else
      CrystalReport1.Destination = crptToPrinter
      CrystalReport1.ProgressDialog = False
    End If
    CrystalReport1.Action = 1
    
    
  Exit Sub
Err:
If Err.Number <> 0 Then
  MsgBox "Error : " & Err.Number & "-" & Err.Description, vbCritical, App.Title
  Err.Clear
End If
End Sub

