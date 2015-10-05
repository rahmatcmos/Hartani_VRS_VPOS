VERSION 5.00
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dxctrls.dll"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLookUpTKP 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11850
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2220
      TabIndex        =   0
      Top             =   180
      Width           =   9555
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Height          =   6015
      Left            =   60
      OleObjectBlob   =   "frmLookUpTKP.frx":0000
      TabIndex        =   3
      Top             =   600
      Width           =   11715
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   30
      Top             =   345
      Visible         =   0   'False
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Tambah <SUB TOTAL>_____Edit <CASH>_____Hapus <HOLD>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   150
      TabIndex        =   2
      Top             =   6780
      Visible         =   0   'False
      Width           =   9345
   End
   Begin VB.Label lbCari 
      BackStyle       =   0  'Transparent
      Caption         =   "CARI KODE / NAMA"
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
      Height          =   405
      Left            =   90
      TabIndex        =   1
      Top             =   210
      Width           =   2895
   End
   Begin CONTROLSLibCtl.dxBackground dxBackground1 
      Height          =   7155
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11970
      _Version        =   65536
      _cx             =   21114
      _cy             =   12621
      StartColor      =   15997586
      EndColor        =   16708297
      ColorFillStyle  =   0
      BackgroundStyle =   0
      DrawPictureStyle=   1
      AnimationInterval=   1000
   End
End
Attribute VB_Name = "frmLookUpTKP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim isHasilKonversi As Boolean
'Dim kodeBrg As String
Public NamaField As String
Dim ambil As Boolean
Dim NamaBrg As String
Dim NoIdbrg As Long
Dim issimpan As Boolean
Dim NoIDSales As Long
Dim SQLQuery As String

Dim cn As New ADODB.Connection
'Dim com As New ADODB.Command
Dim rst As New ADODB.Recordset

Private Sub Form_Activate()
If isSupervisor Then
  Label1.Visible = True
Else
  Label1.Visible = False
End If
End Sub

Sub View()
On Error GoTo pesan
SQLQuery = "SELECT MTukarPoin.NoID, MAlamat.Kode AS NoMember, MAlamat.Nama AS NamaMember, MAlamat.Alamat AS AlamatMember, MTukarPoin.JumlahPoin, MTukarPoin.Kredit AS PoinYgDiambil, MTukarPoin.Saldo AS SisaPoin" & vbCrLf & _
           " From MTukarPoin" & vbCrLf & _
           " LEFT JOIN MAlamat ON MAlamat.NoID=MTukarPoin.IDMember" & vbCrLf & _
           " WHERE MTukarPoin.IDKassa=" & IDPOSDef & " AND MTukarPoin.Tanggal>='" & Format(Now, "yyyy-MM-dd") & "' AND MTukarPoin.Tanggal<'" & Format(DateAdd("d", 1, Now), "yyyy-MM-dd") & "'"

  Adodc1.RecordSource = SQLQuery
  Adodc1.Refresh
  If Not (Adodc1.Recordset.EOF Or Adodc1.Recordset.BOF) Then
    Data1.Recordset.MoveFirst
    Set SelBks = TDBGrid1.SelBookmarks
    While SelBks.Count <> 0
        SelBks.Remove 0
    Wend
    TDBGrid1.SelBookmarks.Add TDBGrid1.Bookmark
  End If

pesan:
  If Err.Number <> 0 Then
    MsgBox "Error : " & Err.Number & ", " & Err.Description, vbCritical
    Err.Clear
  End If
End Sub

Private Sub Form_GotFocus()

End Sub

Private Sub Form_Load()
Dim SQL As String
isHasilKonversi = False
ambil = False

Adodc1.ConnectionString = Cnstr
View
'If isOnline = False Then
'  Data1.DatabaseName = DirDatabase & "\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
'Else
'  Data1.DatabaseName = DirDbServer & "\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
'End If
'SQL = "SELECT [NoID], [Tanggal], [Kode], [KodeMember], [SubTotal]-[DiscNota]-[Pembulatan] AS Total FROM MSales ORDER BY Kode, KodeMember"
'Data1.RecordSource = SQL
'Data1.Refresh
''Data1.Recordset.Index = NamaField
End Sub

'Sub View()
'Dim SelBks
'''mENGGUNAKAN tABEL
''Data1.Recordset.Index = NamaField
''Data1.Recordset.Seek "<=", Text1.Text
''If Data1.Recordset.NoMatch Then
''  Data1.Recordset.MoveFirst
''Else
''  If Text1.Text <> TDBGrid1.Columns(NamaField).Text Then
''    Data1.Recordset.MoveNext
''    If Data1.Recordset.EOF Or (UCase(Left(TDBGrid1.Columns(NamaField).Text, Len(Text1.Text))) <> UCase(Text1.Text)) Then
''      Data1.Recordset.MovePrevious
''    End If
''  End If
''End If
'
''RECORDSET
'' Data1.Recordset.FindNext NamaField & " LIKE '" & Replace(Text1.Text, "'", "''") & "%'"
'' If Data1.Recordset.NoMatch Then Data1.Recordset.MoveFirst
'' Dim SelBks
''  Set SelBks = TDBGrid1.SelBookmarks
''  While SelBks.Count <> 0
''      SelBks.Remove 0
''  Wend
''  TDBGrid1.SelBookmarks.Add TDBGrid1.Bookmark
'  Dim SQL As String
'  If UCase(NamaField) = "Kode" Then
'    SQL = "SELECT [NoID], [Tanggal], [Kode], [KodeMember], [SubTotal]-[DiscNota]-[Pembulatan] AS Total FROM MSales WHERE (UCASE(KodeMember) LIKE '*" & Replace(UCase(Text1.Text), "'", "''") & "*' OR UCASE(Kode) LIKE '*" & Replace(UCase(Text1.Text), "'", "''") & "*') ORDER BY Kode, KodeMember"
'  Else
'    SQL = "SELECT [NoID], [Tanggal], [Kode], [KodeMember], [SubTotal]-[DiscNota]-[Pembulatan] AS Total FROM MSales WHERE (UCASE(KodeMember) LIKE '*" & Replace(UCase(Text1.Text), "'", "''") & "*' OR UCASE(Kode) LIKE '*" & Replace(UCase(Text1.Text), "'", "''") & "*') ORDER BY Kode, KodeMember"
'  End If
'  Data1.RecordSource = SQL
'  Data1.Refresh
'  If Not (Data1.Recordset.EOF Or Data1.Recordset.BOF) Then
'    Data1.Recordset.MoveFirst
'    Set SelBks = TDBGrid1.SelBookmarks
'        While SelBks.Count <> 0
'            SelBks.Remove 0
'        Wend
'        TDBGrid1.SelBookmarks.Add TDBGrid1.Bookmark
'  End If
'End Sub

Private Sub Form_Unload(Cancel As Integer)
Data1.Database.Close
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim idBrg As Long
Dim isSimpanBrg As Boolean
Dim SelBks, SQL As String
Dim hasil As String
hasil = Trim(SendByCode(KeyCode))
Select Case hasil
Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "00", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "'", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "-", "+", "=", "|", "\", ".", ":", ";", "?", "[", "]"
  Text1.Text = Text1.Text & hasil
  Text1.SelStart = Len(Text1.Text)
  DoEvents
  View
 Case "SPC"
  Text1.Text = Text1.Text & " "
  Text1.SelStart = Len(Text1.Text)
  DoEvents
  View
Case "BKS"
  If Len(Text1.Text) > 0 Then Text1.Text = Left(Text1.Text, Len(Text1.Text) - 1)
  Text1.SelStart = Len(Text1.Text)
  DoEvents
  View
Case "CLR"
  Text1.Text = ""
  View
Case "DN"
        Data1.Recordset.MoveNext
        If Data1.Recordset.EOF Then
          Data1.Recordset.MovePrevious
        End If
        Set SelBks = TDBGrid1.SelBookmarks
        While SelBks.Count <> 0
            SelBks.Remove 0
        Wend
        TDBGrid1.SelBookmarks.Add TDBGrid1.Bookmark
Case "UP"
        Data1.Recordset.MovePrevious
        If Data1.Recordset.BOF Then
          Data1.Recordset.MoveNext
        End If
        Set SelBks = TDBGrid1.SelBookmarks
        While SelBks.Count <> 0
            SelBks.Remove 0
        Wend
        TDBGrid1.SelBookmarks.Add TDBGrid1.Bookmark
Case "ENT"
        ambil = True
        
        NoIDSales = TDBGrid1.Columns("NoID")
        Unload Me
Case "ESC"
      ambil = False
      Unload Me
Case "PLU"
        ambil = True
        'kodeBrg = TDBGrid1.Columns("Kode")
        Unload Me
Case "STL"
  If isSupervisor Then 'SUB TOTAL add new
      frmEntriBarang.Tampil True, NoIdbrg, "", NamaBrg, isSimpanBrg
      If isSimpanBrg Then
        Data1.Refresh
        If NamaField = "Kode" Then
          Text1.Text = ""
        Else
          Text1.Text = NamaBrg
        End If
        View
      End If
  End If
Case "CSH"
    If isSupervisor Then 'cash 'edit
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then Exit Sub
      NoIdbrg = Data1.Recordset!NoID
     frmEntriBarang.Tampil False, NoIdbrg, "", NamaBrg, isSimpanBrg
      If isSimpanBrg Then
        Data1.Refresh
        If NamaField = "Kode" Then
          Text1.Text = ""
        Else
          Text1.Text = NamaBrg
        End If
        View
      End If
  End If
End Select
'Dim SelBks
'Dim idBrg As Long
'Dim isSimpanBrg As Boolean
'  If isHasilKonversi And (KeyCode = 8 Or KeyCode = 9 Or KeyCode = 13 Or KeyCode = 27) Then
'    KeyCode = 0
'    Exit Sub
'  End If
'  If Not (isHasilKonversi) Then
'    KeyKode = KeyCode
'    KeyCode = 0
'    isRun = False
'  Else
'      If KeyCode = 40 Then
'        KeyCode = 0
'        Data1.Recordset.MoveNext
'        If Data1.Recordset.EOF Then
'          Data1.Recordset.MovePrevious
'        End If
'        Set SelBks = TDBGrid1.SelBookmarks
'        While SelBks.Count <> 0
'            SelBks.Remove 0
'        Wend
'        TDBGrid1.SelBookmarks.Add TDBGrid1.Bookmark
'      ElseIf KeyCode = 38 Then
'        KeyCode = 0
'        Data1.Recordset.MovePrevious
'        If Data1.Recordset.BOF Then
'          Data1.Recordset.MoveNext
'        End If
'        Set SelBks = TDBGrid1.SelBookmarks
'        While SelBks.Count <> 0
'            SelBks.Remove 0
'        Wend
'        TDBGrid1.SelBookmarks.Add TDBGrid1.Bookmark
'      ElseIf KeyCode = 36 Then 'HOME CLEAR
'        KeyCode = 0
'        Text1.Text = ""
'        Data1.Recordset.MoveFirst
'        Set SelBks = TDBGrid1.SelBookmarks
'        While SelBks.Count <> 0
'            SelBks.Remove 0
'        Wend
'        TDBGrid1.SelBookmarks.Add TDBGrid1.Bookmark
'      ElseIf KeyCode = 35 Then 'END PLU
'        KeyCode = 0
'        ambil = True
'        kodeBrg = TDBGrid1.Columns("Kode")
'        Unload Me
'    ElseIf KeyCode = 33 And isSupervisor Then 'PGUP ADD NEW
'      frmEntriBarang.Tampil True, NoIdbrg, kodeBrg, NamaBrg, isSimpanBrg
'      If isSimpanBrg Then
'        Data1.Refresh
'        If NamaField = "Kode" Then
'          Text1.Text = kodeBrg
'        Else
'          Text1.Text = NamaBrg
'        End If
'      End If
'    ElseIf KeyCode = 34 And isSupervisor Then 'PGDN EDIT
'    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then Exit Sub
'      NoIdbrg = Data1.Recordset!NoId
'     frmEntriBarang.Tampil False, NoIdbrg, kodeBrg, NamaBrg, isSimpanBrg
'      If isSimpanBrg Then
'        Data1.Refresh
'        If NamaField = "Kode" Then
'          Text1.Text = kodeBrg
'        Else
'          Text1.Text = NamaBrg
'        End If
'      End If
'    ElseIf KeyCode = 46 And isSupervisor Then 'DEL HAPUS
'    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then Exit Sub
'      If MsgBox("Hapus Barang " & TDBGrid1.Columns("Nama") & "?", vbQuestion & vbYesNo, "Hapus Data") = vbYes Then
'      Text1.Text = Text1.Text
'        Data1.Recordset.Delete
'        Data1.Recordset.MoveNext
'        If Data1.Recordset.EOF Then
'          Data1.Recordset.MovePrevious
'        End If
'
'        Set SelBks = TDBGrid1.SelBookmarks
'        While SelBks.Count <> 0
'            SelBks.Remove 0
'        Wend
'        TDBGrid1.SelBookmarks.Add TDBGrid1.Bookmark
'
'      End If
'    End If 'Keycode
'    End If 'isHasilKonversi
End Sub
Public Sub Tampil(ByRef isambil As Boolean, NoID As Long)
Me.Show 1
isambil = ambil
NoID = NoIDSales
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
'  If Not (isHasilKonversi) Then
'    KeyAscii = 0
'    isRun = True
'    isHasilKonversi = True
'    SendKeys SendByCode(KeyKode), True
'    isHasilKonversi = False
'  Else
'      If KeyAscii = 27 Then
'        KeyAscii = 0
'        Unload Me
'      ElseIf KeyAscii = 13 Then
'        Text1.Text = ""
'
''    ElseIf KeyAscii = 59 Then ';=plu
''      KeyAscii = 0
''      ambil = True
''      KodeBrg = TDBGrid1.Columns("Kode")
''      Unload Me
'    End If
'  End If
End Sub

'Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
'If Not isRun Then
'        isRun = True
'        isHasilKonversi = True
'        SendKeys (SendByCode(KeyKode)), True
'        isHasilKonversi = False
'    End If
'End Sub
