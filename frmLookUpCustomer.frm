VERSION 5.00
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dxctrls.dll"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmCustomer 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
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
   ScaleWidth      =   9600
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
      Left            =   1650
      TabIndex        =   0
      Top             =   180
      Width           =   7875
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Bindings        =   "frmLookUpCustomer.frx":0000
      Height          =   6015
      Left            =   60
      OleObjectBlob   =   "frmLookUpCustomer.frx":0014
      TabIndex        =   3
      Top             =   600
      Width           =   9465
   End
   Begin VB.Data Data1 
      Connect         =   "Access"
      DatabaseName    =   "D:\JOB\KassaWin\Database\DbMaster.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2190
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TInv"
      Top             =   2520
      Visible         =   0   'False
      Width           =   2055
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
      Width           =   9345
   End
   Begin VB.Label lbCari 
      BackStyle       =   0  'Transparent
      Caption         =   "CARI NAMA"
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
Attribute VB_Name = "frmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim isHasilKonversi As Boolean
Dim kodeBrg As String
Dim DiscountMarketing As Double
Public NamaField As String
Dim ambil As Boolean
Dim NamaBrg As String
Dim NoIdbrg As Long
Dim issimpan As Boolean
Dim SQL As String
Dim isAmbilDiscountMarketing As Boolean
Private Sub Form_Activate()
If isSupervisor Then
  Label1.Visible = True
Else
  Label1.Visible = False
End If
If NamaField = "Kode" Then
Text1.Text = TerakhirCariKodeCust
Else
Text1.Text = TerakhirCariNamaCust
End If
View
End Sub

Private Sub Form_Load()
isHasilKonversi = False
ambil = False
Text1.Text = ""

If isOnline = False Then
  Data1.DatabaseName = DirDatabase & "\dbMaster.mdb"
Else
  Data1.DatabaseName = DirDbServer & "\dbMaster.mdb"
End If

SQL = "SELECT * FROM MCustomer ORDER BY " & NamaField
Data1.RecordSource = SQL
Data1.Refresh

'Data1.RecordSource = "MCustomer"
'Data1.Refresh
'Data1.Recordset.Index = NamaField
End Sub

Sub View()
Dim SelBks
''mENGGUNAKAN tABEL
'Data1.Recordset.Index = NamaField
'Data1.Recordset.Seek "<=", Text1.Text
'If Data1.Recordset.NoMatch Then
'  Data1.Recordset.MoveFirst
'Else
'  If Text1.Text <> TDBGrid1.Columns(NamaField).Text Then
'    Data1.Recordset.MoveNext
'    If Data1.Recordset.EOF Or (UCase(Left(TDBGrid1.Columns(NamaField).Text, Len(Text1.Text))) <> UCase(Text1.Text)) Then
'      Data1.Recordset.MovePrevious
'    End If
'  End If
'End If
''RECORDSET
' 'Data1.Recordset.FindFirst NamaField & " LIKE '" & Replace(Text1.Text, "'", "''") & "%'"
' 'If Data1.Recordset.NoMatch Then Data1.Recordset.MoveFirst
' Dim SelBks
'  Set SelBks = TDBGrid1.SelBookmarks
'  While SelBks.Count <> 0
'      SelBks.Remove 0
'  Wend
'  TDBGrid1.SelBookmarks.Add TDBGrid1.Bookmark
  Dim SQL As String
  If UCase(NamaField) = "NAMA" Then
    SQL = "SELECT * FROM MCustomer WHERE UCASE(NAMA) LIKE '*" & Replace(UCase(Text1.Text), "'", "''") & "*'  ORDER BY " & NamaField
  Else
    SQL = "SELECT * FROM MCustomer WHERE UCASE(KODE) LIKE '*" & Replace(UCase(Text1.Text), "'", "''") & "*'  ORDER BY " & NamaField
  End If
  Data1.RecordSource = SQL
  Data1.Refresh
  If Not (Data1.Recordset.EOF Or Data1.Recordset.BOF) Then
    Data1.Recordset.MoveFirst
    Set SelBks = TDBGrid1.SelBookmarks
        While SelBks.Count <> 0
            SelBks.Remove 0
        Wend
        TDBGrid1.SelBookmarks.Add TDBGrid1.Bookmark
  End If
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
Data1.Database.Close
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim idBrg As Long
Dim isSimpanBrg As Boolean
Dim SelBks
Dim Hasil As String
Hasil = Trim(SendByCode(KeyCode))
Select Case Hasil
Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "00", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "'", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "-", "+", "=", "|", "\", ".", ":", ";", "?", "[", "]"
  Text1.Text = Text1.Text & Hasil
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
        Data1.Recordset.MoveFirst
        Set SelBks = TDBGrid1.SelBookmarks
        While SelBks.Count <> 0
            SelBks.Remove 0
        Wend
        TDBGrid1.SelBookmarks.Add TDBGrid1.Bookmark

Case "DN"
    If Data1.Recordset.RecordCount >= 1 Then
        Data1.Recordset.MoveNext
        If Data1.Recordset.EOF Then
          Data1.Recordset.MovePrevious
        End If
        Set SelBks = TDBGrid1.SelBookmarks
        While SelBks.Count <> 0
            SelBks.Remove 0
        Wend
        TDBGrid1.SelBookmarks.Add TDBGrid1.Bookmark
    End If
Case "UP"
    If Data1.Recordset.RecordCount >= 1 Then
        Data1.Recordset.MovePrevious
        If Data1.Recordset.BOF Then
          Data1.Recordset.MoveNext
        End If
        Set SelBks = TDBGrid1.SelBookmarks
        While SelBks.Count <> 0
            SelBks.Remove 0
        Wend
        TDBGrid1.SelBookmarks.Add TDBGrid1.Bookmark
    End If
Case "ENT"
        TerakhirCariKodeCust = NullToStr(Data1.Recordset.Fields("Kode"))
        TerakhirCariNamaCust = NullToStr(Data1.Recordset.Fields("Nama"))
        
        ambil = True
        kodeBrg = TDBGrid1.Columns("Kode")
        If isAmbilDiscountMarketing Then
        DiscountMarketing = Data1.Recordset.Fields("DiscountMarketing")
        End If
        Unload Me
Case "ESC"
      ambil = False
      Unload Me

Case "PLU"
        ambil = True
        kodeBrg = TDBGrid1.Columns("Kode")
        Unload Me
Case "STL"
  If isSupervisor Then 'SUB TOTAL add new
      frmEntriBarang.Tampil True, NoIdbrg, kodeBrg, NamaBrg, isSimpanBrg
      If isSimpanBrg Then
        Data1.Refresh
        If NamaField = "Kode" Then
          Text1.Text = kodeBrg
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
     frmEntriBarang.Tampil False, NoIdbrg, kodeBrg, NamaBrg, isSimpanBrg
      If isSimpanBrg Then
        Data1.Refresh
        If NamaField = "Kode" Then
          Text1.Text = kodeBrg
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
Public Sub Tampil(ByRef IsAmbil As Boolean, kode As String)
isAmbilDiscountMarketing = False
Me.Show 1
IsAmbil = ambil
kode = kodeBrg
End Sub
Public Sub TampilMarketing(ByRef IsAmbil As Boolean, ByRef kode As String, ByRef DiscountMarketing_ As Double)
isAmbilDiscountMarketing = True
Me.Show 1
IsAmbil = ambil
kode = kodeBrg
DiscountMarketing_ = DiscountMarketing
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
