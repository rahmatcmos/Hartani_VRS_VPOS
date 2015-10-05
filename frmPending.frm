VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmPending 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13575
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4680
   ScaleWidth      =   13575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\program\Database\TempDb.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3180
      TabIndex        =   0
      Top             =   150
      Width           =   10260
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\program\Database\TempDb.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3630
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "MSales"
      Top             =   -60
      Visible         =   0   'False
      Width           =   1140
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Bindings        =   "frmPending.frx":0000
      Height          =   3945
      Left            =   150
      OleObjectBlob   =   "frmPending.frx":0014
      TabIndex        =   2
      Top             =   570
      Width           =   8175
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid2 
      Bindings        =   "frmPending.frx":4156
      Height          =   3945
      Left            =   8325
      OleObjectBlob   =   "frmPending.frx":416A
      TabIndex        =   3
      Top             =   570
      Width           =   5115
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   4545
      Left            =   60
      Top             =   60
      Width           =   13455
   End
   Begin VB.Label lbCari 
      BackStyle       =   0  'Transparent
      Caption         =   "CARI KODE DATA PENDING"
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
      Left            =   150
      TabIndex        =   1
      Top             =   270
      Width           =   5385
   End
End
Attribute VB_Name = "frmPending"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim ambil As Boolean
Dim idp As Long
Dim isHasilKonversi As Boolean

Private Sub Form_Load()
isHasilKonversi = False
  ambil = False
  Data1.DatabaseName = DirDatabase & "\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
  Data2.DatabaseName = DirDatabase & "\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
  Data1.RecordSource = "Select * From MSALES WHERE ISPEnding=TRUE ORDER BY Tanggal DESC, Kode"
  Data1.Refresh
  ViewDetil
End Sub
Sub Tampil(ByRef idPending As Long, ByRef IsAmbil As Boolean)
Me.Show 1
idPending = idp
IsAmbil = ambil
End Sub

Private Sub Form_Unload(Cancel As Integer)
Data1.Database.Close
Data2.Database.Close
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  ViewDetil
End Sub
Private Sub ViewDetil()
Dim SelBks
Dim SQL As String
  If Not (Data1.Recordset.EOF Or Data1.Recordset.BOF) Then
    SQL = "SELECT MSalesD.[NoID], MSalesD.[KodeInv] AS Kode, MSalesD.[NamaInv] AS Nama, MSalesD.[Qty] FROM MSalesD INNER JOIN MSales ON MSalesD.IDSales=MSales.NoID WHERE MSales.NoID=" & NullToNol(Data1.Recordset!NoID) & " ORDER BY MSalesD.[NoID]"
    Data2.RecordSource = SQL
    Data2.Refresh
    If Not (Data2.Recordset.EOF Or Data2.Recordset.BOF) Then
      Data2.Recordset.MoveFirst
      Set SelBks = TDBGrid2.SelBookmarks
          While SelBks.Count <> 0
              SelBks.Remove 0
          Wend
          TDBGrid2.SelBookmarks.Add TDBGrid2.Bookmark
    End If
  End If
End Sub
Private Sub Text1_Change()
If Trim(Text1.Text) = "" Then Exit Sub
Data1.Recordset.FindFirst "left(Kode," & Len(Text1.Text) & ")='" & Replace(Text1.Text, "'", "''") & "'"
If Data1.Recordset.NoMatch Then
  Data1.Recordset.MoveFirst
Else
End If
'RECORDSET
 'Data1.Recordset.FindFirst NamaField & " LIKE '" & Replace(Text1.Text, "'", "''") & "%'"
 'If Data1.Recordset.NoMatch Then Data1.Recordset.MoveFirst
 Dim SelBks
  Set SelBks = TDBGrid1.SelBookmarks
  While SelBks.Count <> 0
      SelBks.Remove 0
  Wend
  TDBGrid1.SelBookmarks.Add TDBGrid1.Bookmark

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim idBrg As Long
Dim isSimpanBrg As Boolean
Dim SelBks
Dim Hasil As String
Hasil = Trim(SendByCode(KeyCode))
Select Case Hasil
Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "'", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "-", "+", "=", "|", "\", ".", ":", ";", "?"
  Text1.Text = Text1.Text & Hasil
  Text1.SelStart = Len(Text1.Text)
 Case "SPC"
  Text1.Text = Text1.Text & " "
  Text1.SelStart = Len(Text1.Text)
'Case "{PGUP}"
'Case "{PGDN}"
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
Case "CLR"
        Text1.Text = ""
        Data1.Recordset.MoveFirst
        Set SelBks = TDBGrid1.SelBookmarks
        While SelBks.Count <> 0
            SelBks.Remove 0
        Wend
        TDBGrid1.SelBookmarks.Add TDBGrid1.Bookmark
Case "ENT"
      ambil = True
      idp = Data1.Recordset!NoID
      Data1.Recordset.Edit
      Data1.Recordset!ispending = False
      Data1.Recordset.Update
      Data1.Recordset.Bookmark = Data1.Recordset.LastModified
      Unload Me
Case "ESC"
      ambil = False
      Unload Me
Case "BKS"
  If Len(Text1.Text) > 0 Then Text1.Text = Left(Text1.Text, Len(Text1.Text) - 1)
   Text1.SelStart = Len(Text1.Text)
End Select
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
'If Not (isHasilKonversi) Then
'    KeyAscii = 0
'    isRun = True
'    isHasilKonversi = True
'    SendKeys SendByCode(KeyKode), True
'    isHasilKonversi = False
'  Else
'    If KeyAscii = 27 Then
'      KeyAscii = 0
'      ambil = False
'      Unload Me
'    ElseIf KeyAscii = 13 Then
'      ambil = True
'      KeyAscii = 0
'      idp = Data1.Recordset!NoId
'      Data1.Recordset.Edit
'      Data1.Recordset!iSPENDING = False
'      Data1.Recordset.Update
'      Data1.Recordset.Bookmark = Data1.Recordset.LastModified
'      Unload Me
'    End If
'  End If
End Sub

