VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmRevisiBayar 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4680
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   4140
      Width           =   3075
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   3640
      Width           =   3075
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   3140
      Width           =   3075
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   2640
      Width           =   3075
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   150
      Width           =   3075
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
      Bindings        =   "frmRevisiBayar.frx":0000
      Height          =   2025
      Left            =   150
      OleObjectBlob   =   "frmRevisiBayar.frx":0014
      TabIndex        =   10
      Top             =   570
      Width           =   6525
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
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
      Height          =   405
      Left            =   480
      TabIndex        =   9
      Top             =   4200
      Width           =   2985
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PEMBAYARAN BANK"
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
      Height          =   405
      Left            =   480
      TabIndex        =   7
      Top             =   3705
      Width           =   2985
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PEMBAYARAN TUNAI"
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
      Height          =   405
      Left            =   360
      TabIndex        =   6
      Top             =   3140
      Width           =   3105
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "DISKON NOTA"
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
      Height          =   405
      Left            =   360
      TabIndex        =   5
      Top             =   2700
      Width           =   3105
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   4545
      Left            =   60
      Top             =   60
      Width           =   6705
   End
   Begin VB.Label lbCari 
      BackStyle       =   0  'Transparent
      Caption         =   "CARI KODE NOTA YG DIREVISI"
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
      Height          =   405
      Left            =   150
      TabIndex        =   1
      Top             =   270
      Width           =   5385
   End
End
Attribute VB_Name = "frmRevisiBayar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim ambil As Boolean
Dim idp As Long
Dim isHasilKonversi As Boolean
Dim SubTotal As Double
Dim BiayaCC As Double
Dim totalBayarCC As Long
Dim Diskon As Double
Dim Tunai As Double
Dim Bank As Double
Dim IDBank As Integer
Dim IDBankServer As Integer
Dim jwb As Boolean
Dim Total As Double

Dim NoAcc As String
Dim KodeBank As String
Dim NamaBank As String
Dim ChargeBank As Double
Dim NamaJenisKartu As String
Private Sub Form_Load()
isHasilKonversi = False
  ambil = False
  Data1.DatabaseName = DirDatabase & "\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
  Data1.RecordSource = "Select * From MSALES WHERE ISPEnding=FALSE ORDER BY Kode"
  Data1.Refresh
End Sub
Sub Tampil(ByRef idPending As Long, ByRef IsAmbil As Boolean)
Me.Show 1
idPending = idp
IsAmbil = ambil
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
tampilBayar
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim idBrg As Long
Dim isSimpanBrg As Boolean
Dim SelBks
Dim hasil As String
hasil = Trim(SendByCode(KeyCode))
Select Case hasil
Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "'", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "-", "+", "=", "|", "\", ".", ":", ";", "?"
  Text1.Text = Text1.Text & hasil
  Text1.SelStart = Len(Text1.Text)
 Case "SPC"
  Text1.Text = Text1.Text & " "
  Text1.SelStart = Len(Text1.Text)
'Case "{PGUP}"
'Case "{PGDN}"
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
        tampilBayar
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
       tampilBayar
Case "CLR"
        Text1.Text = ""
        Data1.Recordset.MoveFirst
        Set SelBks = TDBGrid1.SelBookmarks
        While SelBks.Count <> 0
            SelBks.Remove 0
        Wend
        TDBGrid1.SelBookmarks.Add TDBGrid1.Bookmark
        tampilBayar
Case "ENT"
      ambil = True
        tampilBayar
      Text2.SetFocus
Case "ESC"
      ambil = False
      Unload Me
Case "BKS"
  If Len(Text1.Text) > 0 Then Text1.Text = Left(Text1.Text, Len(Text1.Text) - 1)
   Text1.SelStart = Len(Text1.Text)
End Select
End Sub
Sub tampilBayar()
   If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
      Else
      idp = Data1.Recordset!NoID
      SubTotal = Data1.Recordset!SubTotal
      Total = Data1.Recordset!Hargatotal
      Diskon = Data1.Recordset!DiscNota
      Tunai = Data1.Recordset!UangMuka - Data1.Recordset!Bank
      Bank = Data1.Recordset!Bank
      Text2.Text = Format(Diskon, "##0")
      Text3.Text = Format(Tunai, "##0")
      Text4.Text = Format(Bank, "##0")
      Text5.Text = Format(Total, "##0")
      Text2.SelStart = 0
      Text2.SelLength = Len(Text2.Text)
      End If
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

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim hasil As String
hasil = Trim(SendByCode(KeyCode))
Select Case hasil
Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
  Text2.Text = Text2.Text & hasil
  Text2.SelStart = Len(Text2.Text)
Case "DN"
      Text3.SetFocus
      Text3.SelStart = 0 'Len(Text3.Text)
      'Text2.SelLength = Len(Text2.Text)
      Text3.SelLength = Len(Text3.Text)
Case "UP"
    Text4.SetFocus
      
Case "CLR"
        Text2.Text = ""
Case "ENT"
      Text3.SetFocus
Case "ESC"
      Text1.SetFocus
Case "BKS"
  If Len(Text2.Text) > 0 Then Text2.Text = Left(Text2.Text, Len(Text2.Text) - 1)
   Text2.SelStart = Len(Text2.Text)
End Select
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
Dim hasil As String
hasil = Trim(SendByCode(KeyCode))
Select Case hasil
Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
  Text3.Text = Text3.Text & hasil
  Text3.SelStart = Len(Text3.Text)
Case "DN"
      Text4.SetFocus
Case "UP"
    Text2.SetFocus
      
Case "CLR"
        Text3.Text = ""
Case "ENT"
      Text4.SetFocus
Case "ESC"
      Text2.SetFocus
Case "BKS"
  If Len(Text3.Text) > 0 Then Text3.Text = Left(Text3.Text, Len(Text3.Text) - 1)
   Text3.SelStart = Len(Text3.Text)
End Select

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
Dim hasil As String
hasil = Trim(SendByCode(KeyCode))
Select Case hasil
Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
  Text4.Text = Text4.Text & hasil
  Text4.SelStart = Len(Text4.Text)
Case "DN"
      Text2.SetFocus
Case "UP"
    Text3.SetFocus
      
Case "CLR"
        Text4.Text = ""
Case "ENT"
        If Not IsNumeric(Text2.Text) Then
            Text2.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(Text3.Text) Then
            Text3.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(Text4.Text) Then
            Text4.SetFocus
            Exit Sub
        End If
        If CDbl(Text2.Text) + CDbl(Text3.Text) + CDbl(Text4.Text) <> Total Then
            frmPesan.lbPesan = "JUMLAH PEMBAYARAN BELUM BENAR!"
            frmPesan.Show 1
            Exit Sub
        End If
        If CDbl(Text4.Text) = 0 Then
            IDBank = 0
            IDBankServer = 0
        Else
            frmBank.Tampil jwb, IDBank, IDBankServer, NoAcc, KodeBank, NamaBank, ChargeBank, 67, CDbl(Text4.Text), BiayaCC, totalBayarCC, True, -1, NamaJenisKartu, False ' IIf(IDMember > 0 And DiscINTERNNOTA > 0, True, False)
            If jwb And IDBank <> 0 Then
            Else
                frmPesan.lbPesan = "NAMA BANK HARUS DIPILIH!"
                frmPesan.Show 1
                Exit Sub
            End If
        End If
        Data1.Recordset.Edit
        Data1.Recordset!IDBank = IDBankServer
        Data1.Recordset!NoAcc = NoAcc
        Data1.Recordset!idcustomer = IDBank
        Data1.Recordset!KodeBank = KodeBank
        Data1.Recordset!NamaBank = NamaBank
        Data1.Recordset!NamaJenisKartu = NamaJenisKartu

        Data1.Recordset!Charge = ChargeBank
        Data1.Recordset!DiscNota = CDbl(Text2.Text)
        Data1.Recordset!UangMuka = CDbl(Text3.Text) + CDbl(Text4.Text)
        Data1.Recordset!Bank = totalBayarCC 'CDbl(Text4.Text)
        Data1.Recordset.Update
        Data1.Recordset.Bookmark = Data1.Recordset.LastModified
        frmPesan.lbPesan = "PEMBAYARAN TELAH TERSIMPAN!"
        frmPesan.Show 1
        Text1.SetFocus
      'Text.SetFocus
Case "ESC"
      Text3.SetFocus
Case "BKS"
  If Len(Text4.Text) > 0 Then Text4.Text = Left(Text4.Text, Len(Text4.Text) - 1)
   Text4.SelStart = Len(Text4.Text)
End Select

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

