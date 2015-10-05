VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmAmbilDuit 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1485
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   420
      TabIndex        =   0
      Top             =   690
      Width           =   3765
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   1245
      Left            =   120
      Top             =   120
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   1365
      Left            =   60
      Top             =   60
      Width           =   4575
   End
   Begin VB.Label lbTanya 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Pengambilan :"
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
      Height          =   555
      Left            =   405
      TabIndex        =   1
      Top             =   360
      Width           =   4605
   End
End
Attribute VB_Name = "frmAmbilDuit"
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
Private Sub Form_Load()
isHasilKonversi = False
  jawab = False
  If isOnline = False Then
    Set dbs = OpenDatabase(DirDatabase & "\dbMaster.mdb")
  Else
    Set dbs = OpenDatabase(DirDbServer & "\dbMaster.mdb")
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
dbs.Close
Set dbs = Nothing
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim hasil As String
hasil = Trim(SendByCode(KeyCode))
Select Case hasil
Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "'", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "-", "+", "=", "|", "\", ".", ":", ";", "?"
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
      If IsNumeric(Text1.Text) Then
      Text1.Enabled = False
        If MsgBox("Jumlah Pengambilan Rp. " & Format(CCur(Text1.Text), "###,##0") & " Lanjut Simpan?", vbYesNo, "Pengambilan Uang Tunai") = vbYes Then
          Dim Modal As Double
          Dim dbs As Database
          Dim rs As Recordset
          Dim NoID As Long
          Set dbs = OpenDatabase(DirDatabase & "\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
          Set rs = dbs.OpenRecordset("Select * from MKasKeluar where tanggal=#" & Format(Now, "MM/dd/yyyy") & "# and IDKasir=" & IDUser & " AND Shift=" & NamaShift, dbOpenDynaset)
            NoID = GetNewID("MKasKeluar")
            rs.AddNew
            rs!NoID = NoID
            rs!TANGGAL = Now
            rs!Shift = NamaShift
            rs!Jumlah = CCur(Text1.Text)
            rs!KodeKasir = KodeKasir
            rs!NamaKasir = NamaKasir
            rs!IDKasir = IDUser
            rs!KodePengawas = KodePengawas_
            rs!NamaPengawas = NamaPengawas_
            rs!IdPengawas = IDPengawas_
            rs.Update
            rs.Bookmark = rs.LastModified
          rs.Close
          dbs.Close
          Set rs = Nothing
          Set dbs = Nothing
          PrintAmbilDuit NoID
          jawab = True
          Unload Me
        Else
          Text1.Enabled = False
      End If
      Else
        frmPesan.lbPesan = "Masukan Nominal Uang...!"
        frmPesan.Show 1
      End If
'        Set rs = dbs.OpenRecordset("Select Password from Memp Where isPengawas=true")
'        If rs.EOF And rs.BOF Then
'        Else
'          rs.MoveFirst
'          rs.FindFirst "Password='" & Replace(Text1.Text, "'", "''") & "'"
'          If rs.NoMatch Then
'            Text1.Text = ""
'            Exit Sub
'          Else
'            jawab = True
'            Unload Me
'          End If
'        End If
Case "ESC"
    jawab = False
    Unload Me
End Select
End Sub
Sub PrintAmbilDuit(ByVal NoID As Long)
Dim dbsr As Database
Dim rst As Recordset
On Error GoTo Trace
  CrystalReport1.ReportFileName = App.path & "\Report\KasKeluar.rpt"
  CrystalReport1.DataFiles(0) = App.path & "\Database\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
  'cdate(2011,12,31)
  CrystalReport1.Formulas(0) = "NoID=" & NoID
  CrystalReport1.Formulas(1) = "NamaPerusahaan='" & Replace(NamaToko, "'", "''") & "'"
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
Trace:
  If Err.Number <> 0 Then
    'MsgBox Err.Number & " " & Err.Description, vbCritical, App.Title
    Err.Clear
  End If
End Sub



Private Sub Text1_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Sub Tampil(ByRef jawaban As Boolean, Key As Integer)
  KeyOk = Key
  Me.Show 1
  jawaban = jawab
End Sub
