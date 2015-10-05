VERSION 5.00
Begin VB.Form frmIsiModal 
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
      Caption         =   "Jumlah Modal Hari ini :"
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
Attribute VB_Name = "frmIsiModal"
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
If IsResetPerKasir Then
lbTanya.Caption = "Jumlah Modal Kasir Sekarang:"
Else
lbTanya.Caption = "Masukan Jumlah Modal Shift Sekarang:"
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
        If MsgBox("Modal Rp. " & Format(CCur(Text1.Text), "###,##0") & " Lanjut Simpan?", vbYesNo, "Menyimpan Modal") = vbYes Then
          Dim Modal As Double
          Dim dbs As Database
          Dim rs As Recordset
          Dim NoID As Long
          Set dbs = OpenDatabase(DirDatabase & "\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
          If IsResetPerKasir Then
            Set rs = dbs.OpenRecordset("Select * from MReset where tanggal=#" & Format(Now, "MM/dd/yyyy") & "# and IDKasir=" & IDUser & " AND Shift=" & NamaShift, dbOpenDynaset)
          Else
            Set rs = dbs.OpenRecordset("Select * from MReset where tanggal=#" & Format(Now, "MM/dd/yyyy") & "# and Shift=" & NamaShift, dbOpenDynaset)
          End If
          If rs.EOF And rs.BOF Then
            NoID = GetNewID("MReset")
            rs.AddNew
            rs!NoID = NoID
            rs!TANGGAL = Date
            rs!Shift = NamaShift
            rs!Modal = CCur(Text1.Text)
            rs!KodeKasir = KodeKasir
            rs!NamaKasir = NamaKasir
            rs!IDKasir = IDUser
            rs.Update
            rs.Bookmark = rs.LastModified
          Else
            rs.Edit
            rs!TANGGAL = Date
            rs!Shift = NamaShift
            rs!Modal = CCur(Text1.Text)
            rs!KodeKasir = KodeKasir
            rs!NamaKasir = NamaKasir
            rs!IDKasir = IDUser
            rs.Update
            rs.Bookmark = rs.LastModified
          End If
          rs.Close
          dbs.Close
          Set rs = Nothing
          Set dbs = Nothing
          jawab = True
          Unload Me
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

Private Sub Text1_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Sub Tampil(ByRef jawaban As Boolean, Key As Integer)
  KeyOk = Key
  Me.Show 1
  jawaban = jawab
End Sub
