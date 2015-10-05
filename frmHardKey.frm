VERSION 5.00
Begin VB.Form frmKey 
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
      PasswordChar    =   "#"
      TabIndex        =   0
      Top             =   780
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
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Kunci 2 untuk OK 1 untuk Batal"
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
      Left            =   -30
      TabIndex        =   1
      Top             =   330
      Width           =   4605
   End
End
Attribute VB_Name = "frmKey"
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
Dim KEYCTRL As Boolean
Dim KEYK As Boolean
Dim KEY1 As Boolean
Dim KEY4 As Boolean
Dim KEYBatal As Boolean


Private Sub Form_Load()
KEYCTRL = False
KEYK = False
KEY1 = False
KEY4 = False

isHasilKonversi = False
  jawab = False
'  If GetStatusNetwork = "Local" Then
'    Set dbs = OpenDatabase(DirDatabase & "\dbMaster.mdb")
'  Else
'    Set dbs = OpenDatabase(DirDbServer & "\dbMaster.mdb")
'  End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
'Dim hasil As String
'hasil = Trim(SendByCode(KeyCode))
'Select Case hasil
'Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "'", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "-", "+", "=", "|", "\", ".", ":", ";", "?"
'  Text1.Text = Text1.Text & hasil
'  Text1.SelStart = Len(Text1.Text)
' Case "SPC"
'  Text1.Text = Text1.Text & " "
'  Text1.SelStart = Len(Text1.Text)
'Case "BKS"
'  If Len(Text1.Text) > 0 Then Text1.Text = Left(Text1.Text, Len(Text1.Text) - 1)
'   Text1.SelStart = Len(Text1.Text)
'Case "CLR"
'        Text1.Text = ""
'Case "ENT"
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
'Case "ESC"
'    jawab = False
'    Unload Me
'End Select
If KeyCode = 17 Or KEYCTRL Then
  KEYCTRL = True
  If KeyCode = 75 Or KEYK Then
    KEYK = True
    If (KeyCode = 50 Or KeyCode = 100) Then
      KEYBatal = False
      KEY4 = True
      lbTanya.Caption = "Kembalikan Kunci ke Posisi 0"
      lbTanya.Refresh
      DoEvents
    ElseIf (KeyCode = 49 Or KeyCode = 97) And Not KEY4 Then
        KEYBatal = True
    ElseIf (KeyCode = 48 Or KeyCode = 96) And KEY4 Then
      jawab = True
      KeyCode = 0
      Unload Me
    ElseIf (KeyCode = 48 Or KeyCode = 96) And KEYBatal Then
      jawab = False
      KeyCode = 0
      Unload Me
    End If
  End If
End If
KeyCode = 0
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Sub Tampil(ByRef jawaban As Boolean, Key As Integer)
  KeyOk = Key
  Me.Show 1
  jawaban = jawab
End Sub
