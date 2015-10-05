VERSION 5.00
Begin VB.Form frmGetAngka 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1575
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   105
      TabIndex        =   0
      Top             =   600
      Width           =   5640
   End
   Begin VB.Label lbNamaBarang 
      BackStyle       =   0  'Transparent
      Caption         =   "Masukkan Diskon Barang dalam Rupiah :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   90
      TabIndex        =   1
      Top             =   120
      Width           =   6930
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   1485
      Left            =   60
      Top             =   60
      Width           =   5790
   End
End
Attribute VB_Name = "frmGetAngka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ambil As Boolean
Dim isHasilKonversi As Boolean
Dim harga As Double
Dim JudulForm As String
Public IsMinus As Boolean

Private Sub Form_Activate()
Me.Top = Me.Top + 2500
End Sub

Private Sub Form_Load()
isHasilKonversi = False
Text1.Text = ""
lbNamaBarang = JudulForm
'Text1.SetFocus
End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Trace
Dim hasil As String
hasil = Trim(SendByCode(KeyCode))
KeyCode = 0
Select Case hasil
Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "00", ".", IIf(IsMinus, "-", "")
  Text1.Text = Text1.Text & hasil
  Text1.SelStart = Len(Text1.Text)
Case "BKS"
  If Len(Text1.Text) > 0 Then Text1.Text = Left(Text1.Text, Len(Text1.Text) - 1)
   Text1.SelStart = Len(Text1.Text)
Case "CLR"
    Text1.Text = ""
Case "ENT"
    If IsNumeric(Text1.Text) Then
'     If CCur(Text1.Text) < 0 Then
'       Text1.Text = ""
'       Text1.SetFocus
'     Else
'       harga = CCur(Text1.Text)
'       Unload Me
'     End If
       harga = CCur(Text1.Text)
       Unload Me
    Else
       Text1.Text = ""
       Text1.SetFocus
    End If
Case "ESC"
    harga = -1
    Unload Me
End Select
Trace:
  If Err.Number <> 0 Then
    MsgBox "Angka melebihi standart. " & Err.Number & " : " & Err.Description, vbCritical, App.Title
    Err.Clear
  End If
End Sub
Public Sub Tampil(ByRef Angka As Double, ByVal Judul, ByVal Minus As Boolean)
  JudulForm = Judul
  IsMinus = Minus

  Me.Show 1

  Angka = harga
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
'    If KeyAscii = 13 Then
'      KeyAscii = 0
'        If IsNumeric(Text1.Text) Then
'          If CCur(Text1.Text) <= 0 Then
'            Text1.Text = ""
'            Text1.SetFocus
'          Else
'
'            harga = CCur(Text1.Text)
'            Unload Me
'          End If
'        Else
'            Text1.Text = ""
'            Text1.SetFocus
'        End If
'      ElseIf KeyAscii = 27 Then
'        harga = -1
'        Unload Me
'      End If
'  End If
End Sub

