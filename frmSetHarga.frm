VERSION 5.00
Begin VB.Form frmsetHarga 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1020
   ScaleWidth      =   5595
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
      Left            =   1680
      TabIndex        =   0
      Top             =   420
      Width           =   3705
   End
   Begin VB.Label lbNamaBarang 
      BackStyle       =   0  'Transparent
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
      TabIndex        =   2
      Top             =   120
      Width           =   4995
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   855
      Left            =   60
      Top             =   60
      Width           =   5475
   End
   Begin VB.Label lbCari 
      BackStyle       =   0  'Transparent
      Caption         =   "SET HARGA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
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
      Top             =   450
      Width           =   1575
   End
End
Attribute VB_Name = "frmsetHarga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ambil As Boolean
Dim isHasilKonversi As Boolean
Dim harga As Double
Dim NamaBarang As String

Private Sub Form_Load()
isHasilKonversi = False
Text1.Text = ""
lbNamaBarang = NamaBarang
'Text1.SetFocus
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim hasil As String
hasil = Trim(SendByCode(KeyCode))
KeyCode = 0
Select Case hasil
Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "00", "."
  Text1.Text = Text1.Text & hasil
  Text1.SelStart = Len(Text1.Text)
Case "BKS"
  If Len(Text1.Text) > 0 Then Text1.Text = Left(Text1.Text, Len(Text1.Text) - 1)
   Text1.SelStart = Len(Text1.Text)
Case "CLR"
    Text1.Text = ""

Case "ENT"
    If IsNumeric(Text1.Text) Then
     If CCur(Text1.Text) <= 0 Then
       Text1.Text = ""
       Text1.SetFocus
     Else
       harga = CCur(Text1.Text)
       Unload Me
     End If
    Else
       Text1.Text = ""
       Text1.SetFocus
    End If
Case "ESC"
    harga = -1
    Unload Me
End Select
End Sub
Public Sub Tampil(ByRef HargaJual As Double, ByVal Nama)
  NamaBarang = Nama
  Me.Show 1
  HargaJual = harga
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

