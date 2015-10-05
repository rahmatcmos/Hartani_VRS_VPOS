VERSION 5.00
Begin VB.Form frmAgen 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   ScaleHeight     =   1500
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
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
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   1620
      TabIndex        =   1
      Top             =   285
      Width           =   4845
   End
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
      Left            =   1620
      TabIndex        =   3
      Top             =   795
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grup"
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
      Height          =   240
      Left            =   405
      TabIndex        =   0
      Top             =   330
      Width           =   510
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   1245
      Left            =   90
      Top             =   90
      Width           =   6435
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   1365
      Left            =   60
      Top             =   60
      Width           =   6645
   End
   Begin VB.Label lbTanya 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Komisi (%)"
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
      Height          =   240
      Left            =   405
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   1110
   End
End
Attribute VB_Name = "frmAgen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim jawab As Boolean
Dim Komisi_ As String
Dim Agen_ As String
Dim KeyOk As Integer
Private Sub Form_Load()
isHasilKonversi = False
  jawab = False
  Text2.Text = Agen_
  Text1.Text = Komisi_
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
 
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
    Komisi_ = Text1.Text
    Agen_ = Text2.Text
    jawab = True
    Unload Me
    Else
    frmPesan.lbPesan = "JUMLAH PEMBAYARAN BELUM BENAR!"
    frmPesan.Show 1
      Text2.SetFocus
    End If
Case "ESC"
    Text1.SetFocus
End Select
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Sub Tampil(ByRef jawaban As Boolean, Key As Integer, ByRef Agen As String, ByRef Komisi As Double)
  KeyOk = Key
  Agen_ = Agen
  Komisi_ = Komisi
  Me.Show 1
If jawab Then
  Agen = Agen_
  Komisi = Komisi_
End If
  jawaban = jawab
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim hasil As String
Dim isambil As Boolean
Dim kdMember As String
Dim discMarketing As Double
kdMember = ""
isambil = False
hasil = Trim(SendByCode(KeyCode))
Select Case hasil
Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "'", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "-", "+", "=", "|", "\", ".", ":", ";", "?"
  Text2.Text = Text2.Text & hasil
  Text2.SelStart = Len(Text2.Text)
Case "SPC"
  Text2.Text = Text2.Text & " "
  Text2.SelStart = Len(Text2.Text)
Case "BKS"
  If Len(Text2.Text) > 0 Then Text2.Text = Left(Text2.Text, Len(Text2.Text) - 1)
   Text2.SelStart = Len(Text2.Text)
Case "CLR"
        Text2.Text = ""
Case "ENT"
   Text1.SetFocus
Case "SCK"
    frmCustomer.NamaField = "Kode"
    frmCustomer.lbCari = "CARI KODE"
    frmCustomer.TampilMarketing isambil, kdMember, discMarketing
    If isambil Then
         
        Text2.Text = kdMember
        Text1.Text = Format(discMarketing, "##0.##")
        'GetIDAgen
        'If IDMember < 1 Then
        'Else
          'Unload Me
        'End If
    End If
Case "SCN"
    frmCustomer.NamaField = "Nama"
    frmCustomer.lbCari = "CARI NAMA"
    frmCustomer.TampilMarketing isambil, kdMember, discMarketing
    If isambil Then
        Text2.Text = kdMember
        Text1.Text = Format(discMarketing, "##0.##")
        'GetIDAgen
    End If
Case "ESC"
   jawab = False
    Unload Me
End Select
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
