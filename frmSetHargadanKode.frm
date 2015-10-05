VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmsetHargadanKode 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1455
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
   End
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
      Left            =   2280
      TabIndex        =   3
      Top             =   840
      Width           =   3105
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
      Height          =   420
      Left            =   2280
      TabIndex        =   0
      Top             =   420
      Width           =   3105
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Barcode"
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
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   3255
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
      Height          =   1335
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
Attribute VB_Name = "frmsetHargadanKode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ambil As Boolean
Dim isHasilKonversi As Boolean
Dim harga As Double
Dim NamaBarang As String
Dim Kodebarang As String
Dim BarcodeIn As String
Private Sub Form_Load()
On Error Resume Next
isHasilKonversi = False
Text1.Text = ""
lbNamaBarang = NamaBarang
  MSComm1.CommPort = NoPortBarcode
  MSComm1.PortOpen = True
  BarcodeIn = ""
'Text1.SetFocus
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
  MSComm1.PortOpen = False
End Sub

Private Sub MSComm1_OnComm()
Dim kode As String
Select Case MSComm1.CommEvent
        ' Event messages.
        Case comEvReceive
            Dim Buffer As Variant
            Dim pos As Integer
            Buffer = MSComm1.Input
            'Debug.Print "Receive - " & StrConv(Buffer, vbUnicode)
            BarcodeIn = BarcodeIn & StrConv(Buffer, vbUnicode)
            pos = InStr(1, BarcodeIn, Chr(13))
            If pos Then
                kode = Left(BarcodeIn, pos - 1)
                BarcodeIn = ""
                Text2.Text = kode
                SendKeys "{ENTER}", False
'
'              Unload Me
              Exit Sub
            End If
            'ShowData txtTerm, (StrConv(Buffer, vbUnicode))
End Select
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
       Text2.SetFocus
     End If
    Else
       Text1.Text = ""
       Text1.SetFocus
    End If
Case "ESC"
    harga = -1
    Unload Me
Case "DN"
    Text2.SetFocus
End Select
End Sub
Public Sub Tampil(ByRef HargaJual As Double, ByRef kode As String)
  Kodebarang = kode
  Me.Show 1
  kode = Kodebarang
  HargaJual = harga
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim hasil As String
    hasil = Trim(SendByCode(KeyCode))
    KeyCode = 0
    Select Case hasil
    Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "00", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "'", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "-", "+", "=", "|", "\", ".", ":", ";", "?", "[", "]"
      Text2.Text = Text2.Text & hasil
      Text2.SelStart = Len(Text2.Text)
    Case "BKS"
      If Len(Text2.Text) > 0 Then Text2.Text = Left(Text2.Text, Len(Text2.Text) - 1)
       Text2.SelStart = Len(Text2.Text)
    Case "CLR"
        Text2.Text = ""
    Case "ENT"
        Kodebarang = Text2.Text
        harga = IIf(IsNumeric(Text1.Text), Text1.Text, -1)
        Unload Me
    Case "ESC"
        harga = -1
        Unload Me
    Case "UP"
        Text1.SetFocus
    End Select
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
