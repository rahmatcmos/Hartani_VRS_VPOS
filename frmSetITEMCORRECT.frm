VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmsetItemCorrect 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   810
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2250
      TabIndex        =   0
      Top             =   120
      Width           =   5115
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
      NullDiscard     =   -1  'True
      RThreshold      =   1
      RTSEnable       =   -1  'True
      SThreshold      =   1
      InputMode       =   1
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   705
      Left            =   60
      Top             =   30
      Width           =   7425
   End
   Begin VB.Label lbCari 
      BackStyle       =   0  'Transparent
      Caption         =   "QTY X BARANG"
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
      Top             =   180
      Width           =   2205
   End
End
Attribute VB_Name = "frmsetItemCorrect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ambil As Boolean
Dim isHasilKonversi As Boolean
Dim kode As String
Dim qtyBrg As Double
Dim BarcodeIn As String

Sub BukaCommBarcode()
On Error Resume Next
MSComm1.PortOpen = True
End Sub

Sub TutupCommBarcode()
On Error Resume Next
MSComm1.PortOpen = False
End Sub


Private Sub Form_Activate()
    BukaCommBarcode
End Sub

Private Sub Form_DeActivate()
    TutupCommBarcode
End Sub

Private Sub Form_Load()
On Error Resume Next
  isHasilKonversi = False
  Text1.Text = ""
  qtyBrg = 1
  MSComm1.CommPort = NoPortBarcode
  MSComm1.PortOpen = True
  BarcodeIn = ""
  lbCari.Caption = "QTY X BARANG"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
  MSComm1.PortOpen = False
End Sub

Private Sub MSComm1_OnComm()
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
                Text1.Text = kode
                kode = Left(BarcodeIn, pos - 1)
                BarcodeIn = ""
                
                'kode = Text1.Text
              '  Unload Me
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
    If KeyCode = 13 Then
        KeyCode = 0
        kode = Text1.Text
        Unload Me
        Exit Sub
    End If
    hasil = Trim(SendByCode(KeyCode))
    KeyCode = 0
    Select Case hasil
    Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "'", "!", "@", "#", "$", "%", "^", "&", "(", ")", "_", "-", "+", "=", "|", "\", ".", ":", ";", "?", "/"
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
    Case "*"
          If IsNumeric(Text1.Text) Then
            qtyBrg = CCur(Text1.Text)
          Else
            qtyBrg = 1
          End If
          lbCari.Caption = "KODE BARANG"
          Text1.Text = ""
          Form3.lbQTY = Trim(Str(qtyBrg)) & " X"
    Case "ENT"
        kode = Text1.Text
        Unload Me
    Case "ESC"
        qtyBrg = 1
        Form3.lbQTY = Trim(Str(qtyBrg)) & " X"
        kode = "-1"
        Unload Me
    End Select
End Sub
Public Sub Tampil(ByRef jumBArang As Double, ByRef IDBArang As String)
  Me.Show 1
  IDBArang = kode
  jumBArang = qtyBrg
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
' 'PENGGGUNAAN BARCODE BERSAMA KEYBOARD
'  If (KeyAscii >= 48 And KeyAscii <= 57) Then
'    isRun = True
'    Exit Sub
'  End If
'  If Not (isHasilKonversi) Then
'    KeyAscii = 0
'    isRun = True
'    isHasilKonversi = True
'    SendKeys SendByCode(KeyKode), True
'    isHasilKonversi = False
'  Else
'   If KeyAscii = 13 Then
'          kode = Text1.Text
'          KeyAscii = 0
'          Unload Me
'    ElseIf KeyAscii = 27 Then
'        kode = "-1"
'        KeyAscii = 0
'        Unload Me
'    ElseIf KeyAscii = 42 Then '* quantity
'      If IsNumeric(Text1.Text) Then
'        qtyBrg = CCur(Text1.Text)
'      Else
'        qtyBrg = 1
'      End If
'      Text1.Text = ""
'      KeyAscii = 0
'      Form3.lbQTY = Trim(Str(qtyBrg)) & " X"
'    End If
'  End If
End Sub
