VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmsetItemRetur 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   570
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1830
      TabIndex        =   0
      Top             =   90
      Width           =   3555
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
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
      Height          =   495
      Left            =   60
      Top             =   30
      Width           =   5475
   End
   Begin VB.Label lbCari 
      BackStyle       =   0  'Transparent
      Caption         =   "KODE BARANG"
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
      Left            =   120
      TabIndex        =   1
      Top             =   150
      Width           =   1755
   End
End
Attribute VB_Name = "frmsetItemRetur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ambil As Boolean
Dim isHasilKonversi As Boolean
Dim kode As String
Dim BarcodeIn As String

Private Sub Form_Load()
  isHasilKonversi = False
  Text1.Text = ""
  MSComm1.PortOpen = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
MSComm1.PortOpen = False
End Sub

Private Sub MSComm1_OnComm()
Select Case MSComm1.CommEvent
        ' Event messages.
        Case comEvReceive
            Dim Buffer As Variant
            Dim pos As Integer
            Buffer = MSComm1.Input
            Debug.Print "Receive - " & StrConv(Buffer, vbUnicode)
            BarcodeIn = BarcodeIn & StrConv(Buffer, vbUnicode)
            pos = InStr(1, BarcodeIn, Chr(13))
            If pos Then
                kode = Left(BarcodeIn, pos - 1)
                Text1.Text = kode
                Unload Me
            End If
            'ShowData txtTerm, (StrConv(Buffer, vbUnicode))
End Select
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
 If isHasilKonversi And (KeyCode = 8 Or KeyCode = 9 Or KeyCode = 13 Or KeyCode = 27) Then
    KeyCode = 0
    Exit Sub
  End If
  
  If Not (isHasilKonversi) Then
    KeyKode = KeyCode
    KeyCode = 0
    isRun = False
  Else
  End If
End Sub
Public Sub tampil(ByRef IDBArang As String)
  Me.Show 1
  IDBArang = kode
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  If Not (isHasilKonversi) Then
    KeyAscii = 0
    isRun = True
    isHasilKonversi = True
    SendKeys SendByCode(KeyKode), True
    isHasilKonversi = False
  Else
   If KeyAscii = 13 Then
          kode = Text1.Text
          KeyAscii = 0
          Unload Me
    ElseIf KeyAscii = 27 Then
        kode = "-1"
        KeyAscii = 0
        Unload Me
    End If
    
  End If
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not isRun Then
        isRun = True
        isHasilKonversi = True
        SendKeys (SendByCode(KeyKode)), True
        isHasilKonversi = False
    End If
End Sub
