VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmvoucher 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   810
   ScaleWidth      =   6015
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
      Width           =   3555
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
      Width           =   5865
   End
   Begin VB.Label lbCari 
      BackStyle       =   0  'Transparent
      Caption         =   "KODE VOUCHER"
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
Attribute VB_Name = "frmvoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ambil As Boolean
Dim isHasilKonversi As Boolean
Dim kode As String
Dim Nominal As Double
Dim BarcodeIn As String
Dim IDPenjualan As Long
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
  Nominal = 0
  MSComm1.CommPort = NoPortBarcode
  MSComm1.PortOpen = True
  BarcodeIn = ""

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
  MSComm1.PortOpen = False
End Sub

Private Sub MSComm1_OnComm()
    Select Case MSComm1.CommEvent
        ' Event messages.
        Case comEvReceive
            Dim buffer As Variant
            Dim pos As Integer
            buffer = MSComm1.Input
            'Debug.Print "Receive - " & StrConv(Buffer, vbUnicode)
            BarcodeIn = BarcodeIn & StrConv(buffer, vbUnicode)
            pos = InStr(1, BarcodeIn, Chr(13))
            If pos Then
                
                kode = Left(BarcodeIn, pos - 1)
                Text1.Text = kode
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
        'kode = Text1.Text
        CariNominalOnline
        Unload Me
        Exit Sub
    End If
    hasil = Trim(SendByCode(KeyCode))
    KeyCode = 0
    Select Case hasil
    Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
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
        kode = Text1.Text
        CariNominalOnline
        Unload Me
    Case "ESC"
        kode = "-1"
        Unload Me
    End Select
End Sub

Public Sub Tampil(ByRef NilaiVoucher As Double, ByVal IDJual As Long)
  IDPenjualan = IDJual
  Me.Show 1
  NilaiVoucher = Nominal
End Sub

Sub CariNominal()
On Error GoTo akhir
Dim dbs  As Database
Dim rs As Recordset
DoEvents
Set dbs = OpenDatabase(PathVoucher)
Set rs = dbs.OpenRecordset("Select  * from voucher where kode='" & Text1.Text & "'")
If rs.EOF And rs.BOF Then
  Nominal = 0
Else
  Nominal = IIf(IsNull(rs!Nominal), 0, rs!Nominal)
  If rs!IDSales > 0 Then
  Nominal = 0
  Else
    rs.Edit
    rs!TANGGAL = Date
    rs!kassa = NamaMesin
    rs!IDSales = IDPenjualan
    rs.Update
  End If
End If
Exit Sub
akhir:
Nominal = 0
End Sub
Sub CariNominalOnline()
On Error GoTo akhir
Dim dbs  As ADODB.Connection
Dim rs As ADODB.Recordset
DoEvents

   bacaSettingServer
'    Dim isOnline As Boolean
'    Dim sqlcon As New ADODB.Connection
'    Set sqlcon = New ADODB.Connection
'    sqlcon.ConnectionString = Cnstr
'    sqlcon.Open
'    sqlcon.Execute sql
'    sqlcon.Close
'    Set sqlcon = Nothing
'    ExecuteSQL = "ONLINE"
'    Exit Function


Set dbs = New ADODB.Connection
 dbs.ConnectionString = Cnstr
    dbs.Open
'    Set rs = dbs.Execute("select  * from MVoucher where IsSend=1 and " & _
'"IsBack=0 and TanggalJT>=convert(datetime,'" & Format(Date, "MM/dd/yyyy") & "',101) " & _
'"and (convert(bigint,kode)=" & Text1.Text & " or convert(bigint,Barcode)=" & Text1.Text & ")")

Set rs = dbs.Execute("select  * from MVoucher where IsSend=1 and " & _
"IsBack=0   " & _
"and (convert(bigint,kode)=" & Text1.Text & " or convert(bigint,Barcode)=" & Text1.Text & ")")
If rs.EOF And rs.BOF Then
  Nominal = 0
Else
    If rs!TanggalJT >= Date Then
      Nominal = IIf(IsNull(rs!Nominal), 0, rs!Nominal)
      If rs!IDSales > 0 Then
        Nominal = 0
      Else
        ExecuteSQL "Update MVoucher Set IsBack=1,tanggalBack=getdate(),IDPOS=" & CInt(NamaMesin) & ",KodePos='" & NamaMesin & "',IDSales=" & IDPenjualan & "  where NoID=" & rs!NoID
       End If
    Else
        Nominal = 0
        frmPesan.lbPesan = "Voucher Expired.."
        frmPesan.Show 1
    End If
End If
Exit Sub
akhir:
Nominal = 0
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
'        Nominal = CCur(Text1.Text)
'      Else
'        Nominal = 1
'      End If
'      Text1.Text = ""
'      KeyAscii = 0
'      Form3.lbQTY = Trim(Str(Nominal)) & " X"
'    End If
'  End If
End Sub
