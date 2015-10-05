VERSION 5.00
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dxctrls.dll"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmLookUpTKPSQLServer 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13050
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   13050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TrueOleDBGrid60.TDBGrid TDBGrid1 
      Bindings        =   "frmLookUpTKPSQLServer.frx":0000
      Height          =   6105
      Left            =   30
      OleObjectBlob   =   "frmLookUpTKPSQLServer.frx":0015
      TabIndex        =   4
      Top             =   600
      Width           =   13005
   End
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
      Left            =   1230
      TabIndex        =   0
      Top             =   180
      Width           =   9810
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   480
      Left            =   45
      Top             =   990
      Visible         =   0   'False
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   847
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin CONTROLSLibCtl.dxCheckBox ckAll 
      Height          =   270
      Left            =   11100
      TabIndex        =   5
      Top             =   225
      Visible         =   0   'False
      Width           =   1845
      _Version        =   65536
      _cx             =   3254
      _cy             =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Tampilkan Semua"
      Enabled         =   -1  'True
      AutoSize        =   -1  'True
      BackStyle       =   0
      BackColor       =   15790320
      ForeColor       =   16777215
      ViewStyle       =   1
      Checked         =   0   'False
      GroupIndex      =   -1
      TextLayout      =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   12632256
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OK <ENTER>_____CANCEL <ESC>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   6780
      Visible         =   0   'False
      Width           =   12930
   End
   Begin VB.Label lbCari 
      BackStyle       =   0  'Transparent
      Caption         =   "CARI KET"
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
      TabIndex        =   1
      Top             =   210
      Width           =   2895
   End
   Begin CONTROLSLibCtl.dxBackground dxBackground1 
      Height          =   7245
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   13095
      _Version        =   65536
      _cx             =   23098
      _cy             =   12779
      StartColor      =   33023
      EndColor        =   12640511
      ColorFillStyle  =   0
      BackgroundStyle =   0
      DrawPictureStyle=   1
      AnimationInterval=   1000
   End
End
Attribute VB_Name = "frmLookUpTKPSQLServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim isHasilKonversi As Boolean
'Dim kodeBrg As String
Public NamaField As String
Dim ambil As Boolean
Dim NamaBrg As String
Dim NoIdbrg As Long
Dim issimpan As Boolean
Dim NoIDKas As Long

'Public LookUp As IsLookUpKas
'Dim mCon As New ADODB.Connection
'Dim mCom As New ADODB.Command
'Dim rst As New ADODB.Recordset
Dim SQL As String

Private Sub ckAll_Click()
  TampilkanData
End Sub

Private Sub Form_Activate()
If isSupervisor Then
  Label1.Visible = True
Else
  Label1.Visible = False
End If
End Sub

Private Sub Form_Load()
  Adodc1.ConnectionString = Cnstr
  isHasilKonversi = False
  ambil = False
  TampilkanData
  
  ' Load the layout.
  TDBGrid1.LayoutName = "Layoutku"
  If Dir(App.path & "\" & Me.Name & "_" & TDBGrid1.Name & ".grx") <> "" Then
    TDBGrid1.LayoutFileName = App.path & "\" & Me.Name & "_" & TDBGrid1.Name & ".grx"
    TDBGrid1.LoadLayout
  End If
End Sub

Sub View()
If Text1.Text <> "" Then
  Adodc1.Recordset.Find NamaField & " LIKE '" & Replace(Text1.Text, "'", "''") & "%'", , adSearchForward, 1
End If
Sorot
End Sub

Sub TampilkanData()
Dim SelBks
  Dim SQL As String
'  If LookUp = KasOut Then
'    SQL = "SELECT MKasKeluar.*, MKasKeluar.Kredit AS Jumlah FROM MKasKeluar WHERE " & IIf(ckAll.Checked, "", " IsNull(IsAmbil,0)=0 AND ") & " Dari=1 AND UPPER(" & NamaField & ") LIKE '%" & Replace(Text1.Text, "'", "''") & "%'"
'  Else
'    SQL = "SELECT MKasKeluar.*, MKasKeluar.Debet AS Jumlah FROM MKasKeluar WHERE " & IIf(ckAll.Checked, "", " IsNull(IsAmbil,0)=0 AND ") & " Dari=0 AND UPPER(" & NamaField & ") LIKE '%" & Replace(Text1.Text, "'", "''") & "%'"
'  End If
  SQL = "SELECT MTukarPoin.*, MAlamat.Kode AS KdMember, MAlamat.Nama AS NamaMember, MUser.Nama AS Kasir " & vbCrLf & _
        " From MTukarPoin " & vbCrLf & _
        " LEFT JOIN MAlamat ON MAlamat.NoID=MTukarPoin.IDMember " & vbCrLf & _
        " LEFT JOIN MUser ON MUser.NoID=MTukarPoin.IDKasir" & vbCrLf & _
        " WHERE MTukarPoin.IDKassa=" & IDPOSDef & " AND MTukarPoin.Tanggal>='" & Format(Now, "yyyy-MM-dd") & "' AND MTukarPoin.Tanggal<'" & Format(DateAdd("d", 1, Now), "yyyy-MM-dd") & "'"
  Adodc1.RecordSource = SQL
  Adodc1.Refresh
  TDBGrid1.DataSource = Adodc1
  Sorot
End Sub

Sub Sorot()
  Dim SelBks As TrueOleDBGrid60.SelBookmarks
  Set SelBks = TDBGrid1.SelBookmarks
  
  While SelBks.Count
    SelBks.Remove 0
  Wend
  If Not (Adodc1.Recordset.BOF Or Adodc1.Recordset.EOF) Then
    If IsNull(TDBGrid1.Bookmark) Then
      SelBks.Add TDBGrid1.Bookmark
    End If
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'  Set Adodc1 = Nothing
'Save Layouts
  'TDBGrid1.LayoutName = "Layoutku"
  TDBGrid1.LayoutFileName = App.path & "\" & Me.Name & "_" & TDBGrid1.Name & ".grx"
  TDBGrid1.Layouts.Add "Layoutku"
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim idBrg As Long
Dim isSimpanBrg As Boolean
Dim SelBks, SQL As String
Dim Hasil As String
Hasil = Trim(SendByCode(KeyCode))
Select Case Hasil
Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "00", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "'", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "-", "+", "=", "|", "\", ".", ":", ";", "?", "[", "]"
  Text1.Text = Text1.Text & Hasil
  Text1.SelStart = Len(Text1.Text)
  DoEvents
  View
 Case "SPC"
  Text1.Text = Text1.Text & " "
  Text1.SelStart = Len(Text1.Text)
  DoEvents
  View
Case "BKS"
  If Len(Text1.Text) > 0 Then Text1.Text = Left(Text1.Text, Len(Text1.Text) - 1)
  Text1.SelStart = Len(Text1.Text)
  DoEvents
  View
Case "CLR"
  Text1.Text = ""
  View
Case "DN"
  If Not Adodc1.Recordset.BOF Then
    Adodc1.Recordset.MoveNext
    If Adodc1.Recordset.EOF Then
      Adodc1.Recordset.MovePrevious
    End If
    Set SelBks = TDBGrid1.SelBookmarks
    While SelBks.Count <> 0
      SelBks.Remove 0
    Wend
    TDBGrid1.SelBookmarks.Add TDBGrid1.Bookmark
  End If
Case "UP"
  If Not Adodc1.Recordset.BOF Then
    Adodc1.Recordset.MovePrevious
    If Adodc1.Recordset.BOF Then
      Adodc1.Recordset.MoveNext
    End If
    Set SelBks = TDBGrid1.SelBookmarks
    While SelBks.Count <> 0
      SelBks.Remove 0
    Wend
    TDBGrid1.SelBookmarks.Add TDBGrid1.Bookmark
  End If
Case "ENT"
  ambil = True
  NoIDKas = TDBGrid1.Columns("NoID")
  Unload Me
Case "ESC"
  ambil = False
  NoIDKas = 0
  Unload Me
Case "PLU"
  ambil = True
  NoIDKas = TDBGrid1.Columns("NoID")
  Unload Me
End Select
End Sub
Public Sub Tampil(ByRef IsAmbil As Boolean, NoID As Long)
Me.Show 1
IsAmbil = ambil
NoID = NoIDKas
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
