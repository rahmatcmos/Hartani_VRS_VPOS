VERSION 5.00
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dxctrls.dll"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmHapus 
   Caption         =   "Hapus Data Kasir"
   ClientHeight    =   2520
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2520
   ScaleWidth      =   6660
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   465
      Left            =   4800
      TabIndex        =   7
      Top             =   1725
      Width           =   1515
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4260
      Top             =   1950
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1755
      TabIndex        =   2
      Text            =   "C:\Pos"
      Top             =   360
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Proses"
      Height          =   465
      Left            =   3210
      TabIndex        =   0
      Top             =   1725
      Width           =   1515
   End
   Begin TDBDate6Ctl.TDBDate TDBDate2 
      Height          =   315
      Left            =   4035
      TabIndex        =   4
      Top             =   1050
      Width           =   1785
      _Version        =   65536
      _ExtentX        =   3149
      _ExtentY        =   556
      Calendar        =   "frmhapus.frx":0000
      Caption         =   "frmhapus.frx":0118
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmhapus.frx":0184
      Keys            =   "frmhapus.frx":01A2
      Spin            =   "frmhapus.frx":0200
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd-mm-yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "dd-mm-yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "12-08-2011"
      ValidateMode    =   0
      ValueVT         =   1886388231
      Value           =   40767
      CenturyMode     =   0
   End
   Begin TDBDate6Ctl.TDBDate TDBDate1 
      Height          =   315
      Left            =   1755
      TabIndex        =   6
      Top             =   1020
      Width           =   1785
      _Version        =   65536
      _ExtentX        =   3149
      _ExtentY        =   556
      Calendar        =   "frmhapus.frx":0228
      Caption         =   "frmhapus.frx":0340
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmhapus.frx":03AC
      Keys            =   "frmhapus.frx":03CA
      Spin            =   "frmhapus.frx":0428
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd-mm-yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "dd-mm-yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "12-08-2011"
      ValidateMode    =   0
      ValueVT         =   1886388231
      Value           =   40767
      CenturyMode     =   0
   End
   Begin VB.Label lbProses 
      Caption         =   "Tgl Proses : "
      Height          =   585
      Left            =   45
      TabIndex        =   9
      Top             =   1710
      Width           =   2895
   End
   Begin CONTROLSLibCtl.dxProgressBar pb1 
      Height          =   240
      Left            =   135
      TabIndex        =   8
      Top             =   1395
      Width           =   6450
      _Version        =   65536
      _cx             =   11377
      _cy             =   423
      ForeColor       =   0
      BackColor       =   15790320
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MinPos          =   0
      MaxPos          =   100
      Pos             =   0
      Step            =   10
      ShowText        =   -1  'True
      Orientation     =   0
      StartColor      =   16711680
      EndColor        =   16777215
      DrawBorderStyle =   1
      ShowTextStyle   =   0
      DrawBarStyle    =   2
      DrawBarBorderStyle=   2
   End
   Begin VB.Label Label3 
      Caption         =   "s/d"
      Height          =   585
      Left            =   3585
      TabIndex        =   5
      Top             =   1020
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "Tanggal"
      Height          =   585
      Left            =   1005
      TabIndex        =   3
      Top             =   1050
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Path Kasir"
      Height          =   585
      Left            =   990
      TabIndex        =   1
      Top             =   405
      Width           =   1275
   End
End
Attribute VB_Name = "frmHapus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Dir(Text1.Text & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb") <> "" Then
    If MsgBox("Siap Hapus data Kasir antara tanggal :" & Format(TDBDate1, "dd MMMM yyyy") & " sampai dg " & Format(TDBDate2, "dd MMMM yyyy") & "?", vbQuestion + vbYesNo) = vbYes Then
        HapusData
    End If
Else
    MsgBox "Database tidak ditemukan, silahkan cek path aplikasi!", vbCritical
End If
End Sub
Sub HapusData()
Dim dbs As Database
Dim tgl As Date
nmfilebackup = "tempdb" & NamaMesin & Format(Now, "yyyyMMddHHmmss") & ".mdb"
FileCopy Text1.Text & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb", Text1.Text & "\database\" & nmfilebackup
Set dbs = OpenDatabase(Text1.Text & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
pb1.pos = 0
For i = 0 To (TDBDate2 - TDBDate1)
tgl = DateAdd("d", i, TDBDate1)
    lbProses.Caption = "Tgl Proses :" & Format(tgl, "dd-MM-yyyy")
    dbs.Execute ("DELETE MSALESD.* FROM MSALESD inner join msales on msalesd.idsales=msales.noid where msales.tanggal>=#" & Format(tgl, "MM/dd/yyyy") & "# and msales.tanggal<#" & Format(DateAdd("d", 1, tgl), "MM/dd/yyyy") & "#")
    DoEvents
    dbs.Execute ("DELETE msales.* from msales where msales.tanggal>=#" & Format(tgl, "MM/dd/yyyy") & "# and msales.tanggal<#" & Format(DateAdd("d", 1, tgl), "MM/dd/yyyy") & "#")
    DoEvents
    pb1.pos = (i + 1) * 100 / ((TDBDate2 - TDBDate1) + 1)
    DoEvents
Next
pb1.pos = 100
DoEvents
dbs.Close
Set dbs = Nothing
If MsgBox("Proses selesai!" & vbCrLf & "mau kecilkan/mampatkan database?", vbQuestion + vbYesNo) = vbYes Then
    MamPatkanDatabase
End If
End Sub
Sub MamPatkanDatabase()
On Error GoTo pesan
If Dir(DirDatabase & "\CompactTempDB" & Format(Now, "_yyyyMM") & ".mdb") <> "" Then Kill DirDatabase & "\CompactTempDB" & Format(Now, "_yyyyMM") & ".mdb"
CompactDatabase DirDatabase & "\TempDB" & Format(Now, "_yyyyMM") & ".mdb", DirDatabase & "\CompactTempDB" & Format(Now, "_yyyyMM") & ".mdb", , dbEncrypt + dbVersion30
FileCopy DirDatabase & "\CompactTempDB" & Format(Now, "_yyyyMM") & ".mdb", DirDatabase & "\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
MsgBox "Proses Pemampatan selesai!", vbInformation
Exit Sub
pesan:
MsgBox "ada kesalahan!" & vbCrLf & Err.Description, vbCritical
End Sub
Private Sub Command2_Click()
'CommonDialog1.Filter = "Microsoft Access|*.mdb"
'CommonDialog1.ShowOpen
'Text1.Text = CommonDialog1.Filename
Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = App.Path
TDBDate1 = getfirstdate
TDBDate2 = Date - 14
End Sub
Function getfirstdate() As Date
Dim dbs As Database
Dim rs As Recordset
Set dbs = OpenDatabase(Text1.Text & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb")

Set rs = dbs.OpenRecordset("select min(tanggal) as tanggalawal from msales")
If rs.EOF Or rs.BOF Then
    getfirstdate = Date
Else
If IsNull(rs!tanggalawal) Then
getfirstdate = Date
Else
getfirstdate = rs!tanggalawal
End If
End If
rs.Close
Set rs = Nothing
dbs.Close
Set dbs = Nothing
    
End Function
