VERSION 5.00
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dxctrls.dll"
Object = "{153C51AB-7CB7-45EE-AFDE-3B10157651D6}#1.0#0"; "XShow40.ocx"
Begin VB.Form Form2 
   BackColor       =   &H0080C0FF&
   Caption         =   "VPos Ver 12.3.1"
   ClientHeight    =   10755
   ClientLeft      =   -5670
   ClientTop       =   -2340
   ClientWidth     =   15120
   Icon            =   "frmSPLASH.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   10755
   ScaleWidth      =   15120
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3225
      Left            =   10620
      TabIndex        =   14
      Top             =   7560
      Width           =   4485
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   19.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   2070
         TabIndex        =   19
         Top             =   450
         Width           =   2355
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   19.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   2070
         TabIndex        =   18
         Top             =   1080
         Width           =   2355
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   2070
         TabIndex        =   17
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   2070
         TabIndex        =   15
         Text            =   "Y"
         Top             =   2310
         Width           =   855
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Tunggu         Sedang          Proses.................................."
         ForeColor       =   &H80000004&
         Height          =   225
         Left            =   60
         TabIndex        =   25
         Top             =   2910
         Visible         =   0   'False
         Width           =   4185
      End
      Begin VB.Label lbUser 
         BackStyle       =   0  'Transparent
         Caption         =   "USER ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   330
         TabIndex        =   22
         Top             =   600
         Width           =   1185
      End
      Begin VB.Label lbPassword 
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   330
         TabIndex        =   21
         Top             =   1230
         Width           =   1545
      End
      Begin VB.Label lbShift 
         BackStyle       =   0  'Transparent
         Caption         =   "SHIFT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   330
         TabIndex        =   20
         Top             =   1830
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ONLINE (Y/T)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   330
         TabIndex        =   16
         Top             =   2430
         Width           =   1485
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   30000
      Left            =   1410
      Top             =   1770
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   390
      Top             =   2130
   End
   Begin VB.Frame Frame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9255
      Left            =   0
      TabIndex        =   23
      Top             =   1740
      Width           =   15135
      Begin XShow40.XShow XShow1 
         Height          =   9180
         Left            =   -60
         TabIndex        =   24
         Top             =   0
         Width           =   16950
         Object.Visible         =   -1  'True
         BorderStyle     =   2
         Enabled         =   -1  'True
         Cursor          =   2
         BackColor       =   12640511
         BackgroundStyle =   3
         Center          =   -1  'True
         Delay           =   25
         Stretch         =   0   'False
         Proportional    =   0   'False
         Effect          =   0
         Step            =   4
         ClearClientArea =   0   'False
      End
   End
   Begin CONTROLSLibCtl.dxLabel dxLabela 
      DragIcon        =   "frmSPLASH.frx":0442
      Height          =   705
      Left            =   9300
      TabIndex        =   13
      Top             =   60
      Width           =   5775
      _Version        =   0
      _cx             =   10186
      _cy             =   1244
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "VPos ver 12.3.1"
      BackStyle       =   0
      BackColor       =   16710381
      ForeColor       =   12582912
      LabelStyle      =   1
      Label3dStyle    =   1
      Label3dOrientation=   7
      Label3dDepth    =   5
      PenWidth        =   1
      Angle           =   0
      ShadowColor     =   16744576
   End
   Begin CONTROLSLibCtl.dxBackground dxBackground3 
      Height          =   315
      Left            =   810
      TabIndex        =   12
      Top             =   1800
      Width           =   315
      _Version        =   65536
      _cx             =   556
      _cy             =   556
      StartColor      =   192
      EndColor        =   12648447
      ColorFillStyle  =   0
      BackgroundStyle =   0
      DrawPictureStyle=   1
      AnimationInterval=   1000
   End
   Begin CONTROLSLibCtl.dxBackground dxBackground2 
      Height          =   135
      Left            =   0
      TabIndex        =   11
      Top             =   1620
      Width           =   16020
      _Version        =   65536
      _cx             =   28257
      _cy             =   238
      StartColor      =   192
      EndColor        =   12648447
      ColorFillStyle  =   1
      BackgroundStyle =   0
      DrawPictureStyle=   1
      AnimationInterval=   1000
   End
   Begin CONTROLSLibCtl.dxLabel dxLabel2 
      DragIcon        =   "frmSPLASH.frx":0B2C
      Height          =   345
      Left            =   6960
      TabIndex        =   10
      Top             =   540
      Width           =   1635
      _Version        =   0
      _cx             =   2884
      _cy             =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " "
      BackStyle       =   0
      BackColor       =   16710381
      ForeColor       =   12582912
      LabelStyle      =   1
      Label3dStyle    =   1
      Label3dOrientation=   7
      Label3dDepth    =   5
      PenWidth        =   1
      Angle           =   0
      ShadowColor     =   16744576
   End
   Begin CONTROLSLibCtl.dxLabel dxLabel1 
      DragIcon        =   "frmSPLASH.frx":1216
      Height          =   705
      Left            =   7800
      TabIndex        =   9
      Top             =   780
      Width           =   7425
      _Version        =   0
      _cx             =   13097
      _cy             =   1244
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "By : VPoint, Email : vpointindonesia@yahoo.co.id"
      BackStyle       =   0
      BackColor       =   16710381
      ForeColor       =   12582912
      LabelStyle      =   0
      Label3dStyle    =   2
      Label3dOrientation=   7
      Label3dDepth    =   0
      PenWidth        =   1
      Angle           =   0
      ShadowColor     =   16744576
   End
   Begin VB.Label lbJam 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1230
      TabIndex        =   7
      Top             =   810
      Width           =   1335
   End
   Begin VB.Label lbTanggal 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1230
      TabIndex        =   6
      Top             =   1170
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "J A M "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   270
      TabIndex        =   5
      Top             =   810
      Width           =   825
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "TANGGAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   270
      TabIndex        =   4
      Top             =   1170
      Width           =   1095
   End
   Begin VB.Label lbKassa 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1230
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lbStatus 
      BackStyle       =   0  'Transparent
      Caption         =   ": ONLINE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1230
      TabIndex        =   2
      Top             =   480
      Width           =   2475
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "STATUS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   825
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "KASSA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   825
   End
   Begin CONTROLSLibCtl.dxBackground dxBackground1 
      Height          =   1605
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   17685
      _Version        =   65536
      _cx             =   31194
      _cy             =   2831
      StartColor      =   33023
      EndColor        =   12640511
      ColorFillStyle  =   0
      BackgroundStyle =   0
      DrawPictureStyle=   1
      AnimationInterval=   1000
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbs As Database
Dim rs As Recordset
'Dim posisi As Integer
'Dim skala1 As Integer
'Dim skala2 As Integer
Dim NamaImage() As String

Private Sub Form_Activate()
  On Error Resume Next
  Text3.Text = 1
'  If Format(Time, "HHnnss") > "150000" Then
'    Text3.Text = 2
'  Else
'    Text3.Text = 1
'  End If
  
  Timer1.Enabled = True
  Text1.SetFocus

End Sub

Private Sub Form_DeActivate()
  Timer1.Enabled = False
End Sub

Private Sub Form_Load()
  If x.HasilX = Trial Then
    Me.Caption = "TRIAL " & Me.Caption
  End If
  BuatLogApp ("Form Login...Inisialisasi Form Pesan")

  frmProses.Show 1
  BuatLogApp ("Form Login...Form Pesan loaded")
  BuatLogApp ("Form Login...Inisialisasi Background")
  
  setBackGround
  BuatLogApp ("Form Login...Background loaded")
  'frmProses.SetFocus
  DoEvents
  bolehbergerak = True
  isHasilKonversi = False
  Text3.Text = 1
'  If Format(Time, "HHnnss") > "150000" Then
'    Text3.Text = 2
'  Else
'    Text3.Text = 1
'  End If
'skala2 = 100
'  Remover1.HideMouseCursor True
'  Remover1.HideMouseCursor True
'  Remover1.HideMouseCursor True
'
'  Remover1.HideMouseCursor False
'  Remover1.HideMouseCursor False
'  Remover1.HideMouseCursor False
  
  Text2.PasswordChar = Chr(35)
  isHasilKonversi = False
  lbTanggal = ": " & Format(Date, "dd/MM/yyyy")
  BuatLogApp ("Form Login...Inisialisasi Footer Struk")
  
  'Footer
  If Dir(DirDatabase & "\Setting.mdb") = "" Then
    Footer1 = ""
    Footer2 = ""
    Footer3 = ""
  Else
      Set dbs = OpenDatabase(DirDatabase & "\Setting.mdb")
      Set rs = dbs.OpenRecordset("MSetting")
      If rs.EOF And rs.BOF Then
        Footer1 = ""
        Footer2 = ""
        Footer3 = ""
      Else
        Footer1 = NullToStr(rs!Footer1)
        Footer2 = NullToStr(rs!Footer2)
        Footer3 = NullToStr(rs!Footer3)
      End If
      rs.Close
      Set rs = Nothing
      dbs.Close
      Set dbs = Nothing

 End If
BuatLogApp ("Form Login...Footer Struk sukses diset.")
BuatLogApp ("Form Login...Inisialisasi Setting Kasir (ID dan Hardware)")
  If Dir(DirDatabase & "\TempDB" & Format(Now, "_yyyyMM") & ".mdb") = "" Then
    FileCopy DirDatabase & "\TempDB.mdb", DirDatabase & "\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
  End If
  Set dbs = OpenDatabase(DirDatabase & "\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
  Set rs = dbs.OpenRecordset("Umum")
  If rs.EOF And rs.BOF Then
    rs.AddNew
    rs!kassa = "01"
    rs!PersenBiayaKartuKredit = 0
    rs!IsNotaSelesai = True
    rs!IDSalesAkhir = 1
    rs.Update
    lbkassa.Caption = ": 01"
    NamaMesin = "01"
    NamaToko = ""
    NoPortPrinter = 1
    NoPortDisplay = 2
    NoPortBarcode = 3
    NoPortDrawer = 4
isTampilSaldoStock = False
  Else
    NamaMesin = rs!kassa
    lbkassa.Caption = ": " & NamaMesin
    IsNotaSelesai = rs!IsNotaSelesai
    IDNotaTerakhir = rs!IDSalesAkhir
    Judulstruk = NullToStr(rs!Judul)
    PersenBiayaKartuKredit = rs!PersenBiayaKartuKredit
    NamaToko = NullToStr(Trim(rs!Perusahaan))
    SpasiFooter = NullToNol(Trim(rs!SpasiFooter))
    NoPortPrinter = NullToNol(rs!NamaPrinter)
    NoPortBarcode = NullToNol(rs!namaBarcode)
    NoPortDisplay = NullToNol(rs!NamaCustomerDisplay)
    NoPortDrawer = NullToNol(rs!NamaDrawer)
    isTampilSaldoStock = NullToBool(rs!IsCekStock)
    KodeUserDua = getRegistry("Kode", "Pengawas")
    PersenLap = NullToNol(getRegistry("Prosen", "Pengawas"))
    If isOnline Then
        lbStatus = ": " & "ONLINE"
    Else
        lbStatus = ": " & "Local"
    End If
    'NamaToko =  'Trim(Mid(Judulstruk, 1, InStr(1, Judulstruk, Chr(13)) - 1))
'    Label3.Caption = NamaToko
  End If
  
  dbs.Close
  BuatLogApp ("Form Login...Setting Kasir Ok.")
  BuatLogApp ("Form Login...Baca Setting Tombol.")
  
  Set dbs = OpenDatabase(DirDatabase & "\Tombol.mdb")
    
  Set rs = dbs.OpenRecordset("SELECT Code,Kode FROM Tombol")
  If rs.BOF And rs.EOF Then
  Else
    rs.MoveFirst
    Do While Not rs.EOF
      SendByCode(rs!code) = rs!kode
      rs.MoveNext
    Loop
  End If
    rs.Close
  Set rs = Nothing
  dbs.Close
  Set dbs = Nothing

'If Dir(DirUpdate & "\dbmaster.mdb") <> "" Then
'  FileCopy DirUpdate & "\dbmaster.mdb", DirDatabase & "\dbmaster.mdb"
'  Kill DirUpdate & "\dbmaster.mdb"
'End If
' frmProses.Show 1
'  Set DBS = OpenDatabase(DirDatabase & "\dbMaster.mdb")
'  Set RS = DBS.OpenRecordset("MEmp", dbOpenTable)
'  RS.Index = "Kode"
'
BuatLogApp ("Form Login...Setting Tombol OK.")
BuatLogApp ("Form Login...Setting Com Display...")

  Set comDisplay = CreateObject("MSCOMMLIB.MSCOMM")
BuatLogApp ("Form Login...Setting Com Display OK.")
  
  Set comPrinter = CreateObject("MSCOMMLIB.MSCOMM")
BuatLogApp ("Form Login...Buka Port Display dan Printer ...")
  
Set comDrawer = CreateObject("MSCOMMLIB.MSCOMM")
BuatLogApp ("Form Login...Buka Port Drawer  ...")

  BukaPortDisplay
BuatLogApp ("Form Login...Port Display OK.")
  BukaPortPrinter
BuatLogApp ("Form Login...Port Printer OK.")
  BukaPortDrawer
BuatLogApp ("Form Login...Port Drawer OK.")

    If IsNotaSelesai Then
    PrinBigChar Space(15 - Min(15, Len(NamaToko))) & UCase(NamaToko) & Chr(13) & Chr(10)
    Prin Judulstruk & Chr(13) & Chr(10) & "---------------------------------------"
  End If
  Prin " "
 ' Unload frmProses
 ' DoEvents
BuatLogApp ("Form Login...Sukses di load.")
'   Frame1.Left = Me.Width - Frame1.Width
'  Frame1.Top = Me.Height - Frame1.Height

End Sub

Private Sub Form_Unload(Cancel As Integer)
  If MsgBox("Keluar aplikasi?", vbOKCancel + vbQuestion + vbQuestion, "V POS") = vbOK Then
    MamPatkanDatabase
    End
  Else
    Cancel = 1
  End If
End Sub

Private Sub Label1_DblClick()
  Unload Me
End Sub

Private Sub Text1_Change()
    If Len(Text1.Text) = 6 Then
        If UCase(Text1.Text) = "KELUAR" Then Unload Me
'        Text2.SetFocus
    End If
End Sub

Private Sub Text1_DblClick()
    If MsgBox("Keluar aplikasi?", vbOKCancel + vbQuestion, "V POS") = vbOK Then
        End
    End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Ismati As Boolean
Dim Hasil As String
Hasil = Trim(SendByCode(KeyCode))
KeyCode = 0
Select Case Hasil
Case "00", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "*", "`", "!", "@", "#", "$", "%", "^", "&", "<", "(", ")", "-", "+", "{", "}", "[", "]", "/", "?"
  Text1.Text = Text1.Text & Hasil
  Text1.SelStart = Len(Text1.Text)
Case "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"
  Text1.Text = Text1.Text & Hasil
  Text1.SelStart = Len(Text1.Text)
Case "ESC"
  Ismati = False
          frmQuery.Tampil Ismati, "Tutup Program? " & Chr(13) & "Ya (ENTER), Tidak (ESC)."
          If Ismati Then
            MamPatkanDatabase
            'ditutuP DISPLAYPESAN Space(20), Space(20)
            'Call Shell("Rundll32.exe user,exitwindows")
            End
          End If
Case "ENT"
        If isOnline = False Then
          Set dbs = OpenDatabase(DirDatabase & "\dbMaster.mdb")
        Else
          Set dbs = OpenDatabase(DirDbServer & "\dbMaster.mdb")
        End If
        Set rs = dbs.OpenRecordset("MEmp", dbOpenTable)
        rs.Index = "Kode"
        rs.Seek "=", Text1.Text
        If rs.NoMatch Then
        Else
          Text2.SetFocus
          Text2.SelStart = 0
          Text2.SelLength = Len(Text2.Text)
          'Text2.SelText = Text2.Text
          'SendKeys "{HOME}+{END}"
        End If
        rs.Close
        Set rs = Nothing
        dbs.Close
        Set dbs = Nothing

Case "SCN"
    If Text1.Text = "DATA" Then
        frmSetData.Show 1
    End If
Case "SCK"
    If Text1.Text = "DATA" Then
        frmSetDeviceKASIR.Show 1
    End If

Case "CLR"
  Text1.Text = ""
Case "BKS"
  If Len(Text1.Text) > 0 Then
    Text1.Text = Left(Text1.Text, Len(Text1.Text) - 1)
    Text1.SelStart = Len(Text1.Text)
  End If
End Select
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub



Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Hasil As String
Hasil = Trim(SendByCode(KeyCode))
KeyCode = 0
Select Case Hasil
Case "00", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "*", "`", "!", "@", "#", "$", "%", "^", "&", "<", "(", ")", "-", "+", "{", "}", "[", "]", "/", "?"
  Text2.Text = Text2.Text & Hasil
  Text2.SelStart = Len(Text2.Text)
Case "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"
  Text2.Text = Text2.Text & Hasil
  Text2.SelStart = Len(Text2.Text)
Case "ESC", "ENT"
      If isOnline = False Then
        Set dbs = OpenDatabase(DirDatabase & "\dbMaster.mdb")
      Else
        Set dbs = OpenDatabase(DirDbServer & "\dbMaster.mdb")
      End If
      Set rs = dbs.OpenRecordset("MEmp", dbOpenTable)
      rs.Index = "Kode"
      rs.Seek "=", Text1.Text
      If rs.NoMatch Then
        Text1.SetFocus
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1.Text)
        'SendKeys "{HOME}+{END}"
        rs.Close
        Set rs = Nothing
        dbs.Close
        Set dbs = Nothing
        Exit Sub
      Else
        If rs!kode = Text1.Text And rs!Password = Text2.Text Then
        IDUser = rs!NoID
        NamaKasir = rs!Nama
        KodeKasir = Text1.Text
        isSupervisor = rs!isPengawas
        Timer1.Enabled = False
        rs.Close
        Set rs = Nothing
        dbs.Close
        Set dbs = Nothing
'        Text1.Text = ""
'        Text2.Text = ""
        Text1.SetFocus
          If Hasil = "ENT" Then
'            Label9.Visible = True
'            Label9.Refresh
'            DoEvents
'            Form3.Show
'            Form3.SetFocus
'            Label9.Visible = False
'            Me.Hide
'            Exit Sub
            Text3.SetFocus
          ElseIf Hasil = "ESC" And isSupervisor = True Then
            
           ' frmMaintenance.Show 1
          Else
            Text1.SetFocus
            Text1.SelStart = 0
            Text1.SelLength = Len(Text1.Text)
            'SendKeys "{HOME}+{END}"
          End If
        Timer1.Enabled = True
        Else
        Text1.SetFocus
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1.Text)
        'SendKeys "{HOME}+{END}"
        Exit Sub
        End If
      End If
'rs.Close
'Set rs = Nothing
'dbs.Close
'Set dbs = Nothing
    
Case "CLR"
  Text2.Text = ""
Case "BKS"
  If Len(Text2.Text) > 0 Then
    Text2.Text = Left(Text2.Text, Len(Text2.Text) - 1)
    Text2.SelStart = Len(Text2.Text)
  End If
End Select
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
 KeyAscii = 0
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Hasil As String
Hasil = Trim(SendByCode(KeyCode))
KeyCode = 0
Select Case Hasil
Case "1", "2"
  Text3.Text = Hasil
Case "ENT"
    If Text1.Text = "" Or Text3.Text = "" Then Text1.SetFocus: Exit Sub
    If IsNumeric(Text3.Text) Then
        NamaShift = CInt(Text3.Text)
    Else
        Exit Sub
        NamaShift = Text3.Text
    End If
    'Text1.Text = ""
    'Text2.Text = ""
    'Label9.Visible = True
    'Label9.Refresh
    'DoEvents
    'Form3.Show
    'Form3.SetFocus
    'Label9.Visible = False
    'Me.Hide
    'Exit Sub
    Text4.SetFocus
Case "CLR"
  Text3.Text = ""
Case "ESC"
  Text2.SetFocus
Case "BKS"
  If Len(Text3.Text) > 0 Then
    Text3.Text = Left(Text3.Text, Len(Text3.Text) - 1)
    Text3.SelStart = Len(Text3.Text)
  End If
End Select
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Hasil As String
Hasil = Trim(SendByCode(KeyCode))
KeyCode = 0
Select Case Hasil
Case "Y", "T"
  Text4.Text = Hasil
Case "ENT"
      If isOnline = False Then
        Set dbs = OpenDatabase(DirDatabase & "\dbMaster.mdb")
      Else
        Set dbs = OpenDatabase(DirDbServer & "\dbMaster.mdb")
      End If
      Set rs = dbs.OpenRecordset("MEmp", dbOpenTable)
      rs.Index = "Kode"
      rs.Seek "=", Text1.Text
      If rs.NoMatch Then
        Text1.SetFocus
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1.Text)
        rs.Close
        Set rs = Nothing
        dbs.Close
        Set dbs = Nothing
        Exit Sub
      Else
        rs.Index = "Password"
        rs.Seek "=", Text2.Text
        If rs.NoMatch Then
          Text2.SetFocus
          Text2.SelStart = 0
          Text2.SelLength = Len(Text2.Text)
          rs.Close
          Set rs = Nothing
          dbs.Close
          Set dbs = Nothing
          Exit Sub
        End If
      End If
        IDUser = rs!NoID
        NamaKasir = rs!Nama
        KodeKasir = Text1.Text
        isSupervisor = rs!isPengawas
          'Kasus Match
          rs.Close
          Set rs = Nothing
          dbs.Close
          Set dbs = Nothing
    If Text1.Text = "" Or Text3.Text = "" Then Text1.SetFocus: Exit Sub
    If Text4.Text = "y" Or Text4.Text = "Y" Then
      isRemcomendedOnline = True
    Else
      isRemcomendedOnline = False
    End If
    Dim jwb As Boolean
    
    If SudahIsiModal Then
      jwb = True
    Else
        jwb = False
        frmIsiModal.Tampil jwb, 13
    End If
    If jwb Then
      KodeUserLogin = Text1.Text
      Text1.Text = ""
      Text2.Text = ""
      Label9.Visible = True
      Label9.Refresh
      DoEvents
      Form3.Show
      Form3.SetFocus
      Label9.Visible = False
      Me.Hide
      Exit Sub
    End If
  If Not rs Is Nothing Then
    rs.Close
    Set rs = Nothing
  End If
  If Not dbs Is Nothing Then
    dbs.Close
    Set dbs = Nothing
  End If
Case "CLR"
  Text4.Text = ""
Case "ESC"
  Text3.SetFocus
Case "STL"
  If isSupervisor Then
    frmHapus.Show 1
  End If
Case "BKS"
  If Len(Text3.Text) > 0 Then
    Text4.Text = Left(Text4.Text, Len(Text4.Text) - 1)
    Text4.SelStart = Len(Text4.Text)
  End If
End Select
End Sub
Function SudahIsiModal() As Boolean
    Dim dbs As Database
    Dim rs As Recordset
    If Dir(DirDatabase & "\TempDB" & Format(Now, "_yyyyMM") & ".mdb") = "" Then
      FileCopy DirDatabase & "\TempDB.mdb", DirDatabase & "\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
    End If
    Set dbs = OpenDatabase(DirDatabase & "\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
    If IsResetPerKasir Then
      Set rs = dbs.OpenRecordset("Select * from MReset where tanggal=#" & Format(Now, "MM/dd/yyyy") & "# AND IDKasir=" & IDUser & " AND Shift=" & NamaShift, dbOpenDynaset)
    Else
      Set rs = dbs.OpenRecordset("Select * from MReset where tanggal=#" & Format(Now, "MM/dd/yyyy") & "# AND Shift=" & NamaShift, dbOpenDynaset)
    End If
    If rs.EOF And rs.BOF Then
     SudahIsiModal = False
    Else
     SudahIsiModal = True
    End If
    rs.Close
    Set rs = Nothing
    dbs.Close
    Set dbs = Nothing
End Function
Private Sub Text4_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Timer1_Timer()
  lbJam = ": " & Format(Time, "HH:mm:ss")
' Timer1.Enabled = False
'    posisi = ((posisi + 1) Mod 20)
    If bolehbergerak Then
        'ditutuP DISPLAYPESAN Mid(NamaToko & "  " & NamaToko & "  " & "  " & NamaToko & "  ", 1 + posisi, 20), Mid("KASSA : " & NamaMesin & " , ^" & NamaKasir & "^   " & "KASSA : " & NamaMesin & " , ^" & NamaKasir & "^   " & "KASSA : " & NamaMesin & " , ^" & NamaKasir & "^   ", 1 + posisi, 20)
    End If
'  If Shape3.FillColor >= 16777215 - 1927 Then
'    skala1 = -1927
'  ElseIf Shape3.FillColor <= 1928 Then
'    skala1 = 1927
'  End If
'  Shape3.FillColor = Shape3.FillColor + skala1
'  Text1.BackColor = 16777215 - Shape3.FillColor
'  Text2.BackColor = 16777215 - Shape3.FillColor
'  Text3.BackColor = 16777215 - Shape3.FillColor
'  lbUser.ForeColor = 16777215 - Shape3.FillColor
'  lbPassword.ForeColor = 16777215 - Shape3.FillColor
'  lbShift.ForeColor = 16777215 - Shape3.FillColor
  'MsgBox CStr(Shape3.FillColor)
'  Shape3.Refresh
  DoEvents
'   If dxLabel1.Left >= 4700 Then
'    skala2 = -100
'  ElseIf dxLabel1.Left <= 2100 Then
'    skala2 = 100
'  End If
'  dxLabel1.Move dxLabel1.Left + skala2, dxLabel1.Top
'  dxLabel2.Move dxLabel2.Left + skala2, dxLabel2.Top
  End Sub

Private Sub Timer2_Timer()
  'lbStatus = ": " & GetStatusNetwork
 Timer2.Enabled = False
End Sub

Sub MamPatkanDatabase()
On Error GoTo Trace
If Dir(DirDatabase & "\CompactTempDB" & Format(Now, "_yyyyMM") & ".mdb") <> "" Then Kill DirDatabase & "\CompactTempDB" & Format(Now, "_yyyyMM") & ".mdb"
CompactDatabase DirDatabase & "\TempDB" & Format(Now, "_yyyyMM") & ".mdb", DirDatabase & "\CompactTempDB" & Format(Now, "_yyyyMM") & ".mdb", , dbEncrypt + dbVersion30
FileCopy DirDatabase & "\CompactTempDB" & Format(Now, "_yyyyMM") & ".mdb", DirDatabase & "\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
Trace:
  If Err.Number <> 0 Then
    BuatLogApp Err.Number & ", " & Err.Description
    Err.Clear
  End If
End Sub

Sub setBackGround()
  Dim fz As String
  XShow1.Unlock ("demo")
  fz = Dir(App.path & "\image\background\")
  If fz <> "" Then
    ReDim NamaImage(1)
    NamaImage(1) = App.path & "\image\background\" & fz
    XShow1.Delay = 100
    XShow1.Step = 3
'    XShow1.Stretch = True
    XShow1.Effect = Rnd * 128
    DoEvents
'    MsgBox CStr(XShow1.Effect)
    XShow1.LoadImage NamaImage(Rnd * UBound(NamaImage))
    XShow1.Go
  End If
  fz = Dir
  Do While fz <> ""
    ReDim Preserve NamaImage(UBound(NamaImage) + 1)
    NamaImage(UBound(NamaImage)) = App.path & "\image\background\" & fz
    fz = Dir
  Loop
End Sub

Private Sub XShow1_OnTransitionCompleted()
  Dim x
  DoEvents
'    XShow1.Stretch = True
  XShow1.Delay = 100
  XShow1.Step = 3
  XShow1.Effect = Rnd * 127
  DoEvents
'    MsgBox CStr(XShow1.Effect)
    x = NamaImage(CInt(Rnd * UBound(NamaImage)))
    If x <> "" Then
    XShow1.LoadImage x
    Else
      XShow1.LoadImage NamaImage(1)
    End If
    XShow1.Go
    DoEvents
End Sub
