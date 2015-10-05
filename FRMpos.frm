VERSION 5.00
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dxctrls.dll"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{E994B1F7-F7D0-11D6-A2A1-0010DC1D796E}#19.0#0"; "smbutton.ocx"
Begin VB.Form Form3 
   BackColor       =   &H00FEFAED&
   Caption         =   "VPos ver 13.2.2"
   ClientHeight    =   10455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   Icon            =   "FRMpos.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   10455
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame MenuReedemPoin 
      Caption         =   "Menu Reedem Poin"
      Height          =   3825
      Left            =   3210
      TabIndex        =   71
      Top             =   3300
      Visible         =   0   'False
      Width           =   10365
      Begin VB.Data DataReedemPoin 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   885
         Left            =   4710
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   870
         Visible         =   0   'False
         Width           =   4005
      End
      Begin TrueDBGrid60.TDBGrid TDBGrid3 
         Bindings        =   "FRMpos.frx":146AA
         Height          =   3525
         Left            =   60
         OleObjectBlob   =   "FRMpos.frx":146C7
         TabIndex        =   72
         Top             =   210
         Width           =   9195
      End
      Begin SMButton.Button cmdPilihReedem 
         Height          =   585
         Left            =   9300
         TabIndex        =   73
         Top             =   480
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   1032
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   2
         BackOver        =   16777215
         BorderColor     =   7617536
         BorderColorMedium=   8323072
         ForeHighlightEvent=   1
         ShowFocus       =   -1  'True
         Caption         =   "&Pilih"
         BeTransparentColor=   16646398
      End
      Begin SMButton.Button cmdCancelReedem 
         Height          =   585
         Left            =   9300
         TabIndex        =   74
         Top             =   1080
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   1032
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   2
         BackOver        =   16777215
         BorderColor     =   7617536
         BorderColorMedium=   8323072
         ForeHighlightEvent=   1
         ShowFocus       =   -1  'True
         Caption         =   "&Batal"
         BeTransparentColor=   16646398
      End
   End
   Begin SMButton.Button cmdReedemPoin 
      Height          =   495
      Left            =   90
      TabIndex        =   70
      Top             =   6930
      Visible         =   0   'False
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   873
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      BackOver        =   16777215
      BorderColor     =   7617536
      BorderColorMedium=   8323072
      ForeHighlightEvent=   1
      ShowFocus       =   -1  'True
      Caption         =   "&Reedem Poin"
      BeTransparentColor=   16646398
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\JOB\KassaWin\Database\TempDb.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   885
      Left            =   3570
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "MSalesD"
      Top             =   4560
      Visible         =   0   'False
      Width           =   4005
   End
   Begin SMButton.Button cmdDEL 
      Height          =   345
      Left            =   8820
      TabIndex        =   10
      Top             =   10065
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   609
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      BackOver        =   16777215
      BorderColor     =   7617536
      BorderColorMedium=   8323072
      CaptionEffectEvent=   6
      ForeHighlightEvent=   1
      ShowFocus       =   -1  'True
      Caption         =   "F11 : DEBET"
      BeTransparentColor=   16646398
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   48.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2985
      Locked          =   -1  'True
      TabIndex        =   37
      Text            =   "CHG#          91,350"
      Top             =   1215
      Width           =   12120
   End
   Begin SMButton.Button cmdF1 
      Height          =   345
      Left            =   945
      TabIndex        =   2
      Top             =   10065
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   609
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      BackOver        =   16777215
      BorderColor     =   7617536
      BorderColorMedium=   8323072
      CaptionEffectEvent=   6
      ForeHighlightEvent=   1
      ShowFocus       =   -1  'True
      Caption         =   "F2 : SCN"
      BeTransparentColor=   16646398
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   4110
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   48.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2985
      Locked          =   -1  'True
      TabIndex        =   36
      Text            =   "12345678901234567890"
      Top             =   120
      Width           =   12120
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6150
      Top             =   7380
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Data DataLokal 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   405
      Left            =   4110
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1650
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   90
      TabIndex        =   0
      Top             =   8595
      Width           =   10635
   End
   Begin VB.Timer Timer3 
      Interval        =   30000
      Left            =   3660
      Top             =   1740
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   3990
      Top             =   1290
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   8700
      Top             =   750
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      NullDiscard     =   -1  'True
      RThreshold      =   1
      RTSEnable       =   -1  'True
      SThreshold      =   1
      InputMode       =   1
   End
   Begin SMButton.Button cmdF2 
      Height          =   345
      Left            =   1905
      TabIndex        =   3
      Top             =   10065
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   609
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      BackOver        =   16777215
      BorderColor     =   7617536
      BorderColorMedium=   8323072
      CaptionEffectEvent=   6
      ForeHighlightEvent=   1
      ShowFocus       =   -1  'True
      Caption         =   "F3 : RQTY"
      BeTransparentColor=   16646398
   End
   Begin SMButton.Button cmdF3 
      Height          =   345
      Left            =   2865
      TabIndex        =   4
      Top             =   10065
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   609
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      BackOver        =   16777215
      BorderColor     =   7617536
      BorderColorMedium=   8323072
      CaptionEffectEvent=   6
      ForeHighlightEvent=   1
      ShowFocus       =   -1  'True
      Caption         =   "F4 : VOD"
      BeTransparentColor=   16646398
   End
   Begin SMButton.Button cmdF11 
      Height          =   345
      Left            =   3765
      TabIndex        =   5
      Top             =   10065
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   609
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      BackOver        =   16777215
      BorderColor     =   7617536
      BorderColorMedium=   8323072
      CaptionEffectEvent=   6
      ForeHighlightEvent=   1
      ShowFocus       =   -1  'True
      Caption         =   "F5 : AVD"
      BeTransparentColor=   16646398
   End
   Begin SMButton.Button cmdF12 
      Height          =   345
      Left            =   4575
      TabIndex        =   6
      Top             =   10065
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   609
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      BackOver        =   16777215
      BorderColor     =   7617536
      BorderColorMedium=   8323072
      CaptionEffectEvent=   6
      ForeHighlightEvent=   1
      ShowFocus       =   -1  'True
      Caption         =   "F6 : RTN"
      BeTransparentColor=   16646398
   End
   Begin SMButton.Button cmdPGDN 
      Height          =   345
      Left            =   5385
      TabIndex        =   7
      Top             =   10065
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   609
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      BackOver        =   16777215
      BorderColor     =   7617536
      BorderColorMedium=   8323072
      CaptionEffectEvent=   6
      ForeHighlightEvent=   1
      ShowFocus       =   -1  'True
      Caption         =   "F7 : MBR"
      BeTransparentColor=   16646398
   End
   Begin SMButton.Button cmdPGUP 
      Height          =   345
      Left            =   6150
      TabIndex        =   8
      Top             =   10065
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   609
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      BackOver        =   16777215
      BorderColor     =   7617536
      BorderColorMedium=   8323072
      CaptionEffectEvent=   6
      ForeHighlightEvent=   1
      ShowFocus       =   -1  'True
      Caption         =   "F8 : NS"
      BeTransparentColor=   16646398
   End
   Begin SMButton.Button cmdHOME 
      Height          =   345
      Left            =   6915
      TabIndex        =   9
      Top             =   10065
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   609
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      BackOver        =   16777215
      BorderColor     =   7617536
      BorderColorMedium=   8323072
      CaptionEffectEvent=   6
      ForeHighlightEvent=   1
      ShowFocus       =   -1  'True
      Caption         =   "F9 : RPT"
      BeTransparentColor=   16646398
   End
   Begin SMButton.Button cmdALT 
      Height          =   345
      Left            =   9870
      TabIndex        =   11
      Top             =   10065
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   609
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      BackOver        =   16777215
      BorderColor     =   7617536
      BorderColorMedium=   8323072
      CaptionEffectEvent=   6
      ForeHighlightEvent=   1
      ShowFocus       =   -1  'True
      Caption         =   "F12 : CC"
      BeTransparentColor=   16646398
   End
   Begin SMButton.Button cmdF6 
      Height          =   345
      Left            =   45
      TabIndex        =   1
      Top             =   10065
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   609
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      BackOver        =   16777215
      BorderColor     =   7617536
      BorderColorMedium=   8323072
      CaptionEffectEvent=   6
      ForeHighlightEvent=   1
      ShowFocus       =   -1  'True
      Caption         =   "F1 : SCK"
      BeTransparentColor=   16646398
   End
   Begin SMButton.Button cmdNone 
      Height          =   345
      Left            =   7845
      TabIndex        =   12
      Top             =   10065
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   609
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      BackOver        =   16777215
      BorderColor     =   7617536
      BorderColorMedium=   8323072
      CaptionEffectEvent=   6
      ForeHighlightEvent=   1
      ShowFocus       =   -1  'True
      Caption         =   "F10: VCR"
      BeTransparentColor=   16646398
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Bindings        =   "FRMpos.frx":18221
      Height          =   5700
      Left            =   2220
      OleObjectBlob   =   "FRMpos.frx":18235
      TabIndex        =   75
      Top             =   2580
      Width           =   12840
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reedem Poin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   10950
      TabIndex        =   78
      Top             =   9105
      Width           =   1290
   End
   Begin VB.Label lblReedemPoin 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   12450
      TabIndex        =   77
      Top             =   9105
      Width           =   225
   End
   Begin VB.Label lblNilaiReedem 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   12870
      TabIndex        =   76
      Top             =   9099
      Width           =   1755
   End
   Begin VB.Label lblTotalPoin 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   270
      TabIndex        =   69
      Top             =   6630
      Width           =   1785
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Poin :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   90
      TabIndex        =   68
      Top             =   6300
      Width           =   1875
   End
   Begin VB.Label lbPDP 
      BackStyle       =   0  'Transparent
      Caption         =   "textPDP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   90
      TabIndex        =   67
      Top             =   8355
      Visible         =   0   'False
      Width           =   10545
   End
   Begin VB.Label lbKomisiRp 
      BackStyle       =   0  'Transparent
      Caption         =   "Komisi Rp. :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   90
      TabIndex        =   66
      Top             =   5700
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lbKomisiProsen 
      BackStyle       =   0  'Transparent
      Caption         =   "Komisi (%) :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   90
      TabIndex        =   65
      Top             =   5280
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lbSopir 
      BackStyle       =   0  'Transparent
      Caption         =   "Grup :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   60
      TabIndex        =   64
      Top             =   4890
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dsc Member"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   10950
      TabIndex        =   63
      Top             =   8573
      Width           =   1170
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Ttl"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   10950
      TabIndex        =   62
      Top             =   8310
      Width           =   675
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   10950
      TabIndex        =   61
      Top             =   8836
      Width           =   480
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kembali"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   10950
      TabIndex        =   60
      Top             =   9888
      Width           =   795
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Stl dpt Poin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   10950
      TabIndex        =   59
      Top             =   10155
      Width           =   1185
   End
   Begin VB.Label lbkassa 
      BackStyle       =   0  'Transparent
      Caption         =   "KASSA : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   90
      TabIndex        =   58
      Top             =   2820
      Width           =   2325
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2205
      Left            =   90
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2850
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Jam       : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   90
      TabIndex        =   57
      Top             =   4470
      Width           =   915
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   90
      TabIndex        =   56
      Top             =   4035
      Width           =   915
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   60
      TabIndex        =   55
      Top             =   2505
      Width           =   1815
   End
   Begin VB.Label lbRounding 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   7095
      TabIndex        =   54
      Top             =   9765
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rp."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   8070
      TabIndex        =   53
      Top             =   9750
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pembulatan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   6855
      TabIndex        =   52
      Top             =   9720
      Visible         =   0   'False
      Width           =   1155
   End
   Begin CONTROLSLibCtl.dxBackground dxBackground 
      Height          =   90
      Left            =   30
      TabIndex        =   51
      Top             =   2370
      Width           =   15300
      _Version        =   65536
      _cx             =   26987
      _cy             =   159
      StartColor      =   65535
      EndColor        =   16777215
      ColorFillStyle  =   1
      BackgroundStyle =   0
      DrawPictureStyle=   1
      AnimationInterval=   1000
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      Height          =   10530
      Left            =   0
      Top             =   0
      Width           =   15180
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00728253&
      Height          =   5505
      Left            =   30
      Top             =   2790
      Width           =   2205
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "STL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   14625
      TabIndex        =   50
      Top             =   8310
      Width           =   435
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kurang"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   270
      Left            =   15480
      TabIndex        =   49
      Top             =   4455
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rp."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   270
      Left            =   16650
      TabIndex        =   48
      Top             =   4455
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lbHutang 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   16770
      TabIndex        =   47
      Top             =   4485
      Visible         =   0   'False
      Width           =   1965
   End
   Begin CONTROLSLibCtl.dxLabel dxLabel1 
      DragIcon        =   "FRMpos.frx":23533
      Height          =   585
      Left            =   3960
      TabIndex        =   46
      Top             =   240
      Visible         =   0   'False
      Width           =   5835
      _Version        =   0
      _cx             =   10292
      _cy             =   1032
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "VPOS VERSI 11.1"
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
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rp."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   12240
      TabIndex        =   45
      Top             =   10155
      Width           =   315
   End
   Begin VB.Label lbPoin 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   12870
      TabIndex        =   44
      Top             =   10155
      Width           =   1755
   End
   Begin VB.Label lbNama 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   7335
      TabIndex        =   43
      Top             =   9720
      Width           =   3255
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Disc Item"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   15090
      TabIndex        =   42
      Top             =   6510
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rp."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   12240
      TabIndex        =   41
      Top             =   8573
      Width           =   315
   End
   Begin VB.Label lbDiscBrg 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   15435
      TabIndex        =   40
      Top             =   5760
      Visible         =   0   'False
      Width           =   1965
   End
   Begin CONTROLSLibCtl.dxLabel dxLabel2 
      DragIcon        =   "FRMpos.frx":23C1D
      Height          =   345
      Left            =   90
      TabIndex        =   35
      Top             =   870
      Visible         =   0   'False
      Width           =   5655
      _Version        =   0
      _cx             =   9975
      _cy             =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      BackStyle       =   0
      BackColor       =   16710381
      ForeColor       =   12582912
      LabelStyle      =   1
      Label3dStyle    =   0
      Label3dOrientation=   7
      Label3dDepth    =   5
      PenWidth        =   1
      Angle           =   0
      ShadowColor     =   16744576
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   10950
      TabIndex        =   34
      Top             =   9362
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rp."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   12240
      TabIndex        =   33
      Top             =   9362
      Width           =   315
   End
   Begin VB.Label lbVoucher 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   12870
      TabIndex        =   32
      Top             =   9362
      Width           =   1755
   End
   Begin VB.Label lbQTY 
      BackStyle       =   0  'Transparent
      Caption         =   "1 X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   60
      TabIndex        =   31
      Top             =   9675
      Width           =   3255
   End
   Begin VB.Label lbKartu 
      BackStyle       =   0  'Transparent
      Caption         =   "Silahkan Gesek Kartu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3270
      TabIndex        =   30
      Top             =   9735
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label LbDisc 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   12870
      TabIndex        =   29
      Top             =   8573
      Width           =   1755
   End
   Begin VB.Label lbSubTtl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   12870
      TabIndex        =   28
      Top             =   8310
      Width           =   1755
   End
   Begin VB.Label lbDibayar 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   12870
      TabIndex        =   27
      Top             =   9625
      Width           =   1755
   End
   Begin VB.Label lbTotal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   12870
      TabIndex        =   26
      Top             =   8836
      Width           =   1755
   End
   Begin VB.Label lbKembali 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   12870
      TabIndex        =   25
      Top             =   9888
      Width           =   1755
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      Height          =   2280
      Left            =   30
      Top             =   75
      Width           =   15135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00728253&
      Height          =   1695
      Left            =   45
      Top             =   8325
      Width           =   10755
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00728253&
      BorderColor     =   &H00728253&
      FillColor       =   &H00728253&
      Height          =   2175
      Left            =   0
      Top             =   8280
      Width           =   15060
   End
   Begin VB.Label lbJam 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   1020
      TabIndex        =   24
      Top             =   4500
      Width           =   1815
   End
   Begin VB.Label lbTanggal 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   1020
      TabIndex        =   23
      Top             =   4050
      Width           =   1725
   End
   Begin VB.Label lbStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Status  : Online"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   90
      TabIndex        =   22
      Top             =   3225
      Width           =   2325
   End
   Begin VB.Label lbKasir 
      BackStyle       =   0  'Transparent
      Caption         =   "Kasir   :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   90
      TabIndex        =   21
      Top             =   3630
      Width           =   2565
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rp."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   12240
      TabIndex        =   20
      Top             =   9888
      Width           =   315
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rp."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   12240
      TabIndex        =   19
      Top             =   8836
      Width           =   315
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rp."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   12240
      TabIndex        =   18
      Top             =   9625
      Width           =   315
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rp."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   12240
      TabIndex        =   17
      Top             =   8310
      Width           =   315
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rp."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   15000
      TabIndex        =   16
      Top             =   6240
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label lbBintang 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   14715
      TabIndex        =   15
      Top             =   8580
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dibayar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   10950
      TabIndex        =   14
      Top             =   9625
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   555
      Left            =   4050
      TabIndex        =   13
      Top             =   3060
      Width           =   1275
   End
   Begin CONTROLSLibCtl.dxBackground dxBackground1 
      Height          =   7425
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   15390
      _Version        =   65536
      _cx             =   27146
      _cy             =   13097
      StartColor      =   33023
      EndColor        =   12640511
      ColorFillStyle  =   0
      BackgroundStyle =   0
      DrawPictureStyle=   1
      AnimationInterval=   1000
   End
   Begin CONTROLSLibCtl.dxBackground dxBackground3 
      Height          =   4410
      Left            =   0
      TabIndex        =   39
      Top             =   7320
      Width           =   15540
      _Version        =   65536
      _cx             =   27411
      _cy             =   7779
      Picture         =   "FRMpos.frx":24307
      StartColor      =   12640511
      EndColor        =   33023
      ColorFillStyle  =   0
      BackgroundStyle =   0
      DrawPictureStyle=   1
      AnimationInterval=   1000
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim posisi As Integer
Dim CountReprint As Long
Dim ISCreditCard As Boolean
Dim BiayaCC As Double
Dim dbs As Database
Dim dbSale As Database
Dim rs As Recordset
Dim IDSales As Long
Dim IDSalesD As Long
Dim idSalesDVoucher As Long
Dim Nama As String
Dim NoNota As String
Dim Barcode As String
Dim BarcodeIn As String
Dim IdInv As Long
Dim IDSat As Long
Dim Qty As Double
Dim Konversi As Double
Dim SubTotal As Long
Dim Agen As String
Dim KomisiProsen As Double
Dim KomisiRp As Double
Dim Disc As Long
Dim JumDiscInternRp As Long
Dim DiscINTERNNOTA As Long
Dim PotonganPembulatan As Long
Dim RoundingBawah As Double
Dim SaldoHutang As Long
Dim SaldoStock As Double
Dim Voucher As Long
Dim NilaiVoucher As Double
Dim Total As Long
Dim Dibayar As Long
Dim Kembali As Long
Dim Bank As Long
Dim IDBank As Integer
Dim IDJenisKartu As Integer
Dim IDBankServer As Integer
Dim NoAcc As String
Dim KodeBank As String
Dim NamaBank As String
Dim NamaJenisKartu As String
Dim ChargeBank As Double
Dim IsNotaDariPending As Boolean
Dim Ditutup As Boolean
Dim jawab As Boolean
Dim crdcard As Boolean
Dim Jumlahitem As Double
Dim MacamItem As Double
Dim isBarcode As Boolean
Dim HargaJualKhusus As Double
Dim KodeKhusus As String
Dim AllowBarcode As Boolean
Dim skala1 As Integer
Dim skala2 As Integer
Dim DiscInternRp As Double
Dim DiscInternProsen As Double

Dim DiscRp As Double
Dim DiscProsen As Double
Dim HargaBruto As Double
Dim BelanjaPoin As Double
Dim TotalDiscBrg As Long
Dim AmbilAngka As Double
Dim IsDiscBySupplier As Boolean
Dim IsBarangMember As Boolean
Dim IDSalesAHS As Long

Dim interval As Integer
Dim JumlahBKP As Double
Dim JumlahDPP As Double
Dim JumlahPPN As Double
Dim defNilaiDiskonMember As Double
Dim defBelanjadapat1Poin As Double
Dim defMinialBelanjadapatDiskon As Double
Dim defMinialBelanjadapatDiskon2 As Double
Dim defMinialBelanjaDapatPDP As Double
Dim defIsCCDapatDiskon As Boolean
Public IskartuKredit As Boolean
Dim PoinNotaIni As Double
Dim AmbilVoucher As Double
Dim IDPenerbitVoucher As Integer
Dim qtyVcr As Integer
Dim KodePenerbitVoucher As String
'Dim DiscProsenBawah As Double
'Dim DiscRupiahBawah As Double

Dim IDReedemPoin As Long
Dim ReedemPoin As Long
Dim ReedemNilai As Double

Dim KodeSearch As String
Sub TampilBank(ByVal IsFilterCC As Boolean)
Dim jawaban As Boolean
Dim bayarCC As Long
Dim jwb As Boolean
Dim TasA As Double
Dim TasB As Double
Dim TasC As Double
Dim TasD As Double
Dim Stl As Double
Dim rsStl As Recordset
'dim
    UpdateFooter
    Set rsStl = dbSale.OpenRecordset("SELECT SUM(MSalesD.Jumlah) AS Subtotal FROM MSalesD WHERE MSalesD.IsPoin=True AND MSalesD.IDSales=" & IDSales)
    If Not (rsStl.EOF Or rsStl.BOF) Then
      Stl = NullToNol(rsStl!SubTotal)
    Else
      Stl = 0
    End If
    Set rsStl = Nothing
    If Ditutup = False And lbBintang.Visible Then
    frmBank.Tampil jawaban, IDBank, IDBankServer, NoAcc, KodeBank, NamaBank, ChargeBank, 67, (Total - Dibayar), BiayaCC, bayarCC, IsFilterCC, IDJenisKartu, NamaJenisKartu, IIf(IDMember > 0 And (Stl >= defMinialBelanjadapatDiskon), True, False)
    If jawaban Then
        Bank = bayarCC ' Total - Dibayar
        Dibayar = Total
        TampilBawah
        If Kembali >= 0 Then Ditutup = True
          If NoPortDrawer <> -2 Then
            openDrawer
          End If
        DoEvents
        HitungItem
        DisplayPesan "PAY# " & Space(15 - Min(Len(Format(Dibayar + BiayaCC - Voucher - ReedemNilai, "###,###,##0")), 15)) & Format(Dibayar + BiayaCC - Voucher - ReedemNilai, "###,###,##0"), "CHG# " & Space(15 - Min(Len(Format(Kembali, "###,###,##0")), 15)) & Format(Kembali, "###,###,##0")
        If Kembali >= 0 Then
             Prin "---------------------------------------" & Chr(13) & Chr(10) & _
              "Subtotal" & Space(31 - Len(Format(SubTotal, "###,###,##0"))) & Format(SubTotal, "###,###,##0") & Chr(13) & Chr(10) & _
              IIf(TotalDiscBrg + JumDiscInternRp = 0, "", "Discount Barang" & Space(24 - Len(Format(TotalDiscBrg + JumDiscInternRp, "###,###,##0"))) & Format(TotalDiscBrg + JumDiscInternRp, "###,###,##0") & Chr(13) & Chr(10)) & _
              IIf(Disc + DiscINTERNNOTA + PotonganPembulatan = 0, "", "Discount Nota" & Space(26 - Len(Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0"))) & Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0") & Chr(13) & Chr(10)) & _
              "Voucher " & Space(31 - Len(Format(Voucher, "###,###,##0"))) & Format(Voucher, "###,###,##0") & Chr(13) & Chr(10) & _
              IIf(BiayaCC > 0, "Biaya CC" & Space(31 - Len(Format(BiayaCC, "###,###,##0"))) & Format(BiayaCC, "###,###,##0") & Chr(13) & Chr(10), "") & _
              "Dibayar " & Space(31 - Len(Format(Dibayar + BiayaCC - Voucher - ReedemNilai, "###,###,##0"))) & Format(Dibayar + BiayaCC - Voucher - ReedemNilai, "###,###,##0") & Chr(13) & Chr(10) & _
              "Kembali " & Space(31 - Len(Format(Kembali, "###,###,##0"))) & Format(Kembali, "###,###,##0") & Chr(13) & Chr(10) & _
              IIf(ISSembunyikanFooterbarangKSB, "", "Belanja Barang POIN" & Space(20 - Len(Format(BelanjaPoin, "###,###,##0"))) & Format(BelanjaPoin, "###,###,##0") & Chr(13) & Chr(10)) & _
              "#" & Format(CLng(NoNota), "0#####") & "," & "Kasir" & NamaMesin & ":" & Mid(NamaKasir, 1, 6) & Space(6 - IIf(Len(NamaKasir) <= 6, Len(NamaKasir), 6)) & Format(Date, " dd/MM/yy") & Format(Time, "hh:mm:ss") & Chr(13) & Chr(10) & _
              "Macam Barang =" & Format(MacamItem, "##0") & " , Total Item =" & Format(Jumlahitem, "##0") & Chr(13) & Chr(10) & _
              IIf(KodeMember = "", "", "CUSTOMER: " & KodeMember & "-" & NamaMember & Chr(13) & Chr(10)) & _
              "----TERIMA KASIH ATAS KUNJUNGAN ANDA----" & Chr(13) & Chr(10) & _
              IIf(Footer1 = "", "", Footer1 & Chr(13) & Chr(10)) & _
              IIf(Footer2 = "", "", Footer2 & Chr(13) & Chr(10)) & _
              IIf(Footer3 = "", "", Footer3 & Chr(13) & Chr(10)) & _
              Chr(13) & Chr(10)
              If IsHematKertas Then
                    PrinBigChar Space(15 - Min(15, Len(NamaToko))) & UCase(NamaToko)
                    Prin Judulstruk & Chr(13) & Chr(10) & "---------------------------------------"
              End If
            papercut
        lbBintang.Visible = False
        cmdReedemPoin.Visible = False
        TutupCommBarcode
        Else
          Prin "---------------------------------------" & Chr(13) & Chr(10) & "Dibayar " & Space(31 - Len(Format(Dibayar - Voucher, "###,###,##0"))) & Format(Dibayar - Voucher, "###,###,##0") & Chr(13) & Chr(10) & _
               "Kembali " & Space(31 - Len(Format(Kembali, "###,###,##0"))) & Format(Kembali, "###,###,##0")
        End If 'kembali>0
        DoEvents
        KirimKeServerBeginTrans IDSales
        CetakStruck IDSales, False
        CetakStruckReedem2 IDSales, False
        If IsPakaiKantong Then
          jwb = False
          frmPayment2.Tampil jwb, TasA, TasB, TasC, TasD
          If Not jwb Then
'            Exit Sub
          Else
            Dim rsSales As Recordset
            Set rsSales = dbSale.OpenRecordset("Select * FROM MSales Where NoID=" & IDSales)
            If Not (rsSales.EOF And rsSales.BOF) Then
              rsSales.Edit
              rsSales!TasKresekA = TasA
              rsSales!TasKresekB = TasB
              rsSales!TasKresekC = TasC
              rsSales!TasKresekD = TasD
              rsSales.Update
            End If
            rsSales.Close
            Set rsSales = Nothing
          End If
        End If
        Text1.Text = ""
        
      End If 'jawaban
    End If ' If Ditutup = False And lbBintang.Visible Then
End Sub

Private Sub cmdCancelReedem_Click()
  IDReedemPoin = 0
  Dibayar = Dibayar - ReedemNilai
  ReedemPoin = 0
  ReedemNilai = 0
  TampilBawah
  Text1.Enabled = True
  MenuReedemPoin.Visible = False
  Text1.SetFocus
End Sub

Private Sub cmdDEL_Click()
  Dim IsAmbil As Boolean
  'JIKA KOSONG BERARTI AMBIL PENDINGAN
    If lbBintang.Visible Or Ditutup Then Exit Sub
      If Data1.Recordset.EOF And Data1.Recordset.BOF Then
        Dim rspending As Recordset
        Set rspending = dbSale.OpenRecordset("Select * FROM MSAles Where IsPending=True")
        If rspending.EOF And rspending.BOF Then
          frmPesan.Show 1
        Else
          Dim idPending As Long
           IsAmbil = False
           frmPending.Tampil idPending, IsAmbil
            If IsAmbil Then
            'REVISI 1 NOPEMBER 2011 DI SUBUH HARI
            'PEMBETULAN BUG PENYEBAB NOTA LNCAT
            'HAPUS DULU NOTA TERAKHIR YG BELUM ADA ITEM
                dbSale.Execute "Delete from MSales where Subtotal=0 and NoID=" & IDSales
              'Baru ambil nota Pending
            IsNotaDariPending = True
              IDSales = idPending
              'Data1.RecordSource = "SELECT MSalesD.HargaBruto,MSalesD.DiscRp,MSalesD.DiscProsen,   MSalesD.NoID, MSalesD.IDSales, MSalesD.Transaksi, MSalesD.IDInventor, MSalesD.Qty, MSalesD.Harga, MSalesD.KodeInv, MSalesD.NamaInv, MSalesD.Satuan, MSalesD.Barcode, MSalesD.IDSatuan, MSalesD.Konversi, MSalesD.HargaPokok, MSalesD.IsSend, (MSalesD.Harga*MSalesD.Qty) as Total,(MSalesD.HargaBruto*MSalesD.Qty) as TotalBruto   " & _
              '            "FROM MSalesD WHERE MSalesD.IDSales=" & IDSales & " ORDER BY MSALESD.NOID"
'              Data1.RecordSource = "SELECT MSalesD.HargaBruto,MSalesD.DiscRp,MSalesD.DiscProsen,MSalesD.DiscInternRp,MSalesD.DiscInternProsen," & _
'                        "MSalesD.DiscRp+MSalesD.DiscInternRp as JumDiscRp,MSalesD.DiscProsen+MSalesD.DiscInternProsen as JumDiscProsen," & _
'                        "MSalesD.NoID, MSalesD.IDSales, MSalesD.Transaksi, MSalesD.IDInventor, MSalesD.Qty, MSalesD.Harga, MSalesD.KodeInv," & _
'                        "MSalesD.NamaInv, MSalesD.Satuan, MSalesD.Barcode, MSalesD.IDSatuan, MSalesD.Konversi, MSalesD.HargaPokok, MSalesD.IsSend," & _
'                        "(MSalesD.Harga*MSalesD.Qty) as Total, (MSalesD.HargaBruto*MSalesD.Qty) as TotalBruto,MSalesD.IsMember   " & _
'                        "FROM MSalesD WHERE MSalesD.IDSales=" & IDSales & " ORDER BY MSALESD.NOID"
              Data1.RecordSource = "SELECT MSalesD.*,MSalesD.Qty*MSalesD.Harga as JumlahNetto  " & _
                                    " FROM MSalesD WHERE MSalesD.IDSales=" & IDSales & " ORDER BY MSalesD.NOID"

              Data1.Refresh
              If Data1.Recordset.EOF And Data1.Recordset.BOF Then
              Else
                Prin "---------------------------------------"
                Prin "*** Load From Last Transaction.... ***"
                Prin "---------------------------------------"
                Data1.Recordset.MoveFirst
                Do While Not Data1.Recordset.EOF
                  'cetakdetil Data1.Recordset!kodeinv, Data1.Recordset!namaInv, Data1.Recordset!QTY, Data1.Recordset!harga, Data1.Recordset!QTY * Data1.Recordset!harga
                    If Data1.Recordset!IsMember Then
                      cetakdetil Data1.Recordset!KodeInv, Data1.Recordset!NamaInv, Format(Data1.Recordset!Qty, "##0"), Format(Data1.Recordset!HargaBruto, "###,###,##0"), Format(Data1.Recordset!Qty * Data1.Recordset!HargaBruto, "###,###,##0")
                    Else
                      cetakdetil Data1.Recordset!KodeInv, Data1.Recordset!NamaInv & "*)", Format(Data1.Recordset!Qty, "##0"), Format(Data1.Recordset!HargaBruto, "###,###,##0"), Format(Data1.Recordset!Qty * Data1.Recordset!HargaBruto, "###,###,##0")
                    End If
                   If Data1.Recordset!DiscRp + Data1.Recordset!DiscInternRp > 0 Then
                    Dim abc As Long
                    DoEvents
                    For abc = 1 To 100000
                   
                    Next
                     ' Prin "*** Disc. " & Format(Data1.Recordset!DiscProsen, "#0.00") & "% " & Space(39 - Len("*** Disc. " & Format(Data1.Recordset!DiscProsen, "#0.00") & "% " & "-" & Format(DiscRp * Data1.Recordset!QTY, "###,##0"))) & "-" & Format(Data1.Recordset!DiscRp * Data1.Recordset!QTY, "###,##0")
                       Prin "*** Disc. " & Format(Data1.Recordset!DiscProsen + Data1.Recordset!DiscInternProsen, "#0.00") & "% " & _
                       Space(39 - Len("*** Disc. " & Format(Data1.Recordset!DiscProsen + Data1.Recordset!DiscInternProsen, "#0.00") & "% " & "-" & Format((Data1.Recordset!DiscRp + Data1.Recordset!DiscRp) * Data1.Recordset!Qty, "###,##0"))) & _
                       "-" & Format(Data1.Recordset!DiscInternRp * Data1.Recordset!Qty, "###,##0")
                    End If
                Data1.Recordset.MoveNext
                Loop
                Data1.Recordset.MoveLast
              End If
              CekDataSales
'              DiscProsenBawah = 0
'              DiscRupiahBawah = 0
              HitungSubTotal
            End If
        End If
        rspending.Close
        Ditutup = False
        
      Else
        Pending
      End If
End Sub

Private Sub cmdF1_Click()
'SEKALIGUS DI COPY UPDATE
 If Dir(DirUpdate & "\dbmaster.mdb") <> "" Then
    rs.Close
    Set rs = Nothing
    dbs.Close
    Set dbs = Nothing
    DataReedemPoin.Database.Close
    
'    FileCopy DirUpdate & "\DBMaster.mdb", DirDatabase & "\DBMaster.mdb"
    frmProses.Show 1
    If isOnline = False Then
      Set dbs = OpenDatabase(DirDatabase & "\DbMaster.mdb")
    Else
      Set dbs = OpenDatabase(DirDbServer & "\DbMaster.mdb")
    End If
    Set dbSale = OpenDatabase(DirDatabase & "\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
    Set rs = dbs.OpenRecordset("Tinv", dbOpenTable)
    rs.Index = "Kode"
    
    DataReedemPoin.DatabaseName = DirDbServer & "\DbMaster.mdb"
    DataReedemPoin.RecordSource = "SELECT * " & _
                                  " FROM MReedem WHERE IsActive=True "
    DataReedemPoin.Refresh
  End If

  Dim IsAmbil As Boolean
  Dim kodeBrg As String
  Dim idBrg As Long
  If Ditutup Then Exit Sub
      frmBarang.NamaField = "Kode"
      frmBarang.lbCari = "CARI KODE"
      frmBarang.Tampil IsAmbil, kodeBrg, idBrg
      If IsAmbil And kodeBrg <> "" Then
        rs.Index = "PrimaryKey"
        rs.Seek "=", idBrg
        If rs.NoMatch Then
        Else
            DiscProsen = 0
            DiscRp = 0
            DiscInternRp = 0
            DiscInternProsen = 0
            IsDiscBySupplier = False

          TambahJualD
          Text1.Text = ""
        End If
      End If
End Sub
Private Sub cmdF12_Click()
    If Ditutup Then Exit Sub
    If lbBintang.Visible Then 'discount total prosen
     If defDiscMemberBolehInput Then
      frmGetAngka.Tampil AmbilAngka, "Masukkan Diskon Total dalam Prosen :", True
      If AmbilAngka = -1 Then Exit Sub
'        DiscProsenBawah = AmbilAngka
        Disc = (DefPembulatan * ((AmbilAngka * SubTotal / 100) \ DefPembulatan)) '+ DiscRupiahBawah
        HitungSubTotal
        'TampilBawah
        Text1.Text = ""
     DisplayPesan "DISC NOTA #" & Space(9 - IIf(Len(Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0")) > 9, 0, Len(Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0")))) & Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0"), _
                "TOTAL     #" & Space(9 - IIf(Len(Format(Total, "###,###,##0")) > 9, 0, Len(Format(Total, "###,###,##0")))) & Format(Total, "###,###,##0")
      End If
    Else 'dbp
    frmGetAngka.Tampil AmbilAngka, "Masukkan Diskon Barang dalam Prosen :", True
    If AmbilAngka = -1 Then Exit Sub
    'If Text1.Text = "" Or IsNumeric(Text1.Text) = False Then Exit Sub
      If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        HargaBruto = 0
      Else
        HargaBruto = Data1.Recordset!HargaBruto
      End If
      
      If HargaBruto <> 0 Then
'        If TipeHargaJual <> 2 And (Bulatkan((Data1.Recordset!HargaBruto * (1 - (NullToNol(Data1.Recordset!DiscProsen1) / 100)) * (1 - (NullToNol(Data1.Recordset!DiscProsen2) / 100)) * (1 - (NullToNol(Data1.Recordset!DiscProsen3) / 100))), 0)) >= (HargaBruto * DiscProsen / 100) Then
'          If TipeHargaJual <> 2 And (TipeHargaJual <> 2 And HargaBruto - (HargaBruto * AmbilAngka / 100) < NullToNol(Data1.Recordset!HargaMin)) Then
'            jawab = False
'            frmKey.Tampil jawab, 116
'          Else
'            jawab = True
'          End If
'          If jawab Then
'            DiscProsen = AmbilAngka
'            DiscRp = HargaBruto * DiscProsen / 100
'            IsDiscBySupplier = True
'            UpdateJualD "DBP"
'          End If
'        Else
          DiscProsen = AmbilAngka
          DiscRp = HargaBruto * DiscProsen / 100
          IsDiscBySupplier = True
          UpdateJualD "DBP"
'        End If
        Text1.Text = ""
        Text1.SetFocus
      End If
    End If
End Sub
Private Sub cmdF11_Click()
    If Ditutup Then Exit Sub
   
   ' If Text1.Text = "" Or IsNumeric(Text1.Text) = False Then Exit Sub
    '  DiscRp = Text1.Text
    If lbBintang.Visible Then
          Exit Sub
          frmGetAngka.Tampil AmbilAngka, "Masukkan Diskon Total dalam Rupiah :", True
          If AmbilAngka = -1 Then Exit Sub
'            DiscRupiahBawah = AmbilAngka
'            Disc = (DefPembulatan * ((DiscProsenBawah * SubTotal / 100) \ DefPembulatan)) + AmbilAngka
            Disc = AmbilAngka
            'Disc = Disc + AmbilAngka
            HitungSubTotal
            TampilBawah
                DisplayPesan "DISC NOTA #" & Space(9 - IIf(Len(Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0")) > 9, 0, Len(Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0")))) & Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0"), _
                "TOTAL     #" & Space(9 - IIf(Len(Format(Total, "###,###,##0")) > 9, 0, Len(Format(Total, "###,###,##0")))) & Format(Total, "###,###,##0")

            Text1.Text = ""
    Else
    frmGetAngka.Tampil AmbilAngka, "Masukkan Diskon Barang dalam Rupiah :", True
    If AmbilAngka = -1 Then Exit Sub
      If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        HargaBruto = 0
      Else
        HargaBruto = Data1.Recordset!HargaBruto
      End If
      If HargaBruto <> 0 Then
        If TipeHargaJual <> 2 And (Bulatkan((Data1.Recordset!HargaBruto * (1 - (NullToNol(Data1.Recordset!DiscProsen1) / 100)) * (1 - (NullToNol(Data1.Recordset!DiscProsen2) / 100)) * (1 - (NullToNol(Data1.Recordset!DiscProsen3) / 100))), 0)) >= DiscRp Then
          If TipeHargaJual <> 2 And (HargaBruto - AmbilAngka < NullToNol(Data1.Recordset!HargaMin)) Then
            jawab = False
            frmKey.Tampil jawab, 116
          Else
            jawab = True
          End If
          If jawab Then
            DiscRp = AmbilAngka
            DiscProsen = DiscRp * 100 / HargaBruto
            IsDiscBySupplier = True
            UpdateJualD "DBR"
          End If
        Else
          DiscRp = AmbilAngka
          DiscProsen = DiscRp * 100 / HargaBruto
          IsDiscBySupplier = True
          UpdateJualD "DBR"
        End If
        Text1.Text = ""
        Text1.SetFocus
      End If
  End If
End Sub

Private Sub cmdF2_Click()
  Dim kodeBrg As String
  Dim IsAmbil As Boolean
Dim idBrg As Long
  If lbBintang.Visible = False Then
      'If MsgBox("Ingin mencari Berdasarkan Nama tekan Yes." & vbCrLf & "Cari Berdasarkan Barcode tekan No.", vbYesNo, App.Title) = vbYes Then
        frmBarang.NamaField = "Nama"
        frmBarang.KodeSearch = KodeSearch
        frmBarang.lbCari = "CARI NAMA"
        IsAmbil = False
        frmBarang.Tampil IsAmbil, kodeBrg, idBrg
        If IsAmbil And kodeBrg <> "" Then
          rs.Index = "PrimaryKey"
          rs.Seek "=", idBrg
          If rs.NoMatch Then
          Else
              DiscProsen = 0
              DiscRp = 0
              DiscInternRp = 0
              DiscInternProsen = 0
              IsDiscBySupplier = False
              IsBarangMember = False
            TambahJualD
            Text1.Text = ""
          End If
        End If
  'Else
  '    DiscINTERNNOTA = Total Mod 100
  '    TampilBawah
  End If
End Sub

Private Sub cmdF3_Click()
Dim DefHargaJual As Double
Dim posisi
Dim QtyKurang As Double
  If Ditutup Then Exit Sub
  If Data1.Recordset.EOF Or Data1.Recordset.BOF Then Exit Sub
  If Not lbBintang.Visible And (Data1.Recordset!Transaksi = "PLU" Or Data1.Recordset!Transaksi = "PE") Then 'discount total prosen
    frmGetAngka.Tampil AmbilAngka, "Masukkan Qty Baru :", False
    If AmbilAngka < 0 Or AmbilAngka > 1000000 Then Exit Sub
    QtyKurang = Data1.Recordset!Qty - AmbilAngka
    'If isRemcomendedOnline And AmbilAngka * NullToNol(Data1.Recordset!Konversi) >= SisaStockGudang Then frmPesan.lbPesan = "QTY MELEBIHI SISA STOK !!!!": frmPesan.Show 1: Exit Sub
    If Not Data1.Recordset.EOF And Not Data1.Recordset.BOF Then
          posisi = Data1.Recordset.Bookmark
          'MENCARI HARGA DARI BARANG BERDASAR RS yg sesuai
          rs.Index = "Barcode"
          rs.Seek "=", Data1.Recordset!Barcode
          If rs.NoMatch Then
            DefHargaJual = Data1.Recordset!HargaBruto
          Else
If (NullToBool(rs!IsFamilyGroup) And IsNull(rs!TanggalDariFamily)) Or (NullToBool(rs!IsFamilyGroup) And Date >= NullToDate(rs!TanggalDariFamily) And Date <= NullToDate(rs!TanggalSampaiFamily)) Then
  If JumlahQtyNotaIniSelainBarcode(rs!kode, "xxx") - QtyKurang >= NullToNol(rs!Qtyfamily) Then
      DefHargaJual = NullToNol(rs!HargaFamily)
  Else
    DefHargaJual = NullToNol(rs!HargaJual)
  End If
    If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
     Data1.Recordset.MoveFirst
     Do While Not Data1.Recordset.EOF
         
        If Data1.Recordset!KodeInv = rs!kode And DefHargaJual <> 0 Then
           Data1.Recordset.Edit
           Data1.Recordset!Disc1 = 0
           Data1.Recordset!Disc2 = 0
           Data1.Recordset!Disc3 = 0
           Data1.Recordset!HargaBruto = DefHargaJual
           Data1.Recordset!harga = DefHargaJual
           Data1.Recordset!HargaD = DefHargaJual
           Data1.Recordset!Jumlah = Bulatkan(Data1.Recordset!Qty * DefHargaJual, 0)
           Data1.Recordset.Update
           Data1.Recordset.Bookmark = Data1.Recordset.LastModified
         End If
           Data1.Recordset.MoveNext
      Loop
    End If
Else
          
           If rs!IsKelipatan Then
               If AmbilAngka = 0 Then
                  DefHargaJual = NullToNol(rs!HargaKelipatan)
              Else
                DefHargaJual = Bulatkan(((rs!HargaKelipatan * (AmbilAngka \ rs!QtyKelipatan) * rs!QtyKelipatan) + rs!HargaA * (AmbilAngka - (AmbilAngka \ rs!QtyKelipatan) * rs!QtyKelipatan)) / AmbilAngka, 0)
              End If
            Else 'If NullToBool(rs!IsGrosir) Then
                If AmbilAngka >= NullToNol(rs!Qty3) And NullToNol(rs!Qty3) <> 0 Then
                    DefHargaJual = NullToNol(rs!Harga3)
                ElseIf AmbilAngka >= NullToNol(rs!Qty2) And NullToNol(rs!Qty2) <> 0 Then
                    DefHargaJual = NullToNol(rs!Harga2)
                ElseIf AmbilAngka >= NullToNol(rs!Qty1) And NullToNol(rs!Qty1) <> 0 Then
                    DefHargaJual = NullToNol(rs!Harga1)
                Else
                    DefHargaJual = NullToNol(rs!HargaJual)
                End If
           ' Else
            '    DefHargaJual = NullToNol(rs!HargaA)
          End If

        End If
End If
      If NullToNol(rs!DiscMemberRp2) > 0 And IDMember >= 1 Then 'Kalau Ada Member
'        If Not IsNull(rs!TglDariDiskon2) And Not IsNull(rs!TglSampaiDiskon2) And Date >= NullToDate(rs!TglDariDiskon2) And Date <= NullToDate(rs!TglSampaiDiskon2) And IsQtyPDPAda(NullToNol(rs!IdInventor), AmbilAngka, NullToDate(rs!TglDariDiskon2), NullToDate(rs!TglSampaiDiskon2), CDbl(SubTotal)) Then
'          DiscInternProsen = NullToNol(rs!DiscMemberProsen2)
'          If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
'              HargaBruto = 0
'          Else
'              HargaBruto = Data1.Recordset!HargaBruto
'          End If
'          If HargaBruto <> 0 Then
'            DiscInternRp = NullToNol(rs!DiscMemberRp2) 'NilaiDiskon ' ((HargaBruto * DiscInternProsen / 100) \ 1) * 1  'tanpa koma
'            DiscInternProsen = NullToNol(rs!DiscMemberProsen2)
'            UpdateJualD
'            Text1.Text = ""
'          End If
'        Else 'Kemungkinan ke Promo Diskon
'          If NullToNol(rs!DiscRupiah) > 0 And Not IsNull(rs!DiscExpired) And Not IsNull(rs!DiscMulai) Then
'            If Date >= rs!DiscMulai And Date <= rs!DiscExpired Then
'                DiscInternProsen = NullToNol(rs!DiscProsen)
'                If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
'                    HargaBruto = 0
'                Else
'                    HargaBruto = Data1.Recordset!HargaBruto
'                End If
'                If HargaBruto <> 0 Then
'                  DiscInternRp = NullToNol(rs!DiscRupiah) 'NilaiDiskon ' ((HargaBruto * DiscInternProsen / 100) \ 1) * 1  'tanpa koma
'                  DiscInternProsen = NullToNol(rs!DiscProsen)
'                  UpdateJualD
'                  Text1.Text = ""
'                End If
'            End If
'          End If
'        End If
      ElseIf NullToNol(rs!DiscRupiah) > 0 Then
        If Not IsNull(rs!DiscExpired) And Not IsNull(rs!DiscMulai) Then
            If Date >= rs!DiscMulai And Date <= rs!DiscExpired Then
                DiscInternProsen = NullToNol(rs!DiscProsen)
                If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
                    HargaBruto = 0
                Else
                    HargaBruto = Data1.Recordset!HargaBruto
                End If
                If HargaBruto <> 0 Then
                  If IDMember >= 1 Then 'Kalau Member Ada Disc Sendiri
                    DiscInternRp = NullToNol(rs!DiscMemberRp2) 'NilaiDiskon ' ((HargaBruto * DiscInternProsen / 100) \ 1) * 1  'tanpa koma
                    DiscInternProsen = NullToNol(rs!DiscMemberProsen2)
                  Else
                    DiscInternRp = NullToNol(rs!DiscRupiah) 'NilaiDiskon ' ((HargaBruto * DiscInternProsen / 100) \ 1) * 1  'tanpa koma
                    DiscInternProsen = NullToNol(rs!DiscProsen)
                  End If
'                  UpdateJualD
                  Text1.Text = ""
                End If
            End If
        End If
      ElseIf NullToNol(rs!DiscRupiah) = 0 Then
        If IDMember >= 1 Then 'Kalau Member Ada Disc Sendiri
          DiscInternRp = NullToNol(rs!DiscMemberRp2) 'NilaiDiskon ' ((HargaBruto * DiscInternProsen / 100) \ 1) * 1  'tanpa koma
          DiscInternProsen = NullToNol(rs!DiscMemberProsen2)
        Else
          DiscInternRp = NullToNol(rs!DiscRupiah) 'NilaiDiskon ' ((HargaBruto * DiscInternProsen / 100) \ 1) * 1  'tanpa koma
          DiscInternProsen = NullToNol(rs!DiscProsen)
        End If
    End If

      Data1.Recordset.Bookmark = posisi
      Data1.Recordset.Edit
      Data1.Recordset!Qty = AmbilAngka
      If DefHargaJual <> 0 Then
        Data1.Recordset!harga = DefHargaJual - (DiscInternRp + DiscRp)
        Data1.Recordset!HargaBruto = DefHargaJual
'      Data1.Recordset!HargaD = Bulatkan((DefHargaJual * (1 - (NullToNol(Data1.Recordset!DiscProsen1) / 100)) * (1 - (NullToNol(Data1.Recordset!DiscProsen2) / 100)) * (1 - (NullToNol(Data1.Recordset!DiscProsen3) / 100))), 0)
        Data1.Recordset!Jumlah = Bulatkan(AmbilAngka * (DefHargaJual - (DiscInternRp + DiscRp)), 0) ' (Bulatkan((DefHargaJual * (1 - (NullToNol(Data1.Recordset!DiscProsen1) / 100)) * (1 - (NullToNol(Data1.Recordset!DiscProsen2) / 100)) * (1 - (NullToNol(Data1.Recordset!DiscProsen3) / 100))), 0) - Data1.Recordset!Disc1 - Data1.Recordset!Disc2 - Data1.Recordset!Disc3)
      Else
        Data1.Recordset!Jumlah = Bulatkan(AmbilAngka * Data1.Recordset!harga, 0)
      End If
      Data1.Recordset.Update
      HitungSubTotal
'      If IDMember >= 1 Then
'        UpdatekanDiskonPerBarang
'      End If
    End If
  End If
End Sub

Private Sub UpdatekanDiskonPerBarang()
Dim DB As Database
Dim rs As Recordset
Dim rsSale As Recordset
Dim IDBArang As Long, DiscMbrProsen As Double, TglDari As Date, TglSampai As Date, Jumlah As Double
Dim BarangPDP As Boolean
On Error GoTo Trace:
Set DB = OpenDatabase(App.path & "\Database\DBmaster.mdb")
If IDMember >= 1 And Data1.Recordset.RecordCount >= 1 Then
  Data1.Recordset.MoveFirst
  Do While Not Data1.Recordset.EOF
    Data1.Recordset.Edit
    IDBArang = NullToNol(Data1.Recordset!IdInventor)
    Set rs = DB.OpenRecordset("SELECT * FROM Tinv WHERE IDInventor=" & IDBArang)
    If Not (rs.BOF Or rs.EOF) Then
      Set rsSale = dbSale.OpenRecordset("SELECT SUM(MSalesD.Qty*MSalesD.HargaBruto) AS Subtotal FROM MSalesD WHERE MSalesD.IDSales=" & IDSales)
      If Not (rsSale.EOF Or rsSale.BOF) Then
        Jumlah = NullToNol(rsSale!SubTotal)
      Else
        Jumlah = 0
      End If
      DiscMbrProsen = NullToNol(rs!DiscMemberProsen2)
      TglDari = NullToDate(rs!TglDariDiskon2)
      TglSampai = NullToDate(rs!TglSampaiDiskon2)
      BarangPDP = IIf(IsNull(Data1.Recordset!IsPDP), False, Data1.Recordset!IsPDP)
      'If DiscMbrProsen > 0 And Jumlah >= defMinialBelanjadapatDiskon And Not IsNull(rs!TglDariDiskon2) And Not IsNull(rs!TglSampaiDiskon2) And Date >= TglDari And Date <= TglSampai And IsQtyPDPAda(IDBarang, NullToNol(Data1.Recordset!Qty), TglDari, TglSampai, CDbl(Jumlah)) Then
      If DiscMbrProsen > 0 And Jumlah >= defMinialBelanjadapatDiskon2 And defMinialBelanjadapatDiskon2 > 0 And Not BarangPDP And Not IsNull(rs!TglDariDiskon2) And Not IsNull(rs!TglSampaiDiskon2) And Date >= TglDari And Date <= TglSampai Then
        If NullToNol(Data1.Recordset!HargaBruto) <> 0 Then
          Data1.Recordset!DiscInternProsen = DiscMbrProsen
          Data1.Recordset!DiscInternRp = NullToNol(rs!DiscMemberRp2)
          Data1.Recordset!IsDisc2 = True
        Else
          Data1.Recordset!DiscInternRp = 0
          Data1.Recordset!DiscInternProsen = 0
          Data1.Recordset!IsDisc2 = False
        End If
      Else 'Pindah ke Promo
        DiscMbrProsen = NullToNol(rs!DiscProsen)
        TglDari = NullToDate(rs!DiscMulai)
        TglSampai = NullToDate(rs!DiscExpired)
        If DiscMbrProsen > 0 And Not IsNull(rs!DiscMulai) And Date >= TglDari And Date <= TglSampai Then
          If NullToNol(Data1.Recordset!HargaBruto) <> 0 Then
            Data1.Recordset!DiscInternProsen = DiscMbrProsen
            Data1.Recordset!DiscInternRp = NullToNol(rs!DiscRupiah)
            Data1.Recordset!IsDisc2 = False
          Else
            Data1.Recordset!DiscInternRp = 0
            Data1.Recordset!DiscInternProsen = 0
            Data1.Recordset!IsDisc2 = False
          End If
        Else
          Data1.Recordset!DiscInternRp = 0
          Data1.Recordset!DiscInternProsen = 0
          Data1.Recordset!IsDisc2 = False
        End If
      End If
      Data1.Recordset!HargaPokok = NullToNol(rs!HargaPokok) - NullToNol(Data1.Recordset!DiscInternRp)
    Else
      Data1.Recordset!DiscInternRp = 0
      Data1.Recordset!DiscInternProsen = 0
      Data1.Recordset!IsDisc2 = False
      Data1.Recordset!HargaPokok = NullToNol(rs!HargaPokok) - NullToNol(Data1.Recordset!DiscInternRp)
    End If
    Data1.Recordset!harga = NullToNol(Data1.Recordset!HargaBruto) - NullToNol(Data1.Recordset!DiscInternRp)
    Data1.Recordset!Jumlah = Bulatkan(NullToNol(Data1.Recordset!harga) * NullToNol(Data1.Recordset!Qty), 0)
    Data1.Recordset.Update
    Data1.Recordset.MoveNext
  Loop
Else 'Untuk non Member di kosongi dulu
  If Data1.Recordset.RecordCount >= 1 Then
    Data1.Recordset.MoveFirst
    Do While Not Data1.Recordset.EOF
      Data1.Recordset.Edit
      IDBArang = NullToNol(Data1.Recordset!IdInventor)
      Set rs = DB.OpenRecordset("SELECT * FROM Tinv WHERE IDInventor=" & IDBArang)
      If Not (rs.BOF Or rs.EOF) Then
        Data1.Recordset!HargaPokok = NullToNol(rs!HargaPokok) - NullToNol(Data1.Recordset!DiscInternRp)
      End If
      Data1.Recordset.Update
      Data1.Recordset.MoveNext
    Loop
  End If
End If
HitungSubTotal
Err.Clear
'TampilBawah
Trace:
  If Err.Number <> 0 Then
    'Resume Next
    MsgBox "Error : " & Err.Number & ", " & Err.Description, vbCritical, "VPOS"
    Err.Clear
  End If
  Set rs = Nothing
  Set DB = Nothing
  Set rsSale = Nothing
End Sub

Private Function IsQtyPDPAda(ByVal IDBArang As Long, QtyBarang As Double, TglDari As Date, TglSampai As Date, ByVal SubtotalSekarang As Double) As Boolean
Dim cn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim Hasil As Boolean, SQL As String, QtyPDP As Double, QtyDiPenjualan As Double
On Error GoTo Trace:
Hasil = False
If isRemcomendedOnline Then
  cn.ConnectionString = Cnstr
  cn.Open
  SQL = "SELECT IsNull(MBarang.QtyPDP,0) AS Qty FROM MBarang WHERE NoID=" & IDBArang
  Set rst = cn.Execute(SQL)
  If Not (rst.BOF Or rst.EOF) Then
    QtyPDP = NullToNol(rst!Qty)
  Else
    QtyPDP = 0
  End If
  If rst.State = adStateOpen Then
    rst.Close
  End If
  Set rst = Nothing
  SQL = "SELECT SUM(MJualD.Qty*MJualD.Konversi) AS Qty FROM MBarang LEFT JOIN (MJual INNER JOIN MJualD ON MJual.NoID=MJualD.IDJual) ON MJualD.IDBarang=MBarang.NoID WHERE MJual.IDCustomer>=1 AND MBarang.NoID=" & IDBArang & " AND Mjual.Tanggal>='" & Format(TglDari, "yyyy-MM-dd") & "' And MJual.Tanggal<'" & Format(DateAdd("d", 1, TglSampai), "yyyy-MM-dd") & "'"
  Set rst = cn.Execute(SQL)
  If Not (rst.BOF Or rst.EOF) Then
    QtyDiPenjualan = NullToNol(rst!Qty)
  Else
    QtyDiPenjualan = 0
  End If
  If (Qty - QtyBarang - QtyDiPenjualan >= 0) And SubtotalSekarang >= defMinialBelanjadapatDiskon Then
    Hasil = True
  Else
    Hasil = False
  End If
End If
Trace:
  If Err.Number <> 0 Then
    MsgBox "Error : " & Err.Number & " - " & Err.Description, vbCritical, "VPOS"
    Err.Clear
  End If
  If cn.State = adStateOpen Then
    cn.Close
  End If
  Set cn = Nothing
  If rst.State = adStateOpen Then
    rst.Close
  End If
  Set rst = Nothing
  IsQtyPDPAda = Hasil
End Function

Private Sub cmdF6_Click()
  'If KodeMember <> "" Then frmPesan.lbPesan.Caption = "Customer sudah ada !!!": frmPesan.Show 1: Exit Sub
  Dim rsSales As Recordset
  Dim rsSalesD As Recordset
  Set rsSales = dbSale.OpenRecordset("Select * FROM MSales Where NoID=" & IDSales)
  Set rsSalesD = dbSale.OpenRecordset("Select * FROM MSalesD Where NoID=" & IDSales)
  If rsSales.EOF And rsSales.BOF Then
    frmMember.IsAllow = True
    frmMember.Show 1
    If KodeMember <> "" Then
      rsSales.Edit
      rsSales!KodeMember = NamaMember
      rsSales!IDMember = IDMember
      rsSales.Update
      UpdatekanDiskonPerBarang
    End If
    rsSales.Close
    Set rsSales = Nothing
  Else
    If rsSalesD.RecordCount <= 1 Then
      frmMember.Tampil True
      'If KodeMember <> "" Then
        rsSales.Edit
        rsSales!KodeMember = IIf(NamaMember = "", " ", NamaMember)
        rsSales!IDMember = IDMember
        rsSales.Update
        UpdatekanDiskonPerBarang
     ' End If
    Else
      frmMember.IsAllow = False
      frmMember.Show 1
      'If KodeMember <> "" Then
        rsSales.Edit
        rsSales!KodeMember = NamaMember
        rsSales!IDMember = IDMember
        rsSales.Update
        UpdatekanDiskonPerBarang
      'End If
'      rsSales.Close
'      Set rsSales = Nothing
    End If
    rsSales.Close
    Set rsSales = Nothing
  End If
  rsSalesD.Close
  Set rsSalesD = Nothing
  lbNama.Caption = NamaMember
  GetPoinMember
Text1.SetFocus
End Sub

Private Sub cmdHOME_Click()
Text1.Text = ""
  If Ditutup Then
    IsNotaDariPending = False
    CountReprint = 0
    Disc = 0
    JumDiscInternRp = 0
    DiscINTERNNOTA = 0
    PotonganPembulatan = 0
    TotalDiscBrg = 0
    BuatBaru
    bolehbergerak = True
    Ditutup = False
    BukaCommBarcode
    If Not IsHematKertas Then
            Prin Chr(13) & Chr(10)
            PrinBigChar Space(15 - Min(15, Len(NamaToko))) & UCase(NamaToko)
            DoEvents
            Prin Judulstruk & Chr(13) & Chr(10) & "---------------------------------------"
            DoEvents
    End If
  End If
End Sub
Private Sub cmdPGUP_Click()
Dim TasA As Double
Dim TasB As Double
Dim TasC As Double
Dim TasD As Double
Dim jwb As Boolean
      UpdateFooter
      If Not lbBintang.Visible Then
'        KeyCode = 0
        Exit Sub
      End If
      
      If Len(Text1.Text) > 9 Then
'        KeyCode = 0
        Exit Sub
      End If
      
      If NullToNol(lblReedemPoin.Caption) >= 1 And NullToNol(lblTotalPoin.Caption) - NullToNol(lblReedemPoin.Caption) < 0 Then
        MsgBox "Total Poin member tidak mencukupi untuk melakukan reedem poin.", vbCritical + vbOKOnly
        Exit Sub
      End If
      If Text1.Text = "" Or Text1.Text = "0" Then
'        If IDMember > 0 Then
'            If MsgBox("Apakah Yakin Mau lanjutkan Belanja hutang?", vbYesNo + vbQuestion) = vbYes Then
'               If LimitHutang > 0 Then
'                TampilBawah
'                If Kembali >= 0 Then Ditutup = True
'                'Text1.Text = ""
'                'Text1.Text = " " & Format( Kembali, "###,###,##0")
'                Text1.Text = "CHG#" & Space(10 - Min(Len(Format(Kembali, "###,###,##0")), 10)) & Format(Kembali, "###,###,##0")
'                TutupCommBarcode
''                DiscProsenBawah = 0
''                DiscRupiahBawah = 0
'                Else
'                     If SaldoHutang <= JumlahBolehHutang(LimitHutang, IDMember) Then
'                        TampilBawah
'                        If Kembali >= 0 Then Ditutup = True
'                        'Text1.Text = ""
'                        'Text1.Text = " " & Format( Kembali, "###,###,##0")
'                        Text1.Text = "CHG#" & Space(10 - Min(Len(Format(Kembali, "###,###,##0")), 10)) & Format(Kembali, "###,###,##0")
'                        TutupCommBarcode
''                        DiscProsenBawah = 0
''                        DiscRupiahBawah = 0
'                    Else
'                        MsgBox "Limit Hutang Customer ini tidak mencukupi!", vbCritical, "Tidak dapat menyimpan"
'                    End If
'                End If
'            End If
'        Else
            Dibayar = Total
            Bank = 0
            TampilBawah
            If Kembali >= 0 Then Ditutup = True
            'Text1.Text = ""
            'Text1.Text = " " & Format( Kembali, "###,###,##0")
            Text1.Text = "CHG#" & Space(10 - Min(Len(Format(Kembali, "###,###,##0")), 10)) & Format(Kembali, "###,###,##0")
            TutupCommBarcode
'            DiscProsenBawah = 0
'            DiscRupiahBawah = 0
'        End If
      ElseIf IsNumeric(Text1.Text) Then
        Dibayar = CCur(Text1.Text) + Voucher + ReedemNilai 'Dibayar + CCur(Text1.Text)
'        If IDMember > 0 Then
'          If Dibayar < SaldoHutang Then
'            If MsgBox("Apakah Yakin Mau lanjutkan Belanja hutang?", vbYesNo + vbQuestion) = vbYes Then
'              If JumlahBolehHutang(LimitHutang, IDMember) + Dibayar <= SaldoHutang Then
'                  If SaldoHutang <= JumlahBolehHutang(LimitHutang, IDMember) + Dibayar Then
'                      TampilBawah
'                      If Kembali >= 0 Then Ditutup = True
'                      'Text1.Text = ""
'                      'Text1.Text = " " & Format( Kembali, "###,###,##0")
'                      Text1.Text = "CHG#" & Space(10 - Min(Len(Format(Kembali, "###,###,##0")), 10)) & Format(Kembali, "###,###,##0")
'                      TutupCommBarcode
''                      DiscProsenBawah = 0
''                      DiscRupiahBawah = 0
'                  Else
'                      MsgBox "Limit Hutang Customer ini tidak mencukupi!", vbCritical, "Tidak dapat menyimpan"
'                  End If
'              Else
'                Bank = 0
'                TampilBawah
'                If Kembali >= 0 Then Ditutup = True
'                'Text1.Text = ""
'                'Text1.Text = " " & Format( Kembali, "###,###,##0")
'                Text1.Text = "CHG#" & Space(10 - Min(Len(Format(Kembali, "###,###,##0")), 10)) & Format(Kembali, "###,###,##0")
'                TutupCommBarcode
''                DiscProsenBawah = 0
''                DiscRupiahBawah = 0
'              End If
'            End If
'          Else
'            Bank = 0
'            TampilBawah
'            If Kembali >= 0 Then Ditutup = True
'            'Text1.Text = ""
'            'Text1.Text = " " & Format( Kembali, "###,###,##0")
'            Text1.Text = "CHG#" & Space(10 - Min(Len(Format(Kembali, "###,###,##0")), 10)) & Format(Kembali, "###,###,##0")
'            TutupCommBarcode
''            DiscProsenBawah = 0
''            DiscRupiahBawah = 0
'          End If
'        Else
          Bank = 0
          TampilBawah
          If Kembali >= 0 Then Ditutup = True
          'Text1.Text = ""
          'Text1.Text = " " & Format(Kembali, "###,###,##0")
          Text1.Text = "CHG#" & Space(10 - Min(Len(Format(Kembali, "###,###,##0")), 10)) & Format(Kembali, "###,###,##0")
          TutupCommBarcode
'          DiscProsenBawah = 0
'          DiscRupiahBawah = 0
'        End If
      End If
'        If Ditutup Then
'            openDrawer
'            DoEvents
'        End If
      
      DisplayPesan "PAY# " & Space(15 - Min(Len(Format(Dibayar - Voucher, "###,###,##0")), 15)) & Format(Dibayar - Voucher, "###,###,##0"), "CHG# " & Space(15 - Min(Len(Format(Kembali, "###,###,##0")), 15)) & Format(Kembali, "###,###,##0")
      HitungItem
      If Kembali >= 0 Then
      Dim pPesan As String
      pPesan = "---------------------------------------" & Chr(13) & Chr(10) & _
                "Subtotal" & Space(31 - Len(Format(SubTotal, "###,###,##0"))) & Format(SubTotal, "###,###,##0") & Chr(13) & Chr(10) & _
                IIf(TotalDiscBrg + JumDiscInternRp = 0, "", "Discount Barang" & Space(24 - Len(Format(TotalDiscBrg + JumDiscInternRp, "###,###,##0"))) & Format(TotalDiscBrg + JumDiscInternRp, "###,###,##0") & Chr(13) & Chr(10)) & _
                IIf(Disc + DiscINTERNNOTA + PotonganPembulatan = 0, "", "Discount Nota" & Space(26 - Len(Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0"))) & Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0") & Chr(13) & Chr(10)) & _
                "Voucher " & Space(31 - Len(Format(Voucher, "###,###,##0"))) & Format(Voucher, "###,###,##0") & Chr(13) & Chr(10) & _
                "Dibayar " & Space(31 - Len(Format(Dibayar - Voucher, "###,###,##0"))) & Format(Dibayar - Voucher, "###,###,##0") & Chr(13) & Chr(10) & _
                IIf(SaldoHutang <= 0, "", "Hutang  " & Space(31 - Len(Format(SaldoHutang, "###,###,##0"))) & Format(SaldoHutang, "###,###,##0") & Chr(13) & Chr(10)) & _
                "Kembali " & Space(31 - Len(Format(Kembali, "###,###,##0"))) & Format(Kembali, "###,###,##0") & Chr(13) & Chr(10) & _
                IIf(ISSembunyikanFooterbarangKSB, "", "Belanja Barang POIN" & Space(20 - Len(Format(BelanjaPoin, "###,###,##0"))) & Format(BelanjaPoin, "###,###,##0") & Chr(13) & Chr(10)) & _
                "#" & Format(CLng(NoNota), "0#####") & "," & "Kasir" & NamaMesin & ":" & Mid(NamaKasir, 1, 6) & Space(6 - IIf(Len(NamaKasir) <= 6, Len(NamaKasir), 6)) & Format(Date, " dd/MM/yy") & Format(Time, "hh:mm:ss") & Chr(13) & Chr(10) & _
                "Macam Barang =" & Format(MacamItem, "##0") & " , Total Item =" & Format(Jumlahitem, "##0") & Chr(13) & Chr(10) & _
                IIf(KodeMember = "", "", "CUSTOMER: " & KodeMember & "-" & NamaMember & Chr(13) & Chr(10)) & _
                "----TERIMA KASIH ATAS KUNJUNGAN ANDA----" & Chr(13) & Chr(10) & _
                IIf(Footer1 = "", "", Footer1 & Chr(13) & Chr(10)) & _
                IIf(Footer2 = "", "", Footer2 & Chr(13) & Chr(10)) & _
                IIf(Footer3 = "", "", Footer3 & Chr(13) & Chr(10))
                Dim i As Integer
                For i = 1 To SpasiFooter
                  pPesan = pPesan & Chr(13) & Chr(10)
                Next
      Prin pPesan
              If IsHematKertas Then
                DoEvents
                PrinBigChar Space(15 - Min(15, Len(NamaToko))) & UCase(NamaToko)
                Prin Judulstruk & Chr(13) & Chr(10) & "---------------------------------------"
              End If
            papercut
        If Ditutup Then
          If NoPortDrawer <> -2 Then
            openDrawer
          End If
          CetakStruck IDSales, False
          DoEvents
          If IsPakaiKantong Then
            jwb = False
            frmPayment2.Tampil jwb, TasA, TasB, TasC, TasD
            If Not jwb Then
  '            Exit Sub
            Else
              Dim rsSales As Recordset
              Set rsSales = dbSale.OpenRecordset("Select * FROM MSales Where NoID=" & IDSales)
              If Not (rsSales.EOF And rsSales.BOF) Then
                rsSales.Edit
                rsSales!TasKresekA = TasA
                rsSales!TasKresekB = TasB
                rsSales!TasKresekC = TasC
                rsSales!TasKresekD = TasD
                rsSales.Update
              End If
              rsSales.Close
              Set rsSales = Nothing
            End If
          End If
            If isRemcomendedOnline Then
                KirimKeServerBeginTrans IDSales
                CetakStruckReedem2 IDSales, False
            End If
        End If
       lbBintang.Visible = False
       cmdReedemPoin.Visible = False
       TutupCommBarcode
       
'        Prin "---------------------------------------" & Chr(13) & Chr(10) & _
'                "Subtotal" & Space(31 - Len(Format(SubTotal, "###,###,##0"))) & Format(SubTotal, "###,###,##0") & Chr(13) & Chr(10) & _
'                "Voucher " & Space(31 - Len(Format(Voucher, "###,###,##0"))) & Format(Voucher, "###,###,##0") & Chr(13) & Chr(10) & _
'              "Dibayar " & Space(31 - Len(Format(Dibayar - Voucher, "###,###,##0"))) & Format(Dibayar - Voucher, "###,###,##0") & Chr(13) & Chr(10) & _
'              "Kembali " & Space(31 - Len(Format(Kembali, "###,###,##0"))) & Format(Kembali, "###,###,##0") & Chr(13) & Chr(10) & _
'              "#" & Format(CLNG(NONOTA), "0#####") & "," & "Kasir" & NamaMesin & ":" & Mid(NamaKasir, 1, 6) & Space(6 - IIf(Len(NamaKasir) <= 6, Len(NamaKasir), 6)) & Format(Date, " dd/MM/yy") & Format(Time, "hh:mm:ss") & Chr(13) & Chr(10) & _
'              "Macam Barang =" & Format(MacamItem, "##0") & " , Total Item =" & Format(Jumlahitem, "##0") & Chr(13) & Chr(10) & "----TERIMA KASIH ATAS KUNJUNGAN ANDA----" & Chr(13) & Chr(10) & _
'              IIf(Footer1 = "", "", Footer1 & Chr(13) & Chr(10)) & _
'              IIf(Footer2 = "", "", Footer2 & Chr(13) & Chr(10)) & _
'              IIf(Footer3 = "", "", Footer3 & Chr(13) & Chr(10)) & _
'              Chr(13) & Chr(10)
'
'        PrinBigChar Space(15 - Min(15, Len(NamaToko))) & UCase(NamaToko)
'        Prin Judulstruk & Chr(13) & Chr(10) & "---------------------------------------"
'        papercut
'       lbBintang.Visible = False
'       TutupCommBarcode
      Else
        Prin "---------------------------------------" & Chr(13) & Chr(10) & "Dibayar " & Space(31 - Len(Format(Dibayar - Voucher, "###,###,##0"))) & Format(Dibayar - Voucher, "###,###,##0") & Chr(13) & Chr(10) & "Kembali " & Space(31 - Len(Format(Kembali, "###,###,##0"))) & Format(Kembali, "###,###,##0")
      End If
'      KeyCode = 0
End Sub
Private Sub cmdPGDN_Click()
    If Data1.Recordset.BOF Then Exit Sub
    jawab = False
     If SubTotal < 0 Then
      frmKey.Tampil jawab, 166
      If jawab = False Then Exit Sub
     End If
     lbBintang.Visible = Not lbBintang.Visible
     If lbBintang.Visible Then
        TutupCommBarcode
     Else
        BukaCommBarcode
     End If
     If lbBintang.Visible Then
     'PEMBULATAN
        If IsNumeric(LbDisc.Caption) And CDbl(LbDisc.Caption) >= 1 Then
'          If DiscProsenBawah >= 0 Then
'              Disc = DefPembulatan * ((DiscProsenBawah * SubTotal / 100) \ DefPembulatan)
'          End If
        '  Disc = CDbl(LbDisc.Caption) '(DefPembulatan * ((DiscProsenBawah * SubTotal / 100) \ DefPembulatan)) + DiscRupiahBawah
          HitungSubTotal
          ' TampilBawah
           Text1.Text = ""
           DisplayPesan "DISC NOTA #" & Space(9 - IIf(Len(Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0")) > 9, 0, Len(Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0")))) & Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0"), _
                        "TOTAL     #" & Space(9 - IIf(Len(Format(Total, "###,###,##0")) > 9, 0, Len(Format(Total, "###,###,##0")))) & Format(Total, "###,###,##0")
        Else
          DisplayPesan "TOTAL #" & Space(13 - IIf(Len(Format(Total, "###,###,##0")) > 13, 0, Len(Format(Total, "###,###,##0")))) & Format(Total, "###,###,##0"), "" ', IIf(SubTotal Mod DefPembulatan = 0, Space(20), "DISC PEMBULATAN#" & Space(4 - Len(Format(SubTotal Mod DefPembulatan, "###,###,##0"))) & Format(SubTotal Mod DefPembulatan, "###,###,##0"))
          'DisplayPesan "TOTAL #" & Space(13 - Len(Format(SubTotal, "###,###,##0"))) & Format(SubTotal, "###,###,##0"), Space(20)
        End If
        'Hartani Jika DiPgDown minta kembali ke Barcode bukan. 01/Jun/2012
        If lbBintang.Visible And IsNumeric(Text1.Text) Then
          Text1.Text = Format(CCur(Text1.Text), "###,##0")
        End If
        Text1.SelStart = Len(Text1.Text)
        
        'Request Hartani 01-06-2015
        TampilkanButtonReedem
      Else
        Text1.Text = Format(Text1.Text, "##0")
        Text1.SelStart = Len(Text1.Text)
        cmdReedemPoin.Visible = False
      End If
End Sub

Private Sub cmdPilihReedem_Click()
  If (DataReedemPoin.Recordset.EOF Or DataReedemPoin.Recordset.BOF) Then
    
  ElseIf NullToNol(lblTotalPoin.Caption) - NullToNol(DataReedemPoin.Recordset!Poin) < 0 Then
      MsgBox "Total Poin member tidak mencukupi untuk melakukan reedem poin.", vbCritical + vbOKOnly
      Exit Sub
  Else
    If NullToNol(DataReedemPoin.Recordset!Nilai) + Dibayar <= Total Then
      Dibayar = Dibayar - ReedemNilai
      IDReedemPoin = NullToNol(DataReedemPoin.Recordset!NoID)
      ReedemPoin = NullToNol(DataReedemPoin.Recordset!Poin)
      ReedemNilai = NullToNol(DataReedemPoin.Recordset!Nilai)
      Dibayar = Dibayar + ReedemNilai
      TampilBawah
      Text1.Enabled = True
      MenuReedemPoin.Visible = False
      Text1.SetFocus
    Else
      MsgBox "Total reedem poin melebihi total pembelanjaan.", vbCritical + vbOKOnly
    End If
  End If
End Sub

Private Sub cmdReedemPoin_Click()
  Text1.Enabled = False
  GetPoinMember
  Text1.Enabled = True
  MenuReedemPoin.Visible = True
  TDBGrid3.SetFocus
End Sub

Private Sub Form_Activate()
On Error Resume Next
   'Load the layout.
    TDBGrid1.LayoutName = "Layoutku"
    If Dir(App.path & "\" & Me.Name & "_" & TDBGrid1.Name & ".grx") <> "" Then
      TDBGrid1.LayoutFileName = App.path & "\" & Me.Name & "_" & TDBGrid1.Name & ".grx"
      TDBGrid1.LoadLayout
    End If
If isTampilSaldoStock = True Then
  TDBGrid1.Columns(9).Visible = True
  TDBGrid1.Columns(9).Width = 600
Else
  TDBGrid1.Columns(9).Visible = False
  TDBGrid1.Columns(9).Width = 0
End If
  BarcodeIn = ""
  AllowBarcode = True
  BukaCommBarcode
  Text1.SetFocus
  lbKassa.Caption = "KASSA : " & NamaMesin
End Sub

Private Sub Form_DeActivate()
  AllowBarcode = False
  TutupCommBarcode
End Sub

Private Sub Form_Load()
On Error Resume Next
    If Not IsHematKertas Then
            Prin Chr(13) & Chr(10)
            PrinBigChar Space(15 - Min(15, Len(NamaToko))) & UCase(NamaToko)
            DoEvents
            Prin Judulstruk & Chr(13) & Chr(10) & "---------------------------------------"
            DoEvents
    End If
  skala1 = 50
  Timer2.Enabled = True
  bolehbergerak = True
  isHasilKonversi = False
  Ditutup = False
  Qty = 1
  lbQty = Format(Qty, "##0") & " X"
  If isOnline = False Then
    Set dbs = OpenDatabase(DirDatabase & "\DbMaster.mdb")
  Else
    Set dbs = OpenDatabase(DirDbServer & "\DbMaster.mdb")
  End If
  
  Set dbSale = OpenDatabase(DirDatabase & "\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
  Set rs = dbs.OpenRecordset("Tinv", dbOpenTable)
  rs.Index = "Kode"
  IDSales = GetLastTime 'GetNewID("MSALES") - 1 'Ambil Transaksi terakhir
  CekLastEdit
  DataReedemPoin.DatabaseName = DirDatabase & "\DBMaster.mdb"
  DataReedemPoin.RecordSource = "SELECT * " & _
                                  " FROM MReedem WHERE IsActive=True "
  DataReedemPoin.Refresh
  
  Data1.DatabaseName = DirDatabase & "\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
''  Data1.RecordSource = "SELECT MSalesD.NoID, MSalesD.Transaksi,MSalesD.IDSales, MSalesD.IDInventor, MSalesD.Qty, MSalesD.Harga, MSalesD.KodeInv, MSalesD.NamaInv, MSalesD.Satuan, MSalesD.Barcode, MSalesD.IDSatuan, MSalesD.Konversi, MSalesD.HargaPokok, MSalesD.IsSend, (MSalesD.Harga*MSalesD.Qty) as Total,MSalesD.DiscRupiah,MSalesD.DiscRp " & _
''                        "FROM MSalesD WHERE MSalesD.IDSales=" & IDSales & " ORDER BY MSALESD.NOID "
'   Data1.RecordSource = "SELECT MSalesD.HargaBruto,MSalesD.DiscRp,MSalesD.DiscProsen,MSalesD.DiscInternRp,MSalesD.DiscInternProsen," & _
'                        "MSalesD.DiscRp+MSalesD.DiscInternRp as JumDiscRp,MSalesD.DiscProsen+MSalesD.DiscInternProsen as JumDiscProsen," & _
'                        "MSalesD.NoID, MSalesD.IDSales, MSalesD.Transaksi, MSalesD.IDInventor, MSalesD.Qty, MSalesD.Harga, MSalesD.KodeInv, MSalesD.NamaInv, MSalesD.Satuan, MSalesD.Barcode, MSalesD.IDSatuan, MSalesD.Konversi, MSalesD.HargaPokok, MSalesD.IsSend, (MSalesD.Harga*MSalesD.Qty) as Total, (MSalesD.HargaBruto*MSalesD.Qty) as TotalBruto,MSalesD.IsMember   " & _
'                        "FROM MSalesD WHERE MSalesD.IDSales=" & IDSales & " ORDER BY MSALESD.NOID"
   Data1.RecordSource = "SELECT MSalesD.*,MSalesD.Qty*MSalesD.Harga as JumlahNetto " & _
                        " FROM MSalesD WHERE MSalesD.IDSales=" & IDSales & " ORDER BY MSalesD.NOID"
  Data1.Refresh
  HitungSubTotal
  'DiscProsenBawah = 0
  lbTanggal.Caption = Format(Date, "dd MMM yyyy")
  lbKasir.Caption = "Kasir  : " & NamaKasir
  lbStatus = "Status  : " & GetStatusNetwork
  dxLabel1.Caption = NamaToko
'  dxLabel1.Caption = "SELAMAT DATANG"
  
  TutupCommBarcode
If NoPortBarcode > 0 Then
  MSComm1.CommPort = NoPortBarcode
  MSComm1.PortOpen = True
End If
  BukaPortPrinter
  BukaCommBarcode
  

    
    If isOnline Then
      Me.Caption = IIf(x.HasilX = Trial, "*TRIAL ", "") & UCase("Penjualan Kasir [" & ExecuteSkalarSQL("SELECT Nama FROM MGudang WHERE NoID=" & IDGudangDef) & "]")
    Else
      Me.Caption = IIf(x.HasilX = Trial, "*TRIAL ", "") & UCase("Penjualan Kasir [Lokal]")
    End If
    interval = 1
'  Text1.SetFocus
LoadImage

End Sub

Private Sub LoadImage()
  On Error GoTo Trace:
  Dim path As String
  
  path = App.path & "\image\logo\vpos.jpg"
  If Dir(path) <> "" Then
    Image1.Stretch = True
    Image1.Picture = LoadPicture(path)
  Else
    Image1.Picture = Nothing
  End If
Trace:
  If Err.Number <> 0 Then
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical
    Err.Clear
  End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
  dbs.Close
  dbSale.Close
  MSComm1.PortOpen = False
  
  'Save Layouts
  'TDBGrid1.LayoutName = "Layoutku"
  TDBGrid1.LayoutFileName = App.path & "\" & Me.Name & "_" & TDBGrid1.Name & ".grx"
  TDBGrid1.Layouts.Add "Layoutku"
  Form2.Show
  Form2.SetFocus
End Sub

Private Sub GetPoinMember()
  If IDMember >= 1 Then
    lblTotalPoin.Caption = Format(GetNilaiPoinMember(IDMember), "#,###0")
  Else
    lblTotalPoin.Caption = 0
    
    IDReedemPoin = 0
    ReedemNilai = 0
    ReedemPoin = 0
  End If
  TampilBawah
End Sub

Private Sub TampilkanButtonReedem()
  If IsNumeric(lblTotalPoin.Caption) And CDbl(lblTotalPoin.Caption) >= 1 Then
    cmdReedemPoin.Visible = True
  Else
    cmdReedemPoin.Visible = False
  End If
End Sub

Private Sub MSComm1_OnComm()
'taruh disini kasus, yaitu buffer memory di barcode belum muncul maka ketika
'dibuka, yang sebelumnya yang muncul duluan
''rev tgl 14 juli 2008
'        If (Not AllowBarcode) Or Ditutup Or lbBintang.Visible Then Exit Sub
        If isSupervisor Then
            'frmPesan.lbPesan = "Hanya Kasir Yang Boleh Jual!!!!"
'            frmPesan.Show 1
            Exit Sub
        End If
        Select Case MSComm1.CommEvent
        ' Event messages.
        Case comEvReceive
            Dim buffer As Variant
            Dim pos As Integer
            buffer = MSComm1.Input
            'ditambah per 14 juli 2008
            'artinya buffer di mscomm1 diambil kememory PC (biar kosong)
            'tapi langsung keluar
            If (Not AllowBarcode) Or Ditutup Or lbBintang.Visible Then Exit Sub
            BarcodeIn = BarcodeIn & StrConv(buffer, vbUnicode)
            pos = InStr(1, BarcodeIn, Chr(13))
            Text1.Text = BarcodeIn 'TAMBAHAN
            If pos Then
              Text1.Text = Left(BarcodeIn, pos - 1)
              BarcodeIn = ""
'              If lbBintang.Visible Or Ditutup Then
'                Text1.Text = ""
'                Exit Sub
'              End If
                If Text1.Text = "" Then Exit Sub
                rs.Index = "Barcode"
                rs.Seek "=", Text1.Text
                If rs.NoMatch Then
                  rs.Index = "Kode"
                  rs.Seek "=", Text1.Text
                  If rs.NoMatch Then
                    rs.Index = "Nama"
                    rs.Seek "=", Text1.Text
                    If rs.NoMatch Then
                    Else
                        DiscProsen = 0
                        DiscRp = 0
                        DiscInternRp = 0
                        DiscInternProsen = 0
                        IsDiscBySupplier = False
                      TambahJualD
                      Text1.Text = ""
                    End If
                  Else
                    DiscProsen = 0
                    DiscRp = 0
                    DiscInternRp = 0
                    DiscInternProsen = 0
                    IsDiscBySupplier = False
                  TambahJualD
                  Text1.Text = ""
                  End If
                Else
                    DiscProsen = 0
                    DiscRp = 0
                    DiscInternRp = 0
                    DiscInternProsen = 0
                    IsDiscBySupplier = False
                  TambahJualD
                  Text1.Text = ""
                End If
            End If
            'ShowData txtTerm, (StrConv(Buffer, vbUnicode))
End Select
End Sub

Private Sub Text1_DblClick()
  If MsgBox("Keluar aplikasi?", vbOKCancel + vbQuestion, "V POS") = vbOK Then
    End
  End If
End Sub
Sub AmbilDuit()
Dim JumlahDuit As Double

End Sub

Private Sub TambahkanBarangPDP()
Dim SQL As String
Dim DB As Database
Dim rst As Recordset
Dim Jumlah As Double, JumlahQtyPDP As Double
Dim SelBks
Err.Clear
On Error GoTo Trace:
Set DB = OpenDatabase(App.path & "\DATABASE\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
Set rst = DB.OpenRecordset("SELECT SUM(Jumlah) AS Jml FROM MSALESD WHERE IIF(IsNull(IsPDP), FALSE, IsPDP)=FALSE AND IDSALES=" & IDSales)
If Not (rst.BOF Or rst.EOF) Then
  Jumlah = 0
  rst.MoveFirst
  Do While Not rst.EOF
    Jumlah = Jumlah + NullToNol(rst!Jml)
    rst.MoveNext
  Loop
  rst.Close
  DB.Close
End If

If defMinialBelanjaDapatPDP <= 0 Then
  frmPesan.lbPesan.Caption = "Minimal belanja PDP dinonaktifkan."
  frmPesan.Show 1
ElseIf Jumlah >= defMinialBelanjaDapatPDP Then
  Dim IsAmbil As Boolean
  Dim kodeBrg As String
  Dim idBrg As Long
  If Ditutup Then Exit Sub
    frmBarangPDP.IsAllow = True
    frmBarangPDP.Tampil IsAmbil, idBrg, kodeBrg
    If IsAmbil And kodeBrg <> "" Then
      rs.Index = "PrimaryKey"
      rs.Seek "=", idBrg
      If rs.NoMatch Then
      Else
        JumlahQtyPDP = 0
        Set DB = OpenDatabase(App.path & "\DATABASE\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
        Set rst = DB.OpenRecordset("SELECT * FROM MSalesD WHERE IDSales=" & IDSales)
        If Not (rst.BOF Or rst.EOF) Then
          rst.MoveFirst
          Do While Not rst.EOF
            If NullToBool(rst!IsPDP) And rst!IdInventor = rs!IdInventor And rst!IDInvSat = rs!NoID Then
              JumlahQtyPDP = JumlahQtyPDP + rst!Qty
            End If
            rst.MoveNext
          Loop
          rst.Close
          DB.Close
          If lbBintang.Visible Or Ditutup Or JumlahQtyPDP + Qty > Int(Jumlah / defMinialBelanjaDapatPDP) Then Exit Sub
          If IDMember >= 1 Then
            HargaJualKhusus = NullToNol(rs!HargaJual) - NullToNol(rs!DiscPDPMember)
          Else
            HargaJualKhusus = NullToNol(rs!HargaJual) - NullToNol(rs!DiscPDP)
          End If
          TambahJualDKhususPDP
          Data1.Recordset.Edit
          Data1.Recordset!HargaNormal = NullToNol(rs!HargaJual)
'          Data1.Recordset!HargaBruto = HargaJualKhusus
'          Data1.Recordset!harga = HargaJualKhusus
'          Data1.Recordset!DiscRp = 0
'          Data1.Recordset!DiscProsen = 0
'          Data1.Recordset!DiscInternRp = 0
'          Data1.Recordset!DiscInternProsen = 0
'          Data1.Recordset!Jumlah = Bulatkan(Data1.Recordset!Qty * HargaJualKhusus, 0)
          Data1.Recordset.Update
          Data1.Recordset.Bookmark = Data1.Recordset.LastModified
          Set SelBks = TDBGrid1.SelBookmarks
          While SelBks.Count <> 0
              SelBks.Remove 0
          Wend
          TDBGrid1.SelBookmarks.Add TDBGrid1.Bookmark
          If Data1.Recordset!IsMember Then
            cetakdetil Data1.Recordset!KodeInv, Data1.Recordset!NamaInv, Format(Data1.Recordset!Qty, "##0"), Format(Data1.Recordset!harga, "###,###,##0"), Format(Qty * Data1.Recordset!harga, "###,###,##0")
          Else
            cetakdetil Data1.Recordset!KodeInv, Data1.Recordset!NamaInv & "*)", Format(Data1.Recordset!Qty, "##0"), Format(Data1.Recordset!harga, "###,###,##0"), Format(Qty * Data1.Recordset!harga, "###,###,##0")
          End If
          DisplayPesan Data1.Recordset!NamaInv, Format(Qty, "##0") & " X    " & Format(Data1.Recordset!harga, "###,###,##0")
          Qty = 1
          lbQty = Format(Qty, "##0") & " X"
          HitungSubTotal
        End If
        Text1.Text = ""
        Text1.SetFocus
      End If
    End If
Else
  frmPesan.lbPesan.Caption = "Jumlah pembelian belum memenuhi Minimal belanja PDP."
  frmPesan.Show 1
End If
Trace:
  If Err.Number <> 0 Then
    MsgBox "Error : " & Err.Number & ", " & Err.Description, vbCritical, "VPOS"
    Err.Clear
  End If
  Set DB = Nothing
  Set rst = Nothing
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error Resume Next
Dim kodeBrg As String
Dim Kodebarang As String
Dim JumlahBrg As Double
Dim IsAmbil As Boolean
Dim Hasil As String
Dim jawaban As Boolean
Dim SelBks
Hasil = Trim(SendByCode(KeyCode))
KeyCode = 0
If lbBintang.Visible And (Hasil <> "KMS" And Hasil <> "PMT" And Hasil <> "." And Hasil <> "BKS" And Hasil <> "RND" And Hasil <> "SCN" And Hasil <> "SCA" And Hasil <> "MBR" And Hasil <> "STL" And Hasil <> "CSH" And Hasil <> "DBR" And Hasil <> "DBP" And Hasil <> "DIR" And Hasil <> "DIP" And Hasil <> "RPT" And Hasil <> "VCR" And Hasil <> "0" And Hasil <> "." And Hasil <> "00" And Hasil <> "1" And Hasil <> "2" And Hasil <> "3" And Hasil <> "4" And Hasil <> "5" And Hasil <> "6" And Hasil <> "7" And Hasil <> "8" And Hasil <> "9" And Hasil <> "-") And Hasil <> "CLR" And Hasil <> "ENT" And Hasil <> "BK1" And Hasil <> "BK2" Then Exit Sub
If Ditutup And Hasil <> "RPT" And Hasil <> "CLR" Then Exit Sub

Select Case Hasil
Case "PDP"
  TambahkanBarangPDP
Case "TKP"
  If isRemcomendedOnline Then
    frmTukarPoin.Show 1
  End If
Case "KMS"
      frmAgen.Tampil jawaban, 1, Agen, KomisiProsen
      TampilBawah
Case "PMT"
      jawaban = False
      frmKey.Tampil jawaban, 13
      If jawaban = True Then
        frmAmbilDuit.Tampil jawaban, 13
      End If
Case "DN"
    'Data1.Refresh
    If Data1.Recordset.BOF Or Data1.Recordset.EOF Then Exit Sub
    Data1.Recordset.MoveNext
    If Data1.Recordset.EOF Then
      Data1.Recordset.MovePrevious
    End If
    Set SelBks = TDBGrid1.SelBookmarks
    While SelBks.Count <> 0
        SelBks.Remove 0
    Wend
    TDBGrid1.SelBookmarks.Add TDBGrid1.Bookmark
Case "UP"
    'Data1.Refresh
    If Data1.Recordset.BOF Or (Data1.Recordset.EOF And Data1.Recordset.BOF) Then Exit Sub
    Data1.Recordset.MovePrevious
    If Data1.Recordset.BOF Then
      Data1.Recordset.MoveNext
    End If
    Set SelBks = TDBGrid1.SelBookmarks
    While SelBks.Count <> 0
        SelBks.Remove 0
    Wend
    TDBGrid1.SelBookmarks.Add TDBGrid1.Bookmark
Case "*"
    If Len(Text1.Text) < 8 Then
      If IsNumeric(Text1.Text) Then
        Qty = CCur(Text1.Text)
        If Qty = 0 Then Qty = 1
      Else
        Qty = 1
      End If
      lbQty = Format(Qty, "##0.###") & " X"
      Text1.Text = ""
    End If
Case "0", "00", "1", "2", "3", "4", "5", "6", "7", "8", "9", "*", "`", "!", "@", "#", "$", "%", "^", "&", "<", "(", ")", "-", "+", "{", "}", "[", "]", "/", "?", ".", ","
  Text1.Text = Text1.Text & Hasil
  If lbBintang.Visible And IsNumeric(Text1.Text) Then
    Text1.Text = Format(CCur(Text1.Text), "###,##0")
  End If
  Text1.SelStart = Len(Text1.Text)
Case "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"
  Text1.Text = Text1.Text & Hasil
  Text1.SelStart = Len(Text1.Text)
Case "CLR"
  cmdHOME_Click
Case "BKS"
  If Len(Text1.Text) > 0 Then
    Text1.Text = Left(Text1.Text, Len(Text1.Text) - 1)
    If lbBintang.Visible And IsNumeric(Text1.Text) Then
      Text1.Text = Format(CCur(Text1.Text), "###,##0")
    End If
    Text1.SelStart = Len(Text1.Text)
  End If
Case "MBR"
    cmdF6_Click
    HitungPoin
    TampilBawah
    DisplayPesan "DISC NOTA #" & Space(9 - IIf(Len(Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0")) > 9, 0, Len(Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0")))) & Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0"), "TOTAL     #" & Space(9 - IIf(Len(Format(Total, "###,###,##0")) > 9, 0, Len(Format(Total, "###,###,##0")))) & Format(Total, "###,###,##0")
    UpdateSales
Case "PLU"
If isSupervisor Then
        frmPesan.lbPesan = "Hanya Kasir Yang Boleh!!!!"
        frmPesan.Show 1
        Exit Sub
    End If
    If lbBintang.Visible Or Ditutup Then Exit Sub
    'Request Tidak ada Customer maka tidak boleh transaksi
    If KodeMember = "" Then Exit Sub
    If Text1.Text = "" Then Exit Sub
    rs.Index = "Barcode"
    rs.Seek "=", Text1.Text
    If rs.NoMatch Then
      rs.Index = "Kode"
      rs.Seek "=", Text1.Text
      If rs.NoMatch Then
        rs.Index = "Nama"
        rs.Seek "=", Text1.Text
        If rs.NoMatch Then
        Else
            DiscProsen = 0
            DiscRp = 0
            DiscInternRp = 0
            DiscInternProsen = 0
            IsDiscBySupplier = False

          TambahJualD
          Text1.Text = ""
        End If
      Else
        DiscProsen = 0
        DiscRp = 0
        DiscInternRp = 0
        DiscInternProsen = 0
        IsDiscBySupplier = False
      
      TambahJualD
      Text1.Text = ""
      End If
    Else
        DiscProsen = 0
        DiscRp = 0
        DiscInternRp = 0
        DiscInternProsen = 0
        IsDiscBySupplier = False

      TambahJualD
      Text1.Text = ""
    End If
Case "PE"
'    If lbBintang.Visible Or Ditutup Then Exit Sub
'    frmsetHargadanKode.Tampil HargaJualKhusus, KodeKhusus
'    jawab = False
'    frmKey.Tampil jawab, 114
'    If jawab Then
'      If HargaJualKhusus = -1 Then Exit Sub
'      rs.Index = "Barcode"
'      rs.Seek "=", KodeKhusus
'      If rs.NoMatch Then
'        rs.Index = "Kode"
'        rs.Seek "=", KodeKhusus
'        If rs.NoMatch Then
'          rs.Index = "Nama"
'          rs.Seek "=", KodeKhusus
'          If rs.NoMatch Then
'          Else
'              DiscProsen = 0
'              DiscRp = 0
'              DiscInternRp = 0
'              DiscInternProsen = 0
'              IsDiscBySupplier = False
'
'            TambahJualDKhusus
'            Text1.Text = ""
'          End If
'        Else
'          DiscProsen = 0
'          DiscRp = 0
'          DiscInternRp = 0
'          DiscInternProsen = 0
'          IsDiscBySupplier = False
'
'        TambahJualDKhusus
'        Text1.Text = ""
'        End If
'      Else
'          DiscProsen = 0
'          DiscRp = 0
'          DiscInternRp = 0
'          DiscInternProsen = 0
'          IsDiscBySupplier = False
'
'        TambahJualDKhusus
'        Text1.Text = ""
'      End If
'    End If

  If lbBintang.Visible Or Ditutup Then Exit Sub
'    If frmsetHarga.Tampil(HargaJualKhusus, KodeKhusus) Then
    frmsetHarga.Tampil HargaJualKhusus, KodeKhusus
    If HargaJualKhusus = -1 Then Exit Sub
    'If TipeHargaJual = 2 Or (TipeHargaJual <> 2 And HargaJualKhusus >= NullToNol(Data1.Recordset!HargaMin)) Then
      jawab = False
      frmKey.Tampil jawab, 114
      If jawab Then
        Data1.Recordset.Edit
        Data1.Recordset!HargaBruto = HargaJualKhusus
        Data1.Recordset!harga = HargaJualKhusus
        Data1.Recordset!DiscRp = 0
        Data1.Recordset!DiscProsen = 0
        Data1.Recordset!DiscInternRp = 0
        Data1.Recordset!DiscInternProsen = 0
        Data1.Recordset!Jumlah = Bulatkan(Data1.Recordset!Qty * HargaJualKhusus, 0)
        Data1.Recordset.Update
  '        Data1.Refresh
        Data1.Recordset.Bookmark = Data1.Recordset.LastModified
        If Data1.Recordset!IsMember Then
          cetakdetil Data1.Recordset!KodeInv, Data1.Recordset!NamaInv, Format(Data1.Recordset!Qty, "##0"), Format(Data1.Recordset!harga, "###,###,##0"), Format(Qty * Data1.Recordset!harga, "###,###,##0")
        Else
          cetakdetil Data1.Recordset!KodeInv, Data1.Recordset!NamaInv & "*)", Format(Data1.Recordset!Qty, "##0"), Format(Data1.Recordset!harga, "###,###,##0"), Format(Qty * Data1.Recordset!harga, "###,###,##0")
        End If
        DisplayPesan Data1.Recordset!NamaInv, Format(Qty, "##0") & " X    " & Format(Data1.Recordset!harga, "###,###,##0")
        Qty = 1
        lbQty = Format(Qty, "##0") & " X"
        HitungSubTotal
      End If
'    Else
'      jawab = False
'      frmKey.Tampil jawab, 114
'      If jawab Then
'        Data1.Recordset.Edit
'        Data1.Recordset!HargaBruto = HargaJualKhusus
'        Data1.Recordset!harga = HargaJualKhusus
'        Data1.Recordset!DiscRp = 0
'        Data1.Recordset!DiscProsen = 0
'        Data1.Recordset!DiscInternRp = 0
'        Data1.Recordset!DiscInternProsen = 0
'        Data1.Recordset.Update
'        '        Data1.Refresh
'        Data1.Recordset.Bookmark = Data1.Recordset.LastModified
'        If Data1.Recordset!IsMember Then
'          cetakdetil Data1.Recordset!KodeInv, Data1.Recordset!NamaInv, Format(Data1.Recordset!QTY, "##0"), Format(Data1.Recordset!harga, "###,###,##0"), Format(QTY * Data1.Recordset!harga, "###,###,##0")
'        Else
'          cetakdetil Data1.Recordset!KodeInv, Data1.Recordset!NamaInv & "*)", Format(Data1.Recordset!QTY, "##0"), Format(Data1.Recordset!harga, "###,###,##0"), Format(QTY * Data1.Recordset!harga, "###,###,##0")
'        End If
'        DisplayPesan Data1.Recordset!NamaInv, Format(QTY, "##0") & " X    " & Format(Data1.Recordset!harga, "###,###,##0")
'        QTY = 1
'        lbQTY = Format(QTY, "##0") & " X"
'        HitungSubTotal
'      End If
'    End If
    Text1.Text = ""
    Text1.SetFocus
Case "RBY"
         jawab = False
         frmKey.Tampil jawab, 116
         If jawab Then
        frmRevisiBayar.Show 1
    End If
Case "HLD"
    cmdDEL_Click
Case "SCK"
    cmdF1_Click
Case "SCN"
    KodeSearch = "SCN"
    cmdF2_Click
Case "SCA"
    KodeSearch = "SCA"
    cmdF2_Click
    
Case "VOD"
'If IsNotaDariPending = True Then
'      frmPesan.lbPesan = "Nota Pending Tidak bisa Void!!!"
'      frmPesan.Show 1
'Else
      If Data1.Recordset.EOF And Data1.Recordset.BOF Then
      Else
'hilang        MSComm1.PortOpen = False
        frmsetItemCorrect.Tampil JumlahBrg, Kodebarang
'hilang        MSComm1.PortOpen = True
        If Kodebarang = "-1" Then Exit Sub
        'Data1.Recordset.FindFirst "KodeINV='" & Replace(kodeBarang, "'", "''") & "'"
        'If Data1.Recordset.NoMatch Then
         ' qTY = -1
        'Else
         ' qTY = -1 * Data1.Recordset!qTY
        'End If
          Qty = -1 * Abs(JumlahBrg)
          lbQty = Format(Qty, "##0") & " X"
          If Qty < 0 Then
            jawab = False
            frmKey.Tampil jawab, 114
            If jawab Then
              KeyCode = 0
              rs.Index = "Barcode"
              rs.Seek "=", Kodebarang
              If rs.NoMatch Then
                rs.Index = "Kode"
                rs.Seek "=", Kodebarang
                If rs.NoMatch Then
                  rs.Index = "Nama"
                  rs.Seek "=", Kodebarang
                  If rs.NoMatch Then
                  Else
                    ITEMKoreksi "VOD"
                  End If
                Else
                ITEMKoreksi "VOD"
                End If
              Else
                ITEMKoreksi "VOD"
              End If
              Else
                Qty = 1
                lbQty = Format(Qty, "##0") & " X"
                KeyCode = 0
            End If 'Jawab
          Else 'qty
            Qty = 1
            lbQty = Format(Qty, "##0") & " X"
            KeyCode = 0
          End If 'qty
'      End If
End If 'Nota dari Pending

Case "CRC"
If IsNotaDariPending = True Then
        frmPesan.lbPesan = "Nota Pending Tidak bisa Correct!!!"
        frmPesan.Show 1
Else
      If Data1.Recordset.EOF And Data1.Recordset.BOF Then
      Else
        Data1.Recordset.MoveLast
        Qty = -1 * Data1.Recordset!Qty
        lbQty = Format(Qty, "##0") & " X"
        If Qty < 0 Then
          jawab = False
          frmKey.Tampil jawab, 115
          If jawab Then
            KeyCode = 0
            rs.Index = "Kode"
            rs.Seek "=", Data1.Recordset!KodeInv
            If rs.NoMatch Then
            Else
              ITEMKoreksi "CRC"
            End If
          End If
        Else
          KeyCode = 0
        End If
      End If
End If
Case "AVD"
'If IsNotaDariPending = True Then
'        frmPesan.lbPesan = "Nota Pending Tidak bisa All Void!!!"
'        frmPesan.Show 1
'Else
       If Data1.Recordset.EOF And Data1.Recordset.BOF Then
       Else
            Dim dbHistori As Database
            Dim RsHistori As Recordset
         jawab = False
         frmKey.Tampil jawab, 116
         If jawab Then
           Data1.Recordset.MoveFirst
           Do While Not Data1.Recordset.EOF
           Data1.Recordset.Edit
           Data1.Recordset!Qty = 0
           Data1.Recordset!Jumlah = 0
           Data1.Recordset!Transaksi = "AVD"
           Data1.Recordset.Update
           Data1.Recordset.Bookmark = Data1.Recordset.LastModified
           'SIMPAN HISTORI
               Set dbHistori = OpenDatabase(DirDatabase & "\Histori.mdb")
                Set RsHistori = dbHistori.OpenRecordset("Historikasir")
                RsHistori.AddNew
                'ID
                RsHistori!kassa = NamaMesin
                RsHistori!IDUser = IDUser
                RsHistori!KodeUser = KodeKasir
                RsHistori!Tanggal = Date
                RsHistori!Jam = Time
                RsHistori!IDSales = IDSales
                RsHistori!IDSalesD = IDSalesD
                RsHistori!Transaksi = "AVD"
                RsHistori!IdInventor = Data1.Recordset!IdInventor
                RsHistori!idSatuan = Data1.Recordset!idSatuan
                RsHistori!KodeInventor = Data1.Recordset!KodeInv
                RsHistori!NamaInventor = Data1.Recordset!NamaInv
                RsHistori!HargaPokok = Data1.Recordset!HargaPokok
                RsHistori!HargaJualMaster = Data1.Recordset!harga
                RsHistori!HargaJualKhusus = Data1.Recordset!harga
                RsHistori.Update
                       
                Data1.Recordset.MoveNext
           Loop
           RsHistori.Close
           dbHistori.Close
           KeyCode = 0
           IDBank = 0
           IDJenisKartu = 0
           IDBankServer = 0
           Dibayar = 0
           SubTotal = 0
           Kembali = 0
           Voucher = 0
           Disc = 0
           DiscINTERNNOTA = 0
           PotonganPembulatan = 0
           JumDiscInternRp = 0
           HitungSubTotal
          DisplayPesan "TRANSAKSI DIBATALKAN", "      ALL VOID      "
           Prin "   -ALL VOID - ALL VOID - ALL VOID-    " & Chr(13) & Chr(10) & "---------- STRUK HARAP DISOBEK---------" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
            Chr(13) & Chr(10)
           If IsHematKertas Then
            Prin Chr(13) & Chr(10)
            DoEvents
            PrinBigChar Space(15 - Min(15, Len(NamaToko))) & UCase(NamaToko)
            DoEvents
            Prin Judulstruk & Chr(13) & Chr(10) & "---------------------------------------" & Chr(13) & Chr(10)
            DoEvents
           End If
          papercut
          Ditutup = True
          CetakStruck IDSales, False, , True
          End If
        End If
'    End If
Case "RTN"
    frmsetItemCorrect.Tampil JumlahBrg, Kodebarang
'Hilang    MSComm1.PortOpen = True
    If Kodebarang = "-1" Then Exit Sub
    Qty = -1 * Abs(JumlahBrg)
    lbQty = Format(Qty, "##0") & " X"
    If Qty < 0 Then
      jawab = True
'      frmKey.Tampil jawab, 114
      If jawab Then
        KeyCode = 0
        rs.Index = "Barcode"
        rs.Seek "=", Kodebarang
        If rs.NoMatch Then
          rs.Index = "Kode"
          rs.Seek "=", Kodebarang
          If rs.NoMatch Then
            rs.Index = "Nama"
            rs.Seek "=", Kodebarang
            If rs.NoMatch Then
            Else
              ITEMKoreksi "RTN"
            End If
          Else
          ITEMKoreksi "RTN"
          End If
        Else
          ITEMKoreksi "RTN"
        End If
      End If 'Jawab
    Else 'qty
      KeyCode = 0
    End If 'qty
Case "ENT"
  If lbBintang.Visible Then 'enter
''      If Len(Text1.Text) > 20 Then
'        'lbBintang.Visible = False
'        Bank = Total - Dibayar
'        Dibayar = Total
'        TampilBawah
'        If Kembali >= 0 Then Ditutup = True
'        Text1.Text = ""
'        lbBintang.Visible = False
'        openDrawer
'        HitungItem
'        DisplayPesan "PAY# " & Space(15 - Min(Len(Format(Dibayar - Voucher, "###,###,##0")), 15)) & Format(Dibayar - Voucher, "###,###,##0"), "CHG# " & Space(15 - Max(Len(Format(Kembali, "###,###,##0")), 15)) & Format(Kembali, "###,###,##0")
'        If Kembali >= 0 Then
'            Prin "---------------------------------------" & Chr(13) & Chr(10) & _
'              "Subtotal" & Space(31 - Len(Format(SubTotal, "###,###,##0"))) & Format(SubTotal, "###,###,##0") & Chr(13) & Chr(10) & _
'              IIf(TotalDiscBrg + JumDiscInternRp = 0, "", "Discount Barang" & Space(24 - Len(Format(TotalDiscBrg + JumDiscInternRp, "###,###,##0"))) & Format(TotalDiscBrg + JumDiscInternRp, "###,###,##0") & Chr(13) & Chr(10)) & _
'              IIf(Disc + DiscINTERNNOTA + PotonganPembulatan = 0, "", "Discount Nota" & Space(26 - Len(Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0"))) & Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0") & Chr(13) & Chr(10)) & _
'              "Voucher " & Space(31 - Len(Format(Voucher, "###,###,##0"))) & Format(Voucher, "###,###,##0") & Chr(13) & Chr(10) & _
'              "Dibayar " & Space(31 - Len(Format(Dibayar - Voucher, "###,###,##0"))) & Format(Dibayar - Voucher, "###,###,##0") & Chr(13) & Chr(10) & _
'              "Kembali " & Space(31 - Len(Format(Kembali, "###,###,##0"))) & Format(Kembali, "###,###,##0") & Chr(13) & Chr(10) & _
'              IIf(ISSembunyikanFooterbarangKSB, "", "Belanja Barang POIN" & Space(20 - Len(Format(BelanjaPoin, "###,###,##0"))) & Format(BelanjaPoin, "###,###,##0") & Chr(13) & Chr(10)) & _
'              "#" & Format(CLNG(NONOTA), "0#####") & "," & "Kasir" & NamaMesin & ":" & Mid(NamaKasir, 1, 6) & Space(6 - IIf(Len(NamaKasir) <= 6, Len(NamaKasir), 6)) & Format(Date, " dd/MM/yy") & Format(Time, "hh:mm:ss") & Chr(13) & Chr(10) & _
'              "Macam Barang =" & Format(MacamItem, "##0") & " , Total Item =" & Format(Jumlahitem, "##0") & Chr(13) & Chr(10) & _
'              IIf(KodeMember = "", "", "CUSTOMER: " & KodeMember & "-" & NamaMember & Chr(13) & Chr(10)) & _
'              "----TERIMA KASIH ATAS KUNJUNGAN ANDA----" & Chr(13) & Chr(10) & _
'              IIf(Footer1 = "", "", Footer1 & Chr(13) & Chr(10)) & _
'              IIf(Footer2 = "", "", Footer2 & Chr(13) & Chr(10)) & _
'              IIf(Footer3 = "", "", Footer3 & Chr(13) & Chr(10)) & _
'              Chr(13) & Chr(10)
'              If IsHematKertas Then
'                PrinBigChar Space(15 - Min(15, Len(NamaToko))) & UCase(NamaToko)
'                Prin Judulstruk & Chr(13) & Chr(10) & "---------------------------------------"
'              End If
'            papercut
'        Else
'          Prin "---------------------------------------" & Chr(13) & Chr(10) & "Dibayar " & Space(31 - Len(Format(Dibayar - Voucher, "###,###,##0"))) & Format(Dibayar - Voucher, "###,###,##0") & Chr(13) & Chr(10) & _
'               "Kembali " & Space(31 - Len(Format(Kembali, "###,###,##0"))) & Format(Kembali, "###,###,##0")
'        End If
''      End If
    ElseIf Len(Text1.Text) >= 8 Then
'KHUSUS BILKA MINTA OTOMATIS TAPI BIKIN PUSING TEMPAT LAIN
    'If Len(Text1.Text) = 12 Then Text1.Text = "0" & Text1.Text
    If Len(Text1.Text) = 12 Then
      If Left(Text1.Text, 2) = "25" Or Left(Text1.Text, 6) = "271254" Then 'Buah dan sayur dan telor
          Kodebarang = Left(Text1.Text, 6)
          Qty = CLng(Right(Left(Text1.Text, 11), 5)) / 1000#
      Else
          Kodebarang = Text1.Text
      End If
    ElseIf Len(Text1.Text) = 13 Then
      If Left(Text1.Text, 3) = "025" Or Left(Text1.Text, 7) = "0271254" Then 'Buah dan sayur
          Kodebarang = Right(Left(Text1.Text, 7), 6)
          Qty = CLng(Right(Left(Text1.Text, 12), 5)) / 1000#
      Else
          Kodebarang = Text1.Text
      End If
    Else
      Kodebarang = Text1.Text
    End If

      dbs.Recordsets.Refresh
      rs.Index = "Barcode"
      rs.Seek "=", Kodebarang
      If rs.NoMatch Then
        rs.Index = "Kode"
        rs.Seek "=", Kodebarang
        If rs.NoMatch Then
          rs.Index = "Nama"
          rs.Seek "=", Kodebarang
          If rs.NoMatch Then
          Else 'Nama Cocok
              DiscProsen = 0
                DiscRp = 0
                DiscInternRp = 0
                DiscInternProsen = 0
                IsDiscBySupplier = False

            TambahJualD
            Text1.Text = ""
          End If
          Text1.Text = ""
        Else 'Kode Cocok
            DiscProsen = 0
            DiscRp = 0
            DiscInternRp = 0
            DiscInternProsen = 0
            IsDiscBySupplier = False
        
        TambahJualD
        Text1.Text = ""
        End If
      Else 'Barcode Cocok
        DiscProsen = 0
        DiscRp = 0
        DiscInternRp = 0
        DiscInternProsen = 0
        IsDiscBySupplier = False
        
        TambahJualD
        Text1.Text = ""
      End If
    ElseIf Not lbBintang.Visible And UCase(Text1.Text) = "PRINT" Then
      jawab = False
     ' frmKey.Tampil jawab, 144
      jawab = True
      If jawab Then
        Dim NoIDSales As Long
        If Ditutup Then Exit Sub
        frmLookUpSales.NamaField = "Kode"
        frmLookUpSales.Tampil IsAmbil, NoIDSales
        If IsAmbil And NoIDSales <> 0 Then
          CetakStruck NoIDSales, True
          If isRemcomendedOnline Then
            CetakStruckReedem2 NoIDSales, True
          End If
        End If
      End If
      Text1.Text = ""
'      Text1.SetFocus
    ElseIf Not lbBintang.Visible And UCase(Text1.Text) = "PRNTKP" Then
      'MsgBox "Fitur masih disempurnakan, Terima kasih."
      jawab = False
     ' frmKey.Tampil jawab, 144
      jawab = True
      If jawab And isRemcomendedOnline And Data1.Recordset.RecordCount = 0 Then
        Dim NoIDTKP As Long
        If Ditutup Then Exit Sub
        frmLookUpTKPSQLServer.NamaField = "Keterangan"
        frmLookUpTKPSQLServer.Tampil IsAmbil, NoIDTKP
        If IsAmbil And NoIDTKP <> 0 Then
          CetakStruckTKP NoIDTKP, "COPY"
        End If
      End If
      Text1.Text = ""
    ElseIf Not lbBintang.Visible Then
      rs.Index = "Kode"
      rs.Seek "=", Text1.Text
      If rs.NoMatch Then
        rs.Index = "Barcode"
        rs.Seek "=", Text1.Text
        If rs.NoMatch Then
          rs.Index = "Nama"
          rs.Seek "=", Text1.Text
          If rs.NoMatch Then
          Else
              DiscProsen = 0
                DiscRp = 0
                DiscInternRp = 0
                DiscInternProsen = 0
                IsDiscBySupplier = False
            TambahJualD
            Text1.Text = ""
          End If
          Text1.Text = ""
        Else
            DiscProsen = 0
            DiscRp = 0
            DiscInternRp = 0
            DiscInternProsen = 0
            IsDiscBySupplier = False
        
        TambahJualD
        Text1.Text = ""
        End If
      Else
        DiscProsen = 0
        DiscRp = 0
        DiscInternRp = 0
        DiscInternProsen = 0
        IsDiscBySupplier = False
        
        TambahJualD
        Text1.Text = ""
      End If
    End If
Case "ESC"
    If Data1.Recordset.RecordCount >= 1 Then
      MsgBox "Yakin ingin keluar?" & vbCrLf & "Pending atau Selesaikan Transaksi terlebih Dahulu.", vbInformation + vbOKOnly, "VPOS"
    Else
      Unload Me
      Form2.Show
      Form2.SetFocus
    End If
Case "STL"
    bolehbergerak = False
    If Not lbBintang.Visible Then
      UpdatekanDiskonPerBarang
    End If
    cmdPGDN_Click
    Timer2.Enabled = False
Case "CSH"
    cmdPGUP_Click
Case "CHG"
    frmEntriBarang.Show 1
Case "BK1"
    TampilBank False
Case "BK2" 'Kartu Kredit
    TampilBank True
'     UpdateFooter
'    If Ditutup = False And lbBintang.Visible Then
'    frmBank.Tampil jawaban, IDBank, IDBankServer, 67, CLng((Total - Dibayar) * (1# + PersenBiayaKartuKredit / 100))
'    If jawaban Then
'    Bank = ((Total - Dibayar) * (1# + PersenBiayaKartuKredit / 100))
'    ISCreditCard = True
'    BiayaCC = ((Total - Dibayar) * (PersenBiayaKartuKredit / 100))
'        Dibayar = Total
'        TampilBawah
'        If Kembali >= 0 Then Ditutup = True
'        Text1.Text = ""
'        openDrawer
'        HitungItem
'        DisplayPesan "PAY# " & Space(15 - Min(Len(Format(Dibayar + BiayaCC - Voucher - ReedemNilai, "###,###,##0")), 15)) & Format(Dibayar + BiayaCC - Voucher - ReedemNilai, "###,###,##0"), "CHG# " & Space(15 - Max(Len(Format(Kembali, "###,###,##0")), 15)) & Format(Kembali, "###,###,##0")
'        If Kembali >= 0 Then
'        If ISCreditCard Then
'        Prin "---------------------------------------" & Chr(13) & Chr(10) & _
'                "Subtotal" & Space(31 - Len(Format(SubTotal, "###,###,##0"))) & Format(SubTotal, "###,###,##0") & Chr(13) & Chr(10) & _
'                IIf(TotalDiscBrg = 0, "", "Discount Barang" & Space(24 - Len(Format(TotalDiscBrg, "###,###,##0"))) & Format(TotalDiscBrg, "###,###,##0") & Chr(13) & Chr(10)) & _
'                IIf(Disc = 0, "", "Discount Nota" & Space(26 - Len(Format(Disc, "###,###,##0"))) & Format(Disc, "###,###,##0") & Chr(13) & Chr(10)) & _
'              "Voucher " & Space(31 - Len(Format(Voucher, "###,###,##0"))) & Format(Voucher, "###,###,##0") & Chr(13) & Chr(10) & _
'              "Dibayar (+" & Trim(CStr(PersenBiayaKartuKredit)) & "% CC)" & Space(24 - Len(Trim(CStr(PersenBiayaKartuKredit))) - Len(Format(Dibayar + BiayaCC - Voucher - ReedemNilai, "###,###,##0"))) & Format(Dibayar + BiayaCC - Voucher - ReedemNilai, "###,###,##0") & Chr(13) & Chr(10) & _
'              "Kembali " & Space(31 - Len(Format(Kembali, "###,###,##0"))) & Format(Kembali, "###,###,##0") & Chr(13) & Chr(10) & _
'              "#" & Format(CLNG(NONOTA), "0#####") & "," & "Kasir" & NamaMesin & ":" & Mid(NamaKasir, 1, 6) & Space(6 - IIf(Len(NamaKasir) <= 6, Len(NamaKasir), 6)) & Format(Date, " dd/MM/yy") & Format(Time, "hh:mm:ss") & Chr(13) & Chr(10) & _
'              "Macam Barang =" & Format(MacamItem, "##0") & " , Total Item =" & Format(Jumlahitem, "##0") & Chr(13) & Chr(10) & "----TERIMA KASIH ATAS KUNJUNGAN ANDA----" & Chr(13) & Chr(10) & _
'              IIf(Footer1 = "", "", Footer1 & Chr(13) & Chr(10)) & _
'              IIf(Footer2 = "", "", Footer2 & Chr(13) & Chr(10)) & _
'              IIf(Footer3 = "", "", Footer3 & Chr(13) & Chr(10)) & _
'              Chr(13) & Chr(10)
'
'
'        Else
'            Prin "---------------------------------------" & Chr(13) & Chr(10) & _
'                "Subtotal" & Space(31 - Len(Format(SubTotal, "###,###,##0"))) & Format(SubTotal, "###,###,##0") & Chr(13) & Chr(10) & _
'                IIf(TotalDiscBrg = 0, "", "Discount Barang" & Space(24 - Len(Format(TotalDiscBrg, "###,###,##0"))) & Format(TotalDiscBrg, "###,###,##0") & Chr(13) & Chr(10)) & _
'                IIf(Disc = 0, "", "Discount Nota" & Space(26 - Len(Format(Disc, "###,###,##0"))) & Format(Disc, "###,###,##0") & Chr(13) & Chr(10)) & _
'              "Voucher " & Space(31 - Len(Format(Voucher, "###,###,##0"))) & Format(Voucher, "###,###,##0") & Chr(13) & Chr(10) & _
'              "Dibayar " & Space(31 - Len(Format(Dibayar - Voucher, "###,###,##0"))) & Format(Dibayar - Voucher, "###,###,##0") & Chr(13) & Chr(10) & _
'              "Kembali " & Space(31 - Len(Format(Kembali, "###,###,##0"))) & Format(Kembali, "###,###,##0") & Chr(13) & Chr(10) & _
'              "#" & Format(CLNG(NONOTA), "0#####") & "," & "Kasir" & NamaMesin & ":" & Mid(NamaKasir, 1, 6) & Space(6 - IIf(Len(NamaKasir) <= 6, Len(NamaKasir), 6)) & Format(Date, " dd/MM/yy") & Format(Time, "hh:mm:ss") & Chr(13) & Chr(10) & _
'              "Macam Barang =" & Format(MacamItem, "##0") & " , Total Item =" & Format(Jumlahitem, "##0") & Chr(13) & Chr(10) & "----TERIMA KASIH ATAS KUNJUNGAN ANDA----" & Chr(13) & Chr(10) & _
'              IIf(Footer1 = "", "", Footer1 & Chr(13) & Chr(10)) & _
'              IIf(Footer2 = "", "", Footer2 & Chr(13) & Chr(10)) & _
'              IIf(Footer3 = "", "", Footer3 & Chr(13) & Chr(10)) & _
'              Chr(13) & Chr(10)
'
'        End If
'            PrinBigChar Space(15 - Min(15, Len(NamaToko))) & UCase(NamaToko)
'            Prin Judulstruk & Chr(13) & Chr(10) & "---------------------------------------"
'            papercut
'        lbBintang.Visible = False
'        TutupCommBarcode
'        Else
'          Prin "---------------------------------------" & Chr(13) & Chr(10) & "Dibayar " & Space(31 - Len(Format(Dibayar - Voucher, "###,###,##0"))) & Format(Dibayar - Voucher, "###,###,##0") & Chr(13) & Chr(10) & _
'               "Kembali " & Space(31 - Len(Format(Kembali, "###,###,##0"))) & Format(Kembali, "###,###,##0")
'        End If
'        DoEvents
'        KirimKeServer IDSales
'      End If
'    End If
Case "BK3"
Case "BK4"
Case "BK5"
Case "BK6"
Case "BK7"
Case "BK8"
Case "BK9"
Case "DIP"
    If Ditutup Then Exit Sub
    If lbBintang.Visible Then 'discount total prosen
      frmGetAngka.Tampil AmbilAngka, "Masukkan Diskon Intern Total dalam Prosen :", False
      If AmbilAngka = -1 Then Exit Sub
    
        DiscINTERNNOTA = DefPembulatan * ((AmbilAngka * SubTotal / 100) \ DefPembulatan)
        HitungSubTotal
       ' TampilBawah
        Text1.Text = ""
        DisplayPesan "DISC NOTA #" & Space(9 - IIf(Len(Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0")) > 9, 0, Len(Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0")))) & Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0"), _
                     "TOTAL     #" & Space(9 - IIf(Len(Format(Total, "###,###,##0")) > 9, 0, Len(Format(Total, "###,###,##0")))) & Format(Total, "###,###,##0")

    Else 'dip
    frmGetAngka.Tampil AmbilAngka, "Masukkan Diskon Intern Barang dalam Prosen :", True
    If AmbilAngka = -1 Then Exit Sub
    'If Text1.Text = "" Or IsNumeric(Text1.Text) = False Then Exit Sub
      'DiscInternProsen = AmbilAngka
      If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        HargaBruto = 0
      Else
        HargaBruto = Data1.Recordset!HargaBruto
      End If
      If HargaBruto <> 0 Then
        'Jika ingin dibatasi sesuai harga minimum
        'If HargaBruto - (HargaBruto * DiscInternProsen / 100) < NullToNol(Data1.Recordset!HargaMin) Then
         ' jawab = False
         ' frmKey.Tampil jawab, 116
        'Else
        '  jawab = True
        'End If
        If jawab Then
          DiscInternProsen = AmbilAngka
          DiscInternRp = HargaBruto * DiscInternProsen / 100
          IsDiscBySupplier = True
          UpdateJualD "DIP"
        End If
        Text1.Text = ""
        Text1.SetFocus
      End If
    End If
Case "RND"
    If Ditutup Then Exit Sub
    If lbBintang.Visible Then 'discount total prosen
      frmGetAngka.Tampil AmbilAngka, "Masukkan Pembulatan Total dalam Rupiah :", True
      If AmbilAngka = -1 Then Exit Sub
        RoundingBawah = AmbilAngka
        HitungSubTotal
        Text1.Text = ""
        DisplayPesan "DISC NOTA #" & Space(9 - IIf(Len(Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0")) > 9, 0, Len(Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0")))) & Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0"), _
                     "TOTAL     #" & Space(9 - IIf(Len(Format(Total, "###,###,##0")) > 9, 0, Len(Format(Total, "###,###,##0")))) & Format(Total, "###,###,##0")
      End If 'dbp
Case "QTY"
jawab = False
  frmKey.Tampil jawab, 118
  If jawab Then
    KeyCode = 0
    DiscProsen = 0
    DiscRp = 0
    DiscInternRp = 0
    DiscInternProsen = 0
    IsDiscBySupplier = False
    cmdF3_Click
  End If
    
Case "DBP"
    cmdF12_Click
Case "DBR"
    cmdF11_Click
Case "DIR"
    If Ditutup Then Exit Sub
   
   ' If Text1.Text = "" Or IsNumeric(Text1.Text) = False Then Exit Sub
    '  DiscRp = Text1.Text
    If lbBintang.Visible Then
          frmGetAngka.Tampil AmbilAngka, "Masukkan Diskon Total(INTERN) dalam Rupiah :", True
          If AmbilAngka = -1 Then Exit Sub
        
            DiscINTERNNOTA = DiscINTERNNOTA + AmbilAngka
            HitungSubTotal
          '  TampilBawah
            Text1.Text = ""
  DisplayPesan "DISC NOTA #" & Space(9 - IIf(Len(Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0")) > 9, 0, Len(Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0")))) & Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0"), _
               "TOTAL     #" & Space(9 - IIf(Len(Format(Total, "###,###,##0")) > 9, 0, Len(Format(Total, "###,###,##0")))) & Format(Total, "###,###,##0")

  lbTotal = Format(Total, "###,###,##0")
    Else
    frmGetAngka.Tampil AmbilAngka, "Masukkan Diskon Barang (INTERN) dalam Rupiah :", True
    If AmbilAngka = -1 Then Exit Sub
      DiscInternRp = AmbilAngka
      If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        HargaBruto = 0
      Else
        HargaBruto = Data1.Recordset!HargaBruto
      End If
      If HargaBruto <> 0 Then
        DiscInternProsen = DiscInternRp * 100 / HargaBruto
        IsDiscBySupplier = True
        UpdateJualD "DIR"
        Text1.Text = ""
      End If
  End If
'Case "DBR" 'Disc barang Rupiah
'Case "DBP" 'Disc barang Prosen
Case "DTR" 'Disc Total rupiah
      If lbBintang.Visible Then
      frmGetAngka.Tampil AmbilAngka, "Masukkan Diskon Total dalam Rupiah :", False
      If AmbilAngka = -1 Then Exit Sub
    
        Disc = Disc + AmbilAngka
        HitungSubTotal
  DisplayPesan "DISC NOTA #" & Space(9 - IIf(Len(Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0")) > 9, 0, Len(Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0")))) & Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0"), _
                "TOTAL     #" & Space(9 - IIf(Len(Format(Total, "###,###,##0")) > 9, 0, Len(Format(Total, "###,###,##0")))) & Format(Total, "###,###,##0")

      '  TampilBawah
        Text1.Text = ""
      End If
Case "DTP" 'Disc Total prosen
      If lbBintang.Visible Then
      frmGetAngka.Tampil AmbilAngka, "Masukkan Diskon Total dalam Prosen :", False
      If AmbilAngka = -1 Then Exit Sub
    
        Disc = DefPembulatan * ((AmbilAngka * SubTotal / 100) \ DefPembulatan)
        HitungSubTotal
  DisplayPesan "DISC NOTA #" & Space(9 - IIf(Len(Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0")) > 9, 0, Len(Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0")))) & Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0"), _
                "TOTAL     #" & Space(9 - IIf(Len(Format(Total, "###,###,##0")) > 9, 0, Len(Format(Total, "###,###,##0")))) & Format(Total, "###,###,##0")

        'TampilBawah
        Text1.Text = ""
      End If
Case "VCR" 'Voucher
'      If IDMember < 1 Then
'        frmPesan.lbPesan = "Member harus dimasukkan!!!"
'        frmPesan.Show 1
'      Else
      If lbBintang.Visible Then
'        frmvoucher.Tampil NilaiVoucher, IDSales
        AmbilVoucher = 0
        jawaban = False
        frmPenerbitVoucher.Tampil jawaban, IDPenerbitVoucher, KodePenerbitVoucher, AmbilVoucher, qtyVcr
        If jawaban = True And (qtyVcr * AmbilVoucher) <> 0 Then
          If Voucher + (qtyVcr * AmbilVoucher) > Total Then
            frmPesan.lbPesan = "VOUCHER MELEBIHI TOTAL!"
            frmPesan.Show 1
          Else
                idSalesDVoucher = GetNewID("MSalesDVoucher")
                Data1.Database.Execute "Insert Into MSalesDVoucher(NoID,IDSales,IDPenerbit,NamaPenerbit,Nominal,Qty) values(" & _
                    idSalesDVoucher & "," & IDSales & "," & IDPenerbitVoucher & ",'" & KodePenerbitVoucher & "'," & FixKoma(AmbilVoucher) & "," & qtyVcr & ")"
                NilaiVoucher = qtyVcr * AmbilVoucher '(Text1.Text)
                Voucher = Voucher + NilaiVoucher
                Dibayar = Dibayar + NilaiVoucher
                DisplayPesan "VCR# " & Space(15 - Min(Len(Format(Voucher, "###,###,##0")), 15)) & Format(Voucher, "###,###,##0"), "PAY# " & Space(15 - Min(Len(Format(Dibayar, "###,###,##0")), 15)) & Format(Dibayar, "###,###,##0")
                HitungSubTotal
                TampilBawah
                Text1.Text = ""
          End If
        End If
      End If
'If IsNumeric(Text1.Text) Then
'  If Voucher + CCur(Text1.Text) > Total Then
'    frmPesan.lbPesan = "VOUCHER MELEBIHI TOTAL!"
'    frmPesan.Show 1
'  Else
'        NilaiVoucher = CCur(Text1.Text)
'        Voucher = Voucher + NilaiVoucher
'        Dibayar = Dibayar + NilaiVoucher
'        DisplayPesan "VCR# " & Space(15 - Min(Len(Format(Voucher, "###,###,##0")), 15)) & Format(Voucher, "###,###,##0"), "PAY# " & Space(15 - Max(Len(Format(Dibayar, "###,###,##0")), 15)) & Format(Dibayar, "###,###,##0")
'        HitungSubTotal
'        TampilBawah
'        Text1.Text = ""
'  End If
'End If
'      End If
'Case "DTP" 'Disc Total Prosen
Case "RPT" ' Re Print
If Ditutup Then
  If CountReprint < 1 Then
    CountReprint = CountReprint + 1
    Reprint CountReprint
  Else
    jawab = False
    frmKey.Tampil jawab, 116
    If jawab Then
          CountReprint = CountReprint + 1
          Reprint CountReprint
    End If
  End If
Else
   frmPesan.lbPesan = "Reprint Setelah Transaksi Selesai!!!!"
        frmPesan.Show 1
End If
Case "RST"
    If Not IsResetPerKasir Then
      If Not isSupervisor Then
          frmPesan.lbPesan = "Kasir Tidak Boleh!!!!"
          frmPesan.Show 1
          Exit Sub
      End If
    End If
  jawab = False
  frmKey.Tampil jawab, 118
  If jawab Then
    KeyCode = 0
    'If Left(KodeUserLogin, 2) = "ST" Then
     '   frmReset.Tampil jawab, 118
    'Else
        frmReset1Saja.Tampil jawab, 118
    'End If
  End If
Case "CSO"
Case "NS" 'No Sales
    If NoPortDrawer = -1 Then
        frmDaftarJual.Show 1
    Else
        openDrawer
    End If
Case "AMT"
    If lbBintang.Visible Or Ditutup Then Exit Sub
    If Data1.Recordset.EOF And Data1.Recordset.BOF Then
        Text1.Text = ""
    Else
        Data1.Recordset.MoveLast
        Text1.Text = IIf(IsNull(Data1.Recordset!KodeInv), "", Data1.Recordset!KodeInv)
    End If
    If lbBintang.Visible Then Exit Sub
    If Text1.Text = "" Then Exit Sub
    rs.Index = "Barcode"
    rs.Seek "=", Text1.Text
    If rs.NoMatch Then
      rs.Index = "Kode"
      rs.Seek "=", Text1.Text
      If rs.NoMatch Then
        rs.Index = "Nama"
        rs.Seek "=", Text1.Text
        If rs.NoMatch Then
        Else
            DiscProsen = 0
            DiscRp = 0
            DiscInternRp = 0
            DiscInternProsen = 0
            IsDiscBySupplier = False

          TambahJualD
          Text1.Text = ""
        End If
      Else
        DiscProsen = 0
        DiscRp = 0
        DiscInternRp = 0
        DiscInternProsen = 0
        IsDiscBySupplier = False
      
      TambahJualD
      Text1.Text = ""
      End If
    Else
        DiscProsen = 0
        DiscRp = 0
        DiscInternRp = 0
        DiscInternProsen = 0
        IsDiscBySupplier = False

      TambahJualD
      Text1.Text = ""
    End If
Case "SPC"
'Case "LFT"
'Case "RGT"
'Case "PUP"
'Case "PDN"
End Select
End Sub

Sub TampilBawah()
  PotonganPembulatan = RoundingBawah + ((Round(SubTotal) - Round(Disc) - DiscINTERNNOTA - JumDiscInternRp) Mod DefPembulatan) '- TotalDiscBrg
  'Total = (Round(SubTotal) - Round(Disc)) - PotonganPembulatan - DiscINTERNNOTA '- TotalDiscBrg
  Total = (Round(SubTotal) - Round(Disc)) - PotonganPembulatan - DiscINTERNNOTA
  
  Kembali = Round(Dibayar) - Total
  KomisiRp = (((Total * KomisiProsen / 100) \ 50) * 50)
  lbDiscBrg = Format(TotalDiscBrg + JumDiscInternRp, "###,###,##0")
  lbSubTtl = Format(SubTotal, "###,###,##0")
  lbRounding = Format(PotonganPembulatan, "###,###,##0")
  LbDisc = Format(Disc + DiscINTERNNOTA, "###,###,##0")
  lblNilaiReedem = Format(ReedemNilai, "###,###,##0")
  lblReedemPoin = Format(ReedemPoin, "###,###,##0")
  lbTotal = Format(Total, "###,###,##0")
  lbDibayar = Format(Dibayar + BiayaCC - Voucher - ReedemNilai, "###,###,##0")
 ' If Kembali >= 0 Then
    lbKembali = Format(Kembali, "###,###,##0")
  '  SaldoHutang = 0
  '  lbHutang = Format(SaldoHutang, "###,###,##0")
'Else
 '   lbKembali = Format(0, "###,###,##0")
  '  SaldoHutang = Abs(Kembali)
  '  lbHutang = Format(SaldoHutang, "###,###,##0")
  '  Kembali = 0
'End If
  lbVoucher = Format(Voucher, "###,###,##0")
  lbPoin = Format(BelanjaPoin, "###,###,##0")
  lbSopir.Caption = "Grup :" & Agen
  lbKomisiProsen.Caption = "Komisi (%):" & Format(KomisiProsen, "###,##0.00")
  lbKomisiRp.Caption = "Komisi Rp.:" & Format(KomisiRp, "###,###,##0")
  UpdateSales
'  DiscProsenBawah = 0
'  DiscRupiahBawah = 0
End Sub
Function GetQtyCur(ByVal Barcode As String) As Long
Dim rsTot As Recordset
Dim Qty As Double
Set rsTot = dbSale.OpenRecordset("SELECT SUM(MSalesD.Qty) as Total From MSalesD where BARCODE='" & Replace(Barcode, "'", "''") & "' AND (TRANSAKSI='PLU' OR TRANSAKSI='VOD') AND IDSales=" & IDSales)
If rsTot.BOF And rsTot.EOF Then
  Qty = 0
Else
  If IsNull(rsTot!Total) Then
    Qty = 0
  Else
    Qty = rsTot!Total
  End If
End If
rsTot.Close
GetQtyCur = Qty
End Function
Private Function JanganDiGroup() As Boolean
  On Error GoTo Trace
  JanganDiGroup = NullToBool(rs!IsGroupQty)
  Exit Function
Trace:
  JanganDiGroup = False
End Function

'Private Function SisaStockGudang() As Double
'On Error GoTo Trace
'  Dim x As Double
'  x = 999999999999#
'  If isRemcomendedOnline Then
'    x = NullToNol(ExecuteSkalarSQL("SELECT SUM((MkartuStok.QtyMasuk-MkartuStok.QtyKeluar)*MkartuStok.Konversi) FROM MKartuStok WHERE IDBarang=" & NullToNol(rs!IdInventor) & " AND IDGudang=" & IDGudangDef))
'  End If
'  SisaStockGudang = x
'  Exit Function
'Trace:
'  If Err.Number <> 0 Then
'    BuatLogApp Err.Number & " " & Err.Description
'    Err.Clear
'  End If
'End Function
Sub TambahJualD()
Dim SelBks
'LOGIKA:
'CARI BARCODE YG SAMA JADIKAN 1 BARIS (QTY DITAMBAH)
'JIKA  IS FAMILYGROUP MAKA CARI KODE SAMA, CEK APAKAH QTY MENCUKUPI, JIKA YA UPDATE HARGA
'    If QTY <= 0 Then frmPesan.lbPesan = "MASUKKAN QTY YG BENAR !!!!": frmPesan.Show 1: Exit Sub
'CEK FAMILY GROUP DULU: sudah ada item dg kode sama
Dim IsBuah As Boolean
If Left(rs!kode, 2) = "25" Or UCase(NullToStr(rs!kode)) = UCase("02712543") Or JanganDiGroup Then
  IsBuah = True
Else
  IsBuah = False
End If
Qty = Qty + IsTambahItem(rs!Barcode, Qty, IsBuah)
    'If KodeMember = "" Then frmPesan.lbPesan = "MASUKKAN CUSTOMER !!!!": frmPesan.Show 1: Exit Sub
    TipeHargaJual = 1 'KHUSUS SPBU99

    'If QTY * NullToNol(rs!Konversi) >= SisaStockGudang Then frmPesan.lbPesan = "QTY MELEBIHI SISA STOK !!!!": frmPesan.Show 1: Exit Sub
    Dim HargaJual As Double
    bolehbergerak = False
    Dim DefHargaJual As Double
    'jika max item disc dibatasi
'    If rs!QtyMaxDisc <> 0 Then
'      If (GetQtyCur(rs!BARCODE) + QTY) <= rs!QtyMaxDisc Then
'      Else
'            frmPesan.lbPesan = "BARANG DISCOUNT DIBATASI!" & vbCrLf & "GUNAKAN 'PE' UNTUK HARGA NORMAL!"
'            frmPesan.Show 1
'            Exit Sub
'        Exit Sub
'      End If
'    End If

    IDSalesD = GetNewID("MSALESD")
    
If (NullToBool(rs!IsFamilyGroup) And IsNull(rs!TanggalDariFamily)) Or (NullToBool(rs!IsFamilyGroup) And Date >= NullToDate(rs!TanggalDariFamily) And Date <= NullToDate(rs!TanggalSampaiFamily)) Then
  If JumlahQtyNotaIniSelainBarcode(rs!kode, rs!Barcode) + Qty >= NullToNol(rs!Qtyfamily) Then
      DefHargaJual = NullToNol(rs!HargaFamily)
  Else
    DefHargaJual = NullToNol(rs!HargaJual)
  End If
    If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
     Data1.Recordset.MoveFirst
     Do While Not Data1.Recordset.EOF
        If Data1.Recordset!KodeInv = rs!kode Then
           Data1.Recordset.Edit
           Data1.Recordset!HargaBruto = DefHargaJual
           Data1.Recordset!harga = DefHargaJual
           Data1.Recordset!Jumlah = Bulatkan(Data1.Recordset!Qty * DefHargaJual, 0)
           Data1.Recordset.Update
           Data1.Recordset.Bookmark = Data1.Recordset.LastModified
          End If
        Data1.Recordset.MoveNext
      Loop
    End If
Else
      If NullToBool(rs!IsKelipatan) Then
        DefHargaJual = Bulatkan(((rs!HargaKelipatan * (Qty \ rs!QtyKelipatan) * rs!QtyKelipatan) + rs!HargaA * (Qty - (Qty \ rs!QtyKelipatan) * rs!QtyKelipatan)) / Qty, 0)
      Else 'If NullToBool(rs!IsGrosir) Then'Multi Harga
          If Qty >= NullToNol(rs!Qty3) And NullToNol(rs!Qty3) <> 0 Then
              DefHargaJual = NullToNol(rs!Harga3)
          ElseIf Qty >= NullToNol(rs!Qty2) And NullToNol(rs!Qty2) <> 0 Then
              DefHargaJual = NullToNol(rs!Harga2)
          ElseIf Qty >= NullToNol(rs!Qty1) And NullToNol(rs!Qty1) <> 0 Then
              DefHargaJual = NullToNol(rs!Harga1)
          Else
              DefHargaJual = NullToNol(rs!HargaJual)
          End If
'
'Else
'          DefHargaJual = NullToNol(rs!HargaJual)
      End If
End If
    If DefHargaJual = 0 Then
      frmsetHarga.Tampil HargaJual, rs!kode & " " & rs!Nama
      If HargaJual = -1 Then Exit Sub
    Else
      HargaJual = NullToNol(DefHargaJual)
    End If
    If IsLockDiHargaBeli Then
      If HargaJual < NullToNol(rs!HargaBeliTerakhir) Then
        Dim jwban As Boolean
        frmKey.Tampil jwban, 13
        If Not jwban Then Exit Sub
      End If
    End If
    
    Data1.Recordset.AddNew
    Data1.Recordset!NoID = IDSalesD
    Data1.Recordset!IDSales = IDSales
    Data1.Recordset!IDInvSat = rs!NoID
    Data1.Recordset!IdInventor = rs!IdInventor
    Data1.Recordset!Qty = Qty
    HargaBruto = HargaJual
    Data1.Recordset!HargaBruto = HargaBruto
    Data1.Recordset!DiscRp = 0
    Data1.Recordset!DiscProsen = 0
    Data1.Recordset!DiscInternRp = 0
    Data1.Recordset!DiscInternProsen = 0
    Data1.Recordset!harga = HargaJual
    Data1.Recordset!KodeInv = rs!kode
    Data1.Recordset!NamaInv = rs!Nama
    Data1.Recordset!Satuan = rs!KodeSat
    Data1.Recordset!Barcode = IIf(NullToStr(rs!Barcode) = "", "-", rs!Barcode)
    Data1.Recordset!idSatuan = rs!idSatuan
    Data1.Recordset!Konversi = rs!Konversi
    Data1.Recordset!HargaPokok = rs!HargaPokok
    Data1.Recordset!IsPoin = rs!IsPoin
    Data1.Recordset!IsPDP = False
    Data1.Recordset!IsPoinSupplier = rs!IsPoinSupplier
    Data1.Recordset!IDPoinSupplier = rs!IDPoinSupplier
    Data1.Recordset!Jumlah = Bulatkan(Qty * HargaJual, 0)
    Data1.Recordset!BKP = rs!BKP
    Data1.Recordset!Transaksi = "PLU"

    Data1.Recordset.Update
    Data1.Recordset.Bookmark = Data1.Recordset.LastModified
    If Data1.Recordset!IsPoin Then
     cetakdetil Data1.Recordset!KodeInv, Data1.Recordset!NamaInv, Format(Data1.Recordset!Qty, "##0"), Format(Data1.Recordset!harga, "###,###,##0"), Format(Qty * Data1.Recordset!harga, "###,###,##0")
    Else
     cetakdetil Data1.Recordset!KodeInv, Data1.Recordset!NamaInv & "*)", Format(Data1.Recordset!Qty, "##0"), Format(Data1.Recordset!harga, "###,###,##0"), Format(Qty * Data1.Recordset!harga, "###,###,##0")
    End If
    DisplayPesan Data1.Recordset!NamaInv, Format(Qty, "##0") & " X    " & Format(Data1.Recordset!harga, "###,###,##0")
    Qty = 1
    lbQty = Format(Qty, "##0") & " X"
    HitungSubTotal
    If NullToNol(rs!DiscMemberRp2) > 0 And IDMember >= 1 Then 'Kalau Ada Member
'        If Not IsNull(rs!TglDariDiskon2) And Not IsNull(rs!TglSampaiDiskon2) And Date >= NullToDate(rs!TglDariDiskon2) And Date <= NullToDate(rs!TglSampaiDiskon2) And IsQtyPDPAda(NullToNol(rs!IdInventor), Qty, NullToDate(rs!TglDariDiskon2), NullToDate(rs!TglSampaiDiskon2), CDbl(SubTotal)) Then
'          DiscInternProsen = NullToNol(rs!DiscMemberProsen2)
'          If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
'              HargaBruto = 0
'          Else
'              HargaBruto = Data1.Recordset!HargaBruto
'          End If
'          If HargaBruto <> 0 Then
'            DiscInternRp = NullToNol(rs!DiscMemberRp2) 'NilaiDiskon ' ((HargaBruto * DiscInternProsen / 100) \ 1) * 1  'tanpa koma
'            DiscInternProsen = NullToNol(rs!DiscMemberProsen2)
'            UpdateJualD
'            Text1.Text = ""
'          End If
'        Else 'Kemungkinan ke Promo Diskon
'          If NullToNol(rs!DiscRupiah) > 0 And Not IsNull(rs!DiscExpired) And Not IsNull(rs!DiscMulai) Then
'            If Date >= rs!DiscMulai And Date <= rs!DiscExpired Then
'                DiscInternProsen = NullToNol(rs!DiscProsen)
'                If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
'                    HargaBruto = 0
'                Else
'                    HargaBruto = Data1.Recordset!HargaBruto
'                End If
'                If HargaBruto <> 0 Then
'                  DiscInternRp = NullToNol(rs!DiscRupiah) 'NilaiDiskon ' ((HargaBruto * DiscInternProsen / 100) \ 1) * 1  'tanpa koma
'                  DiscInternProsen = NullToNol(rs!DiscProsen)
'                  UpdateJualD
'                  Text1.Text = ""
'                End If
'            End If
'          End If
'        End If
    ElseIf NullToNol(rs!DiscRupiah) > 0 Then
        If Not IsNull(rs!DiscExpired) And Not IsNull(rs!DiscMulai) Then
            If Date >= rs!DiscMulai And Date <= rs!DiscExpired Then
                DiscInternProsen = NullToNol(rs!DiscProsen)
                If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
                    HargaBruto = 0
                Else
                    HargaBruto = Data1.Recordset!HargaBruto
                End If
                If HargaBruto <> 0 Then
                  DiscInternRp = NullToNol(rs!DiscRupiah) 'NilaiDiskon ' ((HargaBruto * DiscInternProsen / 100) \ 1) * 1  'tanpa koma
                  DiscInternProsen = NullToNol(rs!DiscProsen)
                  UpdateJualD
                  Text1.Text = ""
                End If
            End If
        End If
    End If
    If isRemcomendedOnline And isTampilSaldoStock Then
      SaldoStock = NullToNol(ExecuteSkalarSQL("Select Sum(ISNULL(Konversi,1)*(ISNULL(QtyMasuk,0)-ISNULL(QtyKeluar,0))) as Saldo from MKartuStok where MKartuStok.IDBarang=" & NullToNol(rs!IdInventor)))
      Data1.Recordset.Edit
      Data1.Recordset!SaldoStock = SaldoStock
      Data1.Recordset.Update
      Data1.Recordset.Bookmark = Data1.Recordset.LastModified
    End If
    Set SelBks = TDBGrid1.SelBookmarks
    While SelBks.Count <> 0
        SelBks.Remove 0
    Wend
    TDBGrid1.SelBookmarks.Add TDBGrid1.Bookmark
    If IDMember >= 1 Then
      UpdatekanDiskonPerBarang
    End If
End Sub
Function JumlahQtyNotaIniSelainBarcode(ByVal kode As String, ByVal Barcode As String) As Double
  Dim qtyTot As Double
  qtyTot = 0
    If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
     Data1.Recordset.MoveFirst
     Do While Not Data1.Recordset.EOF
            If Data1.Recordset!KodeInv = kode And Data1.Recordset!Barcode <> Barcode Then
                qtyTot = qtyTot + Data1.Recordset!Qty
            End If
          Data1.Recordset.MoveNext
      Loop
  End If
  JumlahQtyNotaIniSelainBarcode = qtyTot
End Function
Function IsOpenDepartemen(ByVal kode As String) As Boolean
  On Error GoTo Trace
  Dim x As Boolean
  x = False
  Dim DB As Database
  Dim rst As Recordset
  Set DB = OpenDatabase(DirDatabase & "\dbMaster.mdb")
  Set rst = DB.OpenRecordset("SELECT HargaJual FROM TINV WHERE UCASE(Barcode)='" & Replace(kode, "'", "''") & "'")
  If Not (rst.EOF And rst.BOF) Then
    If NullToNol(rst!HargaJual) = 0 Then
      x = True
    Else
      x = False
    End If
  End If
Trace:
  If Err.Number <> 0 Then
    MsgBox Err.Description
    x = False
    Err.Clear
  End If
  
  rst.Close
  Set rst = Nothing
  
  DB.Close
  Set DB = Nothing
  
  IsOpenDepartemen = x
End Function
Function Bulatkan50(ByVal x As Long) As Long
     Dim sisa As Long
     sisa = x Mod 50
     If sisa > 25 Then
      Bulatkan50 = (x \ 50) * 50 + 50
     Else
      Bulatkan50 = (x \ 50) * 50
      End If
End Function
Public Function Bulatkan(ByVal x As Double, ByVal Koma As Integer) As Double
        If Koma >= 0 Then
            Bulatkan = Round(x, CInt(Koma))
            If Round(x - Bulatkan, CInt(Koma + 5)) >= 0.5 / (10 ^ Koma) Then Bulatkan = Bulatkan + 1 / (10 ^ Koma)
        Else
            Bulatkan = x
        End If
    End Function
Sub TambahJualDKhususPDP()
    If Qty <= 0 Then frmPesan.lbPesan = "MASUKKAN QTY YG BENAR !!!!": frmPesan.Show 1: Exit Sub
'    If KodeMember = "" Then frmPesan.lbPesan = "MASUKKAN CUSTOMER !!!!": frmPesan.Show 1: Exit Sub
    'If QTY * NullToNol(rs!Konversi) >= SisaStockGudang Then frmPesan.lbPesan = "QTY MELEBIHI SISA STOK !!!!": frmPesan.Show 1: Exit Sub
    Dim HargaJual As Double
    Dim dbHistori As Database
    Dim RsHistori As Recordset
    bolehbergerak = False
    IDSalesD = GetNewID("MSALESD")
    Data1.Recordset.AddNew
    Data1.Recordset!NoID = IDSalesD
    Data1.Recordset!IDSales = IDSales
    Data1.Recordset!IDInvSat = rs!NoID
    Data1.Recordset!IdInventor = rs!IdInventor
    Data1.Recordset!Qty = Qty
    Data1.Recordset!harga = HargaJualKhusus 'HargaJual
    HargaBruto = HargaJualKhusus
    Data1.Recordset!HargaBruto = HargaBruto
    Data1.Recordset!DiscRp = 0
    Data1.Recordset!DiscProsen = 0
    Data1.Recordset!DiscInternRp = 0
    Data1.Recordset!DiscInternProsen = 0
    Data1.Recordset!KodeInv = rs!kode
    Data1.Recordset!NamaInv = rs!Nama
    Data1.Recordset!Satuan = rs!KodeSat
    Data1.Recordset!Barcode = rs!Barcode
    Data1.Recordset!idSatuan = rs!idSatuan
    Data1.Recordset!Konversi = rs!Konversi
    Data1.Recordset!HargaPokok = rs!HargaPokok
    Data1.Recordset!DiscRp = 0
    Data1.Recordset!DiscProsen = 0
    Data1.Recordset!DiscInternRp = 0
    Data1.Recordset!DiscInternProsen = 0
    Data1.Recordset!harga = HargaJualKhusus
    Data1.Recordset!KodeInv = rs!kode
    Data1.Recordset!NamaInv = rs!Nama
    Data1.Recordset!Satuan = rs!KodeSat
    Data1.Recordset!Barcode = IIf(NullToStr(rs!Barcode) = "", "-", rs!Barcode)
    Data1.Recordset!idSatuan = rs!idSatuan
    Data1.Recordset!Konversi = rs!Konversi
    Data1.Recordset!HargaPokok = rs!HargaPokok
    Data1.Recordset!IsPoin = rs!IsPoin
    Data1.Recordset!IsPoinSupplier = rs!IsPoinSupplier
    Data1.Recordset!IDPoinSupplier = rs!IDPoinSupplier
    Data1.Recordset!Jumlah = Bulatkan(Qty * HargaJualKhusus, 0)
    Data1.Recordset!BKP = rs!BKP
    Data1.Recordset!IsPDP = True
'     If TipeHargaJual = 1 Then
'      Data1.Recordset!DiscProsen1 = NullToNol(rs!DiscProsen1A)
'      Data1.Recordset!DiscProsen2 = NullToNol(rs!DiscProsen2A)
'     ElseIf TipeHargaJual = 2 Then
'      Data1.Recordset!DiscProsen1 = NullToNol(rs!DiscProsen1B)
'      Data1.Recordset!DiscProsen2 = NullToNol(rs!DiscProsen2B)
'     ElseIf TipeHargaJual = 3 Then
'      Data1.Recordset!DiscProsen1 = NullToNol(rs!DiscProsen1C)
'      Data1.Recordset!DiscProsen2 = NullToNol(rs!DiscProsen2C)
'     ElseIf TipeHargaJual = 4 Then
'      Data1.Recordset!DiscProsen1 = NullToNol(rs!DiscProsen1D)
'      Data1.Recordset!DiscProsen2 = NullToNol(rs!DiscProsen2D)
'     ElseIf TipeHargaJual = 5 Then
'      Data1.Recordset!DiscProsen1 = NullToNol(rs!DiscProsen1E)
'      Data1.Recordset!DiscProsen2 = NullToNol(rs!DiscProsen2E)
'     ElseIf TipeHargaJual = 6 Then
'      Data1.Recordset!DiscProsen1 = NullToNol(rs!DiscProsen1F)
'      Data1.Recordset!DiscProsen2 = NullToNol(rs!DiscProsen2F)
'     Else
'      Data1.Recordset!DiscProsen1 = NullToNol(rs!DiscProsen1B)
'      Data1.Recordset!DiscProsen2 = NullToNol(rs!DiscProsen2B)
'     End If
'     Data1.Recordset!DiscProsen3 = 0
'     Data1.Recordset!Disc1 = 0
'     Data1.Recordset!Disc2 = 0
'     Data1.Recordset!Disc3 = 0
    Data1.Recordset!DiscRp = 0
    Data1.Recordset!DiscProsen = 0
    Data1.Recordset!DiscInternRp = 0
    Data1.Recordset!DiscInternProsen = 0
'     Data1.Recordset!HargaD = Bulatkan((Data1.Recordset!HargaBruto * (1 - (NullToNol(Data1.Recordset!DiscProsen1) / 100)) * (1 - (NullToNol(Data1.Recordset!DiscProsen2) / 100)) * (1 - (NullToNol(Data1.Recordset!DiscProsen3) / 100))), 0)
     Data1.Recordset!Jumlah = Bulatkan(Qty * HargaJualKhusus, 0) '(Bulatkan((HargaJual * (1 - (NullToNol(Data1.Recordset!DiscProsen1) / 100)) * (1 - (NullToNol(Data1.Recordset!DiscProsen2) / 100)) * (1 - (NullToNol(Data1.Recordset!DiscProsen3) / 100))), 0) - 0 - 0 - 0)

    Data1.Recordset!Transaksi = "PE"
    Data1.Recordset.Update
    Data1.Recordset.Bookmark = Data1.Recordset.LastModified
    If Data1.Recordset!IsMember Then
      cetakdetil Data1.Recordset!KodeInv, Data1.Recordset!NamaInv, Format(Data1.Recordset!Qty, "##0"), Format(Data1.Recordset!harga, "###,###,##0"), Format(Qty * Data1.Recordset!harga, "###,###,##0")
    Else
      cetakdetil Data1.Recordset!KodeInv, Data1.Recordset!NamaInv & "*)", Format(Data1.Recordset!Qty, "##0"), Format(Data1.Recordset!harga, "###,###,##0"), Format(Qty * Data1.Recordset!harga, "###,###,##0")
    End If
    DisplayPesan Data1.Recordset!NamaInv, Format(Qty, "##0") & " X    " & Format(Data1.Recordset!harga, "###,###,##0")
    Qty = 1
    lbQty = Format(Qty, "##0") & " X"
    Set dbHistori = OpenDatabase(DirDatabase & "\Histori.mdb")
    Set RsHistori = dbHistori.OpenRecordset("Historikasir")
    RsHistori.AddNew
    'ID
    RsHistori!kassa = NamaMesin
    RsHistori!IDUser = IDUser
    RsHistori!KodeUser = KodeKasir
    RsHistori!Tanggal = Date
    RsHistori!Jam = Time
    RsHistori!IDSales = IDSales
    RsHistori!IDSalesD = IDSalesD
    RsHistori!Transaksi = "PE"
    RsHistori!IdInventor = Data1.Recordset!IdInventor
    RsHistori!idSatuan = Data1.Recordset!idSatuan
    RsHistori!KodeInventor = Data1.Recordset!KodeInv
    RsHistori!NamaInventor = Data1.Recordset!NamaInv
    RsHistori!HargaPokok = Data1.Recordset!HargaPokok
    RsHistori!HargaJualMaster = rs!HargaJual
    RsHistori!HargaJualKhusus = HargaJualKhusus
    RsHistori.Update
    RsHistori.Close
    dbHistori.Close
    HitungSubTotal
End Sub
Sub Reprint(Optional ByVal ReprintKe As Integer = 1)
Dim i As Integer
Dim j As Integer
Dim z As Double
    'Prin Judulstruk
    If Not IsHematKertas Then
            Prin Chr(13) & Chr(10)
            PrinBigChar Space(15 - Min(15, Len(NamaToko))) & UCase(NamaToko)
            DoEvents
            Prin Judulstruk & Chr(13) & Chr(10) & "---------------------------------------"
            DoEvents
        End If
    PrinBigChar "----REPRINT----"
    Prin "---------------------------------------"
    If Data1.Recordset.EOF And Data1.Recordset.BOF Then Exit Sub
    Data1.Recordset.MoveFirst
    
    Do While Not Data1.Recordset.EOF
        DoEvents
        If Data1.Recordset!IsMember Then
          cetakdetil Data1.Recordset!KodeInv, Data1.Recordset!NamaInv, Format(Data1.Recordset!Qty, "##0"), Format(Data1.Recordset!HargaBruto, "###,###,##0"), Format(Data1.Recordset!Qty * Data1.Recordset!HargaBruto, "###,###,##0")
        Else
           cetakdetil Data1.Recordset!KodeInv, Data1.Recordset!NamaInv & "*)", Format(Data1.Recordset!Qty, "##0"), Format(Data1.Recordset!HargaBruto, "###,###,##0"), Format(Data1.Recordset!Qty * Data1.Recordset!HargaBruto, "###,###,##0")
        End If
      If Data1.Recordset!DiscRp + Data1.Recordset!DiscInternRp > 0 Then
      DoEvents
      For i = 1 To 32000
        For j = 1 To 160
        z = 100 * 98
      'perlambat
        Next
      Next
         Prin "*** Disc. " & Format(Data1.Recordset!DiscProsen + Data1.Recordset!DiscInternProsen, "#0.00") & _
         "% " & Space(39 - Len("*** Disc. " & Format((Data1.Recordset!DiscProsen + Data1.Recordset!DiscInternProsen), "#0.00") & _
         "% " & "-" & Format((Data1.Recordset!DiscInternRp + Data1.Recordset!DiscRp) * Data1.Recordset!Qty, "###,##0"))) & "-" & Format((Data1.Recordset!DiscInternRp + Data1.Recordset!DiscRp) * Data1.Recordset!Qty, "###,##0")
      End If
      
      DoEvents
      For i = 1 To 32000
      For j = 1 To 60
'        z = 100 * 98
      'perlambat
      Next
      Next
      Data1.Recordset.MoveNext
    Loop
'    DoEvents
'    Prin Chr(13) & Chr(10) & "---------------------------------------" & Chr(13) & Chr(10)
    DoEvents
    Prin "---------------------------------------" & Chr(13) & Chr(10) & "Subtotal" & Space(31 - Len(Format(SubTotal, "###,###,##0"))) & Format(SubTotal, "###,###,##0") & Chr(13) & Chr(10) & _
               IIf(TotalDiscBrg + JumDiscInternRp = 0, "", "Discount Barang" & Space(24 - Len(Format(TotalDiscBrg + JumDiscInternRp, "###,###,##0"))) & Format(TotalDiscBrg + JumDiscInternRp, "###,###,##0") & Chr(13) & Chr(10)) & _
               IIf(Disc + DiscINTERNNOTA = 0, "", "Discount Nota" & Space(26 - Len(Format(Disc + DiscINTERNNOTA, "###,###,##0"))) & Format(Disc + DiscINTERNNOTA, "###,###,##0") & Chr(13) & Chr(10)) & _
              "Voucher " & Space(31 - Len(Format(Voucher, "###,###,##0"))) & Format(Voucher, "###,###,##0") & Chr(13) & Chr(10) & _
              IIf(BiayaCC > 0, "Biaya CC" & Space(31 - Len(Format(BiayaCC, "###,###,##0"))) & Format(BiayaCC, "###,###,##0") & Chr(13) & Chr(10), "") & _
              "Dibayar " & Space(31 - Len(Format(Dibayar + BiayaCC - Voucher - ReedemNilai, "###,###,##0"))) & Format(Dibayar + BiayaCC - Voucher - ReedemNilai, "###,###,##0") & Chr(13) & Chr(10) & _
              IIf(SaldoHutang <= 0, "", "Hutang  " & Space(31 - Len(Format(SaldoHutang, "###,###,##0"))) & Format(SaldoHutang, "###,###,##0") & Chr(13) & Chr(10)) & _
              "Kembali " & Space(31 - Len(Format(Kembali, "###,###,##0"))) & Format(Kembali, "###,###,##0") & Chr(13) & Chr(10) & _
              IIf(ISSembunyikanFooterbarangKSB, "", "Belanja Barang POIN" & Space(20 - Len(Format(BelanjaPoin, "###,###,##0"))) & Format(BelanjaPoin, "###,###,##0") & Chr(13) & Chr(10)) & _
              "#" & Format(CLng(NoNota), "0#####") & "," & "Kasir" & NamaMesin & ":" & Mid(NamaKasir, 1, 6) & Space(6 - IIf(Len(NamaKasir) <= 6, Len(NamaKasir), 6)) & Format(Date, " dd/MM/yy") & Format(Time, "hh:mm:ss") & Chr(13) & Chr(10) & _
              "Macam Barang =" & Format(MacamItem, "##0") & " , Total Item =" & Format(Jumlahitem, "##0") & Chr(13) & Chr(10) & _
              IIf(KodeMember = "", "", "customer: " & KodeMember & "-" & NamaMember & Chr(13) & Chr(10)) & _
              "----TERIMA KASIH ATAS KUNJUNGAN ANDA----" & Chr(13) & Chr(10) & _
              IIf(Footer1 = "", "", Footer1 & Chr(13) & Chr(10)) & _
              IIf(Footer2 = "", "", Footer2 & Chr(13) & Chr(10)) & _
              IIf(Footer3 = "", "", Footer3 & Chr(13) & Chr(10)) & _
              Chr(13) & Chr(10)
        DoEvents
        If IsHematKertas Then
            Prin Chr(13) & Chr(10)
            PrinBigChar Space(15 - Min(15, Len(NamaToko))) & UCase(NamaToko)
            DoEvents
            Prin Judulstruk & Chr(13) & Chr(10) & "---------------------------------------"
            DoEvents
        End If
    papercut
    CetakStruck IDSales, True, ReprintKe
    If isRemcomendedOnline Then
      CetakStruckReedem2 IDSales, True
    End If
    DoEvents
End Sub
Sub Reprintold()
Dim i As Integer
Dim j As Integer
Dim z As Long
Dim Timer1 As Long
If Dir(DirDatabase & "\Setting.mdb") = "" Then
    Timer1 = 500
  Else
  Dim dbs1 As Database
  Dim rs1 As Recordset
      Set dbs1 = OpenDatabase(DirDatabase & "\Setting.mdb")
      Set rs1 = dbs1.OpenRecordset("MSetting")
      If rs1.EOF And rs1.BOF Then
        Timer1 = 500
      Else
        Footer1 = NullToNol(rs1!TimerReprint)
      End If
      rs1.Close
    dbs1.Close
    
    Set dbs1 = Nothing
 End If

    'Prin Judulstruk
    PrinBigChar "----REPRINT----"
    Prin "---------------------------------------"
    If Data1.Recordset.EOF And Data1.Recordset.BOF Then Exit Sub
    Data1.Recordset.MoveFirst
    
    Do While Not Data1.Recordset.EOF
      cetakdetil Data1.Recordset!KodeInv, Data1.Recordset!NamaInv, Format(Data1.Recordset!Qty, "##0"), Format(Data1.Recordset!harga, "###,###,##0"), Format(Data1.Recordset!Qty * Data1.Recordset!harga, "###,###,##0")
      DoEvents
      For i = 1 To 32000
      For j = 1 To Timer1 '150 '60
        z = 100 * 98
      'perlambat
      Next
      Next
      Data1.Recordset.MoveNext
    Loop
    DoEvents
    Prin "---------------------------------------" & Chr(13) & Chr(10) & _
    "Subtotal" & Space(31 - Len(Format(SubTotal, "###,###,##0"))) & Format(SubTotal, "###,###,##0") & Chr(13) & Chr(10) & _
                "Voucher " & Space(31 - Len(Format(Voucher, "###,###,##0"))) & Format(Voucher, "###,###,##0") & Chr(13) & Chr(10) & _
              "Dibayar " & Space(31 - Len(Format(Dibayar - Voucher, "###,###,##0"))) & Format(Dibayar - Voucher, "###,###,##0") & Chr(13) & Chr(10) & _
               "Kembali " & Space(31 - Len(Format(Kembali, "###,###,##0"))) & Format(Kembali, "###,###,##0") & Chr(13) & Chr(10) & _
              IIf(ISSembunyikanFooterbarangKSB, "", "Belanja Barang POIN" & Space(20 - Len(Format(BelanjaPoin, "###,###,##0"))) & Format(BelanjaPoin, "###,###,##0") & Chr(13) & Chr(10)) & _
              "#" & Format(CLng(NoNota), "0#####") & "," & "Kasir" & NamaMesin & ":" & Mid(NamaKasir, 1, 6) & Space(6 - IIf(Len(NamaKasir) <= 6, Len(NamaKasir), 6)) & Format(Date, " dd/MM/yy") & Format(Time, "hh:mm:ss") & Chr(13) & Chr(10) & _
              "Macam Barang =" & Format(MacamItem, "##0") & " , Total Item =" & Format(Jumlahitem, "##0") & Chr(13) & Chr(10) & _
              IIf(KodeMember = "", "", "CUSTOMER: " & KodeMember & "-" & NamaMember & Chr(13) & Chr(10)) & _
              "----TERIMA KASIH ATAS KUNJUNGAN ANDA----" & Chr(13) & Chr(10) & _
              IIf(Footer1 = "", "", Footer1 & Chr(13) & Chr(10)) & _
              IIf(Footer2 = "", "", Footer2 & Chr(13) & Chr(10)) & _
              IIf(Footer3 = "", "", Footer3 & Chr(13) & Chr(10)) & _
              Chr(13) & Chr(10)
              
              PrinBigChar Space(15 - Min(15, Len(NamaToko))) & UCase(NamaToko)
              Prin Judulstruk & Chr(13) & Chr(10) & "---------------------------------------"
    papercut
    DoEvents
End Sub
Sub ReprintNew()
Dim cetak As String
Dim kode As String
Dim Nama As String
Dim Qty As String
Dim hargaSatuan As String
Dim Jumlah As String
    'Prin Judulstruk
    PrinBigChar "----REPRINT----"
    cetak = "---------------------------------------" & Chr(13) & Chr(10)
    If Data1.Recordset.EOF And Data1.Recordset.BOF Then Exit Sub
    Data1.Recordset.MoveFirst
    Do While Not Data1.Recordset.EOF
        kode = Data1.Recordset!KodeInv
        Nama = Data1.Recordset!NamaInv
        Qty = Format(Data1.Recordset!Qty, "##0")
        hargaSatuan = Format(Data1.Recordset!harga, "###,###,##0")
        Jumlah = Format(Data1.Recordset!Qty * Data1.Recordset!harga, "###,###,##0")
       cetak = cetak & Nama & Chr(13) & Chr(10) & kode & Space(16 - Min(Len(Left(kode, 13)) + Len(Qty), 16)) & Qty & " X " & Space(8 - Min(Len(hargaSatuan), 8)) & hargaSatuan & "=" & Space(11 - Len(Jumlah)) & Jumlah
      Data1.Recordset.MoveNext
    Loop
    
    cetak = cetak & Chr(13) & Chr(10) & "---------------------------------------" & Chr(13) & Chr(10)
    
    cetak = cetak & "Subtotal" & Space(31 - Len(Format(SubTotal, "###,###,##0"))) & Format(SubTotal, "###,###,##0") & Chr(13) & Chr(10) & _
                "Voucher " & Space(31 - Len(Format(Voucher, "###,###,##0"))) & Format(Voucher, "###,###,##0") & Chr(13) & Chr(10) & _
              "Dibayar " & Space(31 - Len(Format(Dibayar - Voucher, "###,###,##0"))) & Format(Dibayar - Voucher, "###,###,##0") & Chr(13) & Chr(10) & _
               "Kembali " & Space(31 - Len(Format(Kembali, "###,###,##0"))) & Format(Kembali, "###,###,##0") & Chr(13) & Chr(10) & _
              IIf(ISSembunyikanFooterbarangKSB, "", "Belanja Barang POIN" & Space(20 - Len(Format(BelanjaPoin, "###,###,##0"))) & Format(BelanjaPoin, "###,###,##0") & Chr(13) & Chr(10)) & _
              "#" & Format(CLng(NoNota), "0#####") & "," & "Kasir" & NamaMesin & ":" & Mid(NamaKasir, 1, 6) & Space(6 - IIf(Len(NamaKasir) <= 6, Len(NamaKasir), 6)) & Format(Date, " dd/MM/yy") & Format(Time, "hh:mm:ss") & Chr(13) & Chr(10) & _
              "Macam Barang =" & Format(MacamItem, "##0") & " , Total Item =" & Format(Jumlahitem, "##0") & Chr(13) & Chr(10) & _
              IIf(KodeMember = "", "", "CUSTOMER: " & KodeMember & "-" & NamaMember & Chr(13) & Chr(10)) & _
              "----TERIMA KASIH ATAS KUNJUNGAN ANDA----" & Chr(13) & Chr(10) & _
              IIf(Footer1 = "", "", Footer1 & Chr(13) & Chr(10)) & _
              IIf(Footer2 = "", "", Footer2 & Chr(13) & Chr(10)) & _
              IIf(Footer3 = "", "", Footer3 & Chr(13) & Chr(10)) & _
              Chr(13) & Chr(10)
              
              PrinBigChar Space(15 - Min(15, Len(NamaToko))) & UCase(NamaToko)
              Prin Judulstruk & Chr(13) & Chr(10) & "---------------------------------------"
    Prin cetak
    DoEvents
    papercut
End Sub

Sub ITEMKoreksi(ByVal Transaksi As String)
Dim DefHargaJual As Double
    If Transaksi = "VOD" Then
        If BolehVoid(rs!kode, Qty) = False Then
            frmPesan.lbPesan = "BARANG DIVOID HARUS ADA!!!!"
            Qty = 1
            lbQty = Format(Qty, "##0") & " X"
            frmPesan.Show 1
            Exit Sub
        End If
    End If
    Dim dbHistori As Database
    Dim RsHistori As Recordset
    Dim HargaJual As Double
    IDSalesD = GetNewID("MSALESD")
'HARGA MENGGUNAKAN MULTI HARGA
        If Abs(Qty) >= NullToNol(rs!Qty3) And NullToNol(rs!Qty3) <> 0 Then
              DefHargaJual = NullToNol(rs!Harga3)
          ElseIf Abs(Qty) >= NullToNol(rs!Qty2) And NullToNol(rs!Qty2) <> 0 Then
              DefHargaJual = NullToNol(rs!Harga2)
          ElseIf Abs(Qty) >= NullToNol(rs!Qty1) And NullToNol(rs!Qty1) <> 0 Then
              DefHargaJual = NullToNol(rs!Harga1)
          Else
              DefHargaJual = NullToNol(rs!HargaJual)
          End If
'    If TipeHargaJual = 1 Then
'      DefHargaJual = NullToNol(rs!HargaA)
'    ElseIf TipeHargaJual = 2 Then
'      DefHargaJual = NullToNol(rs!HargaB)
'    ElseIf TipeHargaJual = 3 Then
'      DefHargaJual = NullToNol(rs!HargaC)
'    ElseIf TipeHargaJual = 4 Then
'      DefHargaJual = NullToNol(rs!HargaD)
'    ElseIf TipeHargaJual = 5 Then
'      DefHargaJual = NullToNol(rs!HargaE)
'    ElseIf TipeHargaJual = 6 Then
'      DefHargaJual = NullToNol(rs!HargaF)
'    Else
      'DefHargaJual = NullToNol(rs!HargaJual)
'    End If
    
    If DefHargaJual = 0 Then
        If Transaksi = "RTN" Or Transaksi = "VOD" Then
          frmsetHarga.Tampil HargaJual, rs!kode & " " & rs!Nama
        Else
          HargaJual = Data1.Recordset!harga
        End If
    Else
      HargaJual = DefHargaJual
    End If
    Data1.Recordset.AddNew
    Data1.Recordset!NoID = IDSalesD
    Data1.Recordset!IDSales = IDSales
    Data1.Recordset!IDInvSat = rs!NoID
    Data1.Recordset!IdInventor = rs!IdInventor
    Data1.Recordset!Qty = Qty
    Data1.Recordset!HargaBruto = HargaJual
    Data1.Recordset!harga = HargaJual
    Data1.Recordset!KodeInv = rs!kode
    Data1.Recordset!NamaInv = rs!Nama
    Data1.Recordset!Satuan = IIf(IsNull(rs!KodeSat), "", rs!KodeSat)
    Data1.Recordset!Barcode = IIf(IsNull(rs!Barcode), rs!kode, rs!Barcode)
    Data1.Recordset!idSatuan = rs!idSatuan
    Data1.Recordset!Konversi = rs!Konversi
    Data1.Recordset!Transaksi = Transaksi
    Data1.Recordset!HargaPokok = rs!HargaPokok
    Data1.Recordset!IsPoin = rs!IsPoin
    Data1.Recordset!IsPoinSupplier = rs!IsPoinSupplier
    Data1.Recordset!IDPoinSupplier = rs!IDPoinSupplier
    Data1.Recordset!BKP = rs!BKP
    Data1.Recordset!DiscRp = 0
    Data1.Recordset!DiscProsen = 0
    Data1.Recordset!DiscInternRp = 0
    Data1.Recordset!DiscInternProsen = 0

'    If TipeHargaJual = 1 Then
'      Data1.Recordset!DiscProsen1 = NullToNol(rs!DiscProsen1A)
'      Data1.Recordset!DiscProsen2 = NullToNol(rs!DiscProsen2A)
'      Data1.Recordset!HargaMin = NullToNol(rs!HargaMinA)
'     ElseIf TipeHargaJual = 2 Then
'      Data1.Recordset!DiscProsen1 = NullToNol(rs!DiscProsen1B)
'      Data1.Recordset!DiscProsen2 = NullToNol(rs!DiscProsen2B)
'      Data1.Recordset!HargaMin = NullToNol(rs!HargaMinB)
'     ElseIf TipeHargaJual = 3 Then
'      Data1.Recordset!DiscProsen1 = NullToNol(rs!DiscProsen1C)
'      Data1.Recordset!DiscProsen2 = NullToNol(rs!DiscProsen2C)
'      Data1.Recordset!HargaMin = NullToNol(rs!HargaMinC)
'     ElseIf TipeHargaJual = 4 Then
'      Data1.Recordset!DiscProsen1 = NullToNol(rs!DiscProsen1D)
'      Data1.Recordset!DiscProsen2 = NullToNol(rs!DiscProsen2D)
'      Data1.Recordset!HargaMin = NullToNol(rs!HargaMinD)
'     ElseIf TipeHargaJual = 5 Then
'      Data1.Recordset!DiscProsen1 = NullToNol(rs!DiscProsen1E)
'      Data1.Recordset!DiscProsen2 = NullToNol(rs!DiscProsen2E)
'      Data1.Recordset!HargaMin = NullToNol(rs!HargaMinE)
'     ElseIf TipeHargaJual = 6 Then
'      Data1.Recordset!DiscProsen1 = NullToNol(rs!DiscProsen1F)
'      Data1.Recordset!DiscProsen2 = NullToNol(rs!DiscProsen2F)
'      Data1.Recordset!HargaMin = NullToNol(rs!HargaMinF)
'     Else
'      Data1.Recordset!DiscProsen1 = NullToNol(rs!DiscProsen1B)
'      Data1.Recordset!DiscProsen2 = NullToNol(rs!DiscProsen2B)
'      Data1.Recordset!HargaMin = NullToNol(rs!HargaMinB)
'     End If
'     Data1.Recordset!DiscProsen3 = 0
'     Data1.Recordset!Disc1 = 0
'     Data1.Recordset!Disc2 = 0
'     Data1.Recordset!Disc3 = 0
     Data1.Recordset!Jumlah = Bulatkan(Qty * HargaJual, 0) '(Bulatkan((HargaJual * (1 - (NullToNol(Data1.Recordset!DiscProsen1) / 100)) * (1 - (NullToNol(Data1.Recordset!DiscProsen2) / 100)) * (1 - (NullToNol(Data1.Recordset!DiscProsen3) / 100))), 0) - 0 - 0 - 0)
     
    Data1.Recordset.Update
    Data1.Recordset.Bookmark = Data1.Recordset.LastModified
    DisplayPesan Data1.Recordset!NamaInv, Format(Qty, "##0") & " X " & Space(17 - Len(Format(Qty, "##0")) - Len(Format(Data1.Recordset!harga, "###,###,##0"))) & Format(Data1.Recordset!harga, "###,###,##0")
    If Data1.Recordset!IsMember Then
      cetakdetil Data1.Recordset!KodeInv, Data1.Recordset!NamaInv, Format(Data1.Recordset!Qty, "##0"), Format(Data1.Recordset!harga, "###,###,##0"), Format(Qty * Data1.Recordset!harga, "###,###,##0")
    Else
      cetakdetil Data1.Recordset!KodeInv, Data1.Recordset!NamaInv & "*)", Format(Data1.Recordset!Qty, "##0"), Format(Data1.Recordset!harga, "###,###,##0"), Format(Qty * Data1.Recordset!harga, "###,###,##0")
    End If
    Qty = 1
    lbQty = Format(Qty, "##0") & " X"
    Set dbHistori = OpenDatabase(DirDatabase & "\Histori.mdb")
    Set RsHistori = dbHistori.OpenRecordset("Historikasir")
    RsHistori.AddNew
    'ID
    RsHistori!kassa = NamaMesin
    RsHistori!IDUser = IDUser
    RsHistori!KodeUser = KodeKasir
    RsHistori!Tanggal = Date
    RsHistori!Jam = Time
    RsHistori!IDSales = IDSales
    RsHistori!IDSalesD = IDSalesD
    RsHistori!Transaksi = Transaksi
    RsHistori!IdInventor = Data1.Recordset!IdInventor
    RsHistori!idSatuan = Data1.Recordset!idSatuan
    RsHistori!KodeInventor = Data1.Recordset!KodeInv
    RsHistori!NamaInventor = Data1.Recordset!NamaInv
    RsHistori!HargaPokok = Data1.Recordset!HargaPokok
    RsHistori!HargaJualMaster = Data1.Recordset!harga
    RsHistori!HargaJualKhusus = Data1.Recordset!harga
    RsHistori.Update
    RsHistori.Close
    dbHistori.Close
        If NullToNol(rs!DiscMemberRp2) > 0 Then
'          If Not IsNull(rs!TglDariDiskon2) And Not IsNull(rs!TglSampaiDiskon2) And Date >= NullToDate(rs!TglDariDiskon2) And Date <= NullToDate(rs!TglSampaiDiskon2) And IsQtyPDPAda(NullToNol(rs!IdInventor), Qty, NullToDate(rs!TglDariDiskon2), NullToDate(rs!TglSampaiDiskon2), CDbl(SubTotal)) Then
'            DiscInternProsen = NullToNol(rs!DiscProsen)
'            If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
'                HargaBruto = 0
'            Else
'                HargaBruto = Data1.Recordset!HargaBruto
'            End If
'            If HargaBruto <> 0 Then
'              DiscInternRp = NullToNol(rs!DiscMemberRp2) 'NilaiDiskon ' ((HargaBruto * DiscInternProsen / 100) \ 1) * 1  'tanpa koma
'              DiscInternProsen = NullToNol(rs!DiscMemberProsen2)
'              UpdateJualD
'              Text1.Text = ""
'            End If
'          Else 'Kemungkinan ke Promo Diskon
'            If NullToNol(rs!DiscRupiah) > 0 And Not IsNull(rs!DiscExpired) And Not IsNull(rs!DiscMulai) Then
'              If Date >= rs!DiscMulai And Date <= rs!DiscExpired Then
'                  DiscInternProsen = NullToNol(rs!DiscProsen)
'                  If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
'                      HargaBruto = 0
'                  Else
'                      HargaBruto = Data1.Recordset!HargaBruto
'                  End If
'                  If HargaBruto <> 0 Then
'                    DiscInternRp = NullToNol(rs!DiscRupiah) 'NilaiDiskon ' ((HargaBruto * DiscInternProsen / 100) \ 1) * 1  'tanpa koma
'                    DiscInternProsen = NullToNol(rs!DiscProsen)
'                    UpdateJualD
'                    Text1.Text = ""
'                  End If
'              End If
'            End If
'          End If
        ElseIf NullToNol(rs!DiscRupiah) > 0 Then
          If Not IsNull(rs!DiscExpired) And Not IsNull(rs!DiscMulai) Then
              If Date >= rs!DiscMulai And Date <= rs!DiscExpired Then
                  DiscInternProsen = NullToNol(rs!DiscProsen)
                  If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
                      HargaBruto = 0
                  Else
                      HargaBruto = Data1.Recordset!HargaBruto
                  End If
                  If HargaBruto <> 0 Then
                    DiscInternRp = NullToNol(rs!DiscRupiah) 'NilaiDiskon ' ((HargaBruto * DiscInternProsen / 100) \ 1) * 1  'tanpa koma
                    DiscInternProsen = NullToNol(rs!DiscProsen)
                    UpdateJualD
                    Text1.Text = ""
                  End If
              End If
          End If
        End If
    HitungSubTotal
    If IDMember >= 1 Then
      UpdatekanDiskonPerBarang
    End If
'  If Transaksi = "VOD" Then
'        If BolehVoid(rs!kode, QTY) = False Then
'            frmPesan.lbPesan = "BARANG DIVOID HARUS ADA!!!!"
'            QTY = 1
'            lbQTY = Format(QTY, "##0") & " X"
'            frmPesan.Show 1
'            Exit Sub
'        End If
'    End If
'    Dim dbHistori As Database
'    Dim RsHistori As Recordset
'    Dim HargaJual As Double
'    IDSalesD = GetNewID("MSALESD")
'    If rs!HargaJual = 0 Then
'        If Transaksi = "RTN" Or Transaksi = "VOD" Then
'          frmsetHarga.Tampil HargaJual, rs!kode & " " & rs!Nama
'        Else
'          HargaJual = Data1.Recordset!harga
'        End If
'    Else
'      HargaJual = rs!HargaJual
'    End If
'    Data1.Recordset.AddNew
'    Data1.Recordset!NoId = IDSalesD
'    Data1.Recordset!IDSales = IDSales
'    Data1.Recordset!IdInventor = rs!NoId
'    Data1.Recordset!QTY = QTY
'    Data1.Recordset!HargaBruto = HargaJual
'    Data1.Recordset!harga = HargaJual
'    Data1.Recordset!KodeInv = rs!kode
'    Data1.Recordset!NamaInv = rs!Nama
'    Data1.Recordset!Satuan = IIf(IsNull(rs!KodeSat), "", rs!KodeSat)
'    Data1.Recordset!Barcode = IIf(IsNull(rs!Barcode), rs!kode, rs!Barcode)
'    Data1.Recordset!idSatuan = rs!idSatuan
'    Data1.Recordset!Konversi = rs!Konversi
'    Data1.Recordset!Transaksi = Transaksi
'    Data1.Recordset!HargaPokok = rs!HargaPokok
'
'    Data1.Recordset.Update
'    Data1.Recordset.Bookmark = Data1.Recordset.LastModified
'    DisplayPesan Data1.Recordset!NamaInv, Format(QTY, "##0") & " X " & Space(17 - Len(Format(QTY, "##0")) - Len(Format(Data1.Recordset!harga, "###,###,##0"))) & Format(Data1.Recordset!harga, "###,###,##0")
'    cetakdetil Data1.Recordset!KodeInv, Data1.Recordset!NamaInv, Format(Data1.Recordset!QTY, "##0"), Format(Data1.Recordset!harga, "###,###,##0"), Format(QTY * Data1.Recordset!harga, "###,###,##0")
'    QTY = 1
'    lbQTY = Format(QTY, "##0") & " X"
'    Set dbHistori = OpenDatabase(DirDatabase & "\Histori.mdb")
'    Set RsHistori = dbHistori.OpenRecordset("Historikasir")
'    RsHistori.AddNew
'    'ID
'    RsHistori!Kassa = NamaMesin
'    RsHistori!IDUser = IDUser
'    RsHistori!KodeUser = KodeKasir
'    RsHistori!Tanggal = Date
'    RsHistori!Jam = Time
'    RsHistori!IDSales = IDSales
'    RsHistori!IDSalesD = IDSalesD
'    RsHistori!Transaksi = Transaksi
'    RsHistori!IdInventor = Data1.Recordset!IdInventor
'    RsHistori!idSatuan = Data1.Recordset!idSatuan
'    RsHistori!KodeInventor = Data1.Recordset!KodeInv
'    RsHistori!NamaInventor = Data1.Recordset!NamaInv
'    RsHistori!HargaPokok = Data1.Recordset!HargaPokok
'    RsHistori!HargaJualMaster = Data1.Recordset!harga
'    RsHistori!HargaJualKhusus = Data1.Recordset!harga
'    RsHistori.Update
'    RsHistori.Close
'    dbHistori.Close
'    UpdateJualD
'    HitungSubTotal
    
End Sub
Function BolehVoid(ByVal kodeVoid As String, ByVal Qty As Double) As Boolean
    Dim dbs As Database
    Dim rs As Recordset
    Set dbs = OpenDatabase(App.path & "\DATABASE\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
    Set rs = dbs.OpenRecordset("SELECT SUM(qty) AS Jml FROM MSALESD WHERE  IDSALES=" & IDSales & " AND KodeInv='" & kodeVoid & "'")
    If rs.EOF Or rs.BOF Then
        BolehVoid = False
    Else
        If IsNull(rs!Jml) Then
            BolehVoid = False
        Else
            If rs!Jml < Abs(Qty) Then
                BolehVoid = False
            Else
                BolehVoid = True
            End If
        End If
    End If
End Function
Function IsTambahItem(ByVal Barcode As String, ByVal Qty As Double, Optional ByVal IsBuah As Boolean = False) As Double
    Dim dbs As Database
    Dim rs As Recordset
    Set dbs = OpenDatabase(App.path & "\DATABASE\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
    Set rs = dbs.OpenRecordset("SELECT NoID,Qty FROM MSALESD WHERE IDSALES=" & IDSales & " AND Transaksi='PLU' AND Barcode='" & Barcode & "'")
    If rs.EOF Or rs.BOF Then
        IsTambahItem = 0
    Else
        If IsOpenDepartemen(Barcode) Or IsBuah Then
          IsTambahItem = 0
        Else
          If IsNull(rs!NoID) Then
              IsTambahItem = 0
          Else
              If rs!NoID <= 0 Then
                  IsTambahItem = 0
              Else
                  'dbs.Execute "delete * from MsalesD where NoID=" & rs!NoID
                  IsTambahItem = NullToNol(rs!Qty)
                    Data1.Recordset.FindFirst "NoID=" & rs!NoID
                    If Data1.Recordset.NoMatch Then
                    Else
                      Data1.Recordset.Delete
                    End If
  
              End If
          End If
        End If
    End If
rs.Close
Set rs = Nothing
dbs.Close
Set dbs = Nothing
DoEvents
End Function
    
Function IsItemAda(ByVal Kodebarang As String) As Boolean
    Dim dbs As Database
    Dim rs As Recordset
    Set dbs = OpenDatabase(App.path & "\DATABASE\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
    Set rs = dbs.OpenRecordset("SELECT COUNT(NoID) AS Jml FROM MSALESD WHERE  IDSALES=" & IDSales & " AND KodeInv='" & Kodebarang & "'")
    If rs.EOF Or rs.BOF Then
        IsItemAda = False
    Else
        If NullToNol(rs!Jml) > 0 Then
            IsItemAda = True
        Else
            IsItemAda = True
        End If
    End If
End Function
    
Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Timer1_Timer()
'  lbJam = Format(Time, "HH : mm : ss")
'Timer1.Enabled = False
'    Text1.SetFocus
End Sub
Function DiscountPLUTampil(ByVal pbarcode As String) As Boolean
Dim rsTot As Recordset
Set rsTot = dbSale.OpenRecordset("SELECT MSalesD.DiscRp,MSalesD.DiscInternRp,MSalesD.DiscProsen,MSalesD.DiscInternProsen From MSalesD where Transaksi='PLU' And Barcode='" & pbarcode & "' AND IDSales=" & IDSales)
If rsTot.BOF And rsTot.EOF Then
  DiscountPLUTampil = False
    DiscRp = 0
    DiscProsen = 0
    DiscInternRp = 0
    DiscInternProsen = 0
    
Else
    DiscRp = NullToNol(rsTot!DiscRp)
    DiscProsen = NullToNol(rsTot!DiscProsen)
    DiscInternRp = NullToNol(rsTot!DiscInternRp)
    DiscInternProsen = NullToNol(rsTot!DiscInternProsen)
    
  If IsNull(rsTot!DiscRp) And IsNull(rsTot!DiscRp) Then
    DiscountPLUTampil = False
  ElseIf rsTot!DiscRp > 0 Or rsTot!DiscInternRp > 0 Then
  DiscountPLUTampil = True
  Else
  DiscountPLUTampil = False
    'DiscountPLUTampil = IIf(rsTot!DiscRp > 0, True, False)
  End If
End If
End Function
Sub HitungSubTotal()
Dim rsTot As Recordset 'SUM(int(MSalesD.Harga*MSalesD.Qty+0.5)) as Total,
Set rsTot = dbSale.OpenRecordset("SELECT SUM(Jumlah) as Total, SUM(int(MSalesD.DiscRp*MSalesD.Qty+0.5)) as TotalDiscBrg," & _
            "SUM(int(MSalesD.DiscInternRp*MSalesD.Qty+0.5)) as TotalDiscInternBrg From MSalesD where IDSales=" & IDSales)
If rsTot.BOF And rsTot.EOF Then
  SubTotal = 0
  TotalDiscBrg = 0
  JumDiscInternRp = 0
Else
  If IsNull(rsTot!Total) Then
    SubTotal = 0
  Else
    SubTotal = rsTot!Total
  End If
  
  If IsNull(rsTot!TotalDiscBrg) Then
    TotalDiscBrg = 0
  Else
    TotalDiscBrg = rsTot!TotalDiscBrg '(rsTot!TotalDiscBrg \ DefPembulatan) * DefPembulatan
  End If
  
  If IsNull(rsTot!TotalDiscInternBrg) Then
    JumDiscInternRp = 0
  Else
    JumDiscInternRp = rsTot!TotalDiscInternBrg '(rsTot!TotalDiscInternBrg \ DefPembulatan) * DefPembulatan
  End If
End If
rsTot.Close
Set rsTot = dbSale.OpenRecordset("SELECT SUM(Jumlah) as JUMBKP From MSalesD  where MSalesD.BKP=True and IDSales=" & IDSales)
If rsTot.BOF And rsTot.EOF Then
  JumlahBKP = 0
  JumlahDPP = 0
  JumlahPPN = 0
Else
  JumlahBKP = NullToNol(rsTot!JUMBKP)
  JumlahDPP = Bulatkan(JumlahBKP / 1.1, 0)
  JumlahPPN = JumlahBKP - JumlahDPP ' Bulatkan(JumlahDPP * 0.1, 0)
End If
rsTot.Close
Set rsTot = dbSale.OpenRecordset("SELECT SUM(Jumlah) as JumBarangPoin From MSalesD  where MSalesD.IsPoin=True and IDSales=" & IDSales) 'and DiscInternRp=0
If rsTot.BOF And rsTot.EOF Then
  BelanjaPoin = 0
Else
  BelanjaPoin = NullToNol(rsTot!JumBarangPoin)
End If
rsTot.Close
If defMinialBelanjaDapatPDP > 0 Then
  Set rsTot = dbSale.OpenRecordset("SELECT SUM(Jumlah) as JumBarangPDP From MSalesD where MSalesD.IsPDP=False AND IDSales=" & IDSales) 'and DiscInternRp=0
  If rsTot.BOF And rsTot.EOF Then
    lbPDP.Visible = False
  Else
    If NullToNol(rsTot!JumBarangPDP) >= defMinialBelanjaDapatPDP Then
      lbPDP.Visible = True
      lbPDP.Caption = "Selamat anda bisa mendapatkan barang-barang tertentu dengan harga khusus."
    Else
      lbPDP.Visible = False
    End If
  End If
  rsTot.Close
Else
  lbPDP.Visible = False
End If
'End If
HitungPoin
 'HitungDiscountOtomatis
DoEvents
End Sub
Sub HitungPoin()
Dim rs As Recordset
Dim dbs As Database, i As Integer
Set dbs = OpenDatabase(App.path & "\database\dbmaster.mdb")
Set rs = dbs.OpenRecordset("SELECT * from MSettingPoin")

If rs.BOF And rs.EOF Then
  defNilaiDiskonMember = 0
  defBelanjadapat1Poin = 0
  defMinialBelanjadapatDiskon = 0
  defMinialBelanjadapatDiskon2 = 0
  defMinialBelanjaDapatPDP = 0
  defIsCCDapatDiskon = False
  'Ambil per customer
  defNilaiDiskonMember = defDiscMember
  defMinialBelanjadapatDiskon = defMinMemberdapatDisc
  defMinialBelanjadapatDiskon2 = defMinMemberdapatDisc
  defMinialBelanjaDapatPDP = 0
  defIsCCDapatDiskon = True
Else
  defNilaiDiskonMember = NullToNol(rs!NilaiDiskon)
  defBelanjadapat1Poin = NullToNol(rs!NilaiBelanja1Poin)
  defMinialBelanjadapatDiskon = NullToNol(rs!MinimumBelanjaDapatPoin)
  defMinialBelanjadapatDiskon2 = NullToNol(rs!MinimumBelanjaDapatPoin2)
  defMinialBelanjaDapatPDP = NullToNol(rs!MinimumBelanjaDapatPDP)
  defIsCCDapatDiskon = NullToBool(rs!CreditCardDapatDiskon)
End If
rs.Close
Set rs = Nothing
dbs.Close
Set dbs = Nothing

If IDMember > 0 And (Bank = 0 Or (Bank > 0 And IskartuKredit = True And defIsCCDapatDiskon = False)) Then
  If BelanjaPoin >= defMinialBelanjadapatDiskon Then
    DiscINTERNNOTA = BelanjaPoin * defNilaiDiskonMember / 100
  Else
    DiscINTERNNOTA = 0
  End If
Else
    DiscINTERNNOTA = 0
End If
  If defBelanjadapat1Poin = 0 Then
    PoinNotaIni = 0
  Else
    PoinNotaIni = BelanjaPoin \ defBelanjadapat1Poin
  End If
  TampilBawah

'DisplayPesan "DISC NOTA #" & Space(9 - IIf(Len(Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0")) > 9, 0, Len(Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0")))) & Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0"), "TOTAL     #" & Space(9 - IIf(Len(Format(Total, "###,###,##0")) > 9, 0, Len(Format(Total, "###,###,##0")))) & Format(Total, "###,###,##0")

End Sub

Sub HitungDiscountOtomatis()
'    Dim typediscount As Byte '1bertingkat, 2 terbesar
'    If Dir(App.Path & "\database\Mdiscount.mdb") = "" Then Exit Sub
'    Dim dbsdis As Database
'    Dim rsdis As Recordset
'    Set dbsdis = OpenDatabase(App.Path & "\database\Mdiscount.mdb")
'    Set rsdis = dbsdis.OpenRecordset("select * from MTypeDiscount")
'    If rsdis.EOF And rsdis.BOF Then
'      typediscount = 2
'    Else
'      typediscount = rsdis!NoId
'    End If
'    rsdis.Close
'    Set rsdis = dbsdis.OpenRecordset("select * from MDiscount order BY NoID")
'    If rsdis.EOF And rsdis.BOF Then
'    Else
'      rsdis.MoveFirst
'      If typediscount = 2 Then
'          Do While Not rsdis.EOF
'            If SubTotal >= rsdis!SubTotalDari And SubTotal < rsdis!SubTotalSampai Then
'              Disc = Round(rsdis!DiscountProsen * SubTotal \ 100, 0) + rsdis!DiscountRupiah
'              Exit Do
'            End If
'
'            rsdis.MoveNext
'          Loop
'      End If
'    End If
'    dbsdis.Close
End Sub


Sub UpdateJualD(Optional ByVal tombol As String)
    Dim HargaJual As Double
    If Data1.Recordset!Transaksi = "VOD" And tombol = "" Then
        DiscountPLUTampil (Data1.Recordset!Barcode) 'ambil discount aslinya
    End If
    HargaBruto = Data1.Recordset!HargaBruto
    HargaJual = HargaBruto - DiscRp - DiscInternRp
    Dim jwb As Boolean
    bolehbergerak = False
    Data1.Recordset.Edit
    Data1.Recordset!HargaBruto = HargaBruto
    Data1.Recordset!harga = HargaJual
    Data1.Recordset!DiscRp = DiscRp
    Data1.Recordset!DiscProsen = DiscProsen
    Data1.Recordset!DiscInternRp = DiscInternRp
    Data1.Recordset!DiscInternProsen = DiscInternProsen
'    Data1.Recordset!Disc1 = DiscRp
'    Data1.Recordset!Disc2 = 0
'    Data1.Recordset!Disc3 = 0
'    Data1.Recordset!HargaD = Bulatkan((Data1.Recordset!HargaBruto * (1 - (NullToNol(Data1.Recordset!DiscProsen1) / 100)) * (1 - (NullToNol(Data1.Recordset!DiscProsen2) / 100)) * (1 - (NullToNol(Data1.Recordset!DiscProsen3) / 100))), 0)
    Data1.Recordset!Jumlah = Bulatkan(Data1.Recordset!Qty * HargaJual, 0) ' (Bulatkan((HargaBruto * (1 - (NullToNol(Data1.Recordset!DiscProsen1) / 100)) * (1 - (NullToNol(Data1.Recordset!DiscProsen2) / 100)) * (1 - (NullToNol(Data1.Recordset!DiscProsen3) / 100))), 0) - Data1.Recordset!Disc1 - Data1.Recordset!Disc2 - Data1.Recordset!Disc3)
'    Data1.Recordset!IsDiscSupplier = IsDiscBySupplier
    
    Data1.Recordset.Update
    Data1.Recordset.Bookmark = Data1.Recordset.LastModified
    'cetakdetil Data1.Recordset!kodeinv, Data1.Recordset!namaInv, Format(Data1.Recordset!QTY, "##0"), Format(Data1.Recordset!harga, "###,###,##0"), Format(QTY * Data1.Recordset!harga, "###,###,##0")
    'If DiscRp > 0 Or DiscInternRp > 0 Then
    ' Prin "*** Disc. " & Format(DiscProsen + DiscInternProsen, "#0.00") & "% " & Space(39 - Len("*** Disc. " & Format(DiscProsen + DiscInternProsen, "#0.00") & "% " & "-" & Format((DiscRp + DiscInternRp) * Data1.Recordset!QTY, "###,##0"))) & "-" & Format((DiscRp + DiscInternRp) * Data1.Recordset!QTY, "###,##0") 'Data1.Recordset!KodeInv, Data1.Recordset!NamaInv, Format(Data1.Recordset!QTY, "##0"), Format(Data1.Recordset!harga, "###,###,##0"), Format(QTY * Data1.Recordset!harga, "###,###,##0")
     '38
     'Prin Nama & Chr(13) & Chr(10) & kode & Space(16 - Min(Len(Left(kode, 13)) + Len(QTY), 16)) & QTY & " X " & Space(8 - Min(Len(hargaSatuan), 8)) & hargaSatuan & "=" & Space(11 - Len(Jumlah)) & Jumlah
   'End If
    
    If (tombol = "DBP" Or tombol = "DBR") And DiscRp > 0 Then
     Prin "*** Disc. " & Format(DiscProsen, "#0.00") & "% " & Space(39 - Len("*** Disc. " & Format(DiscProsen, "#0.00") & "% " & "-" & Format((DiscRp) * Data1.Recordset!Qty, "###,##0"))) & "-" & Format((DiscRp) * Data1.Recordset!Qty, "###,##0")     'Data1.Recordset!KodeInv, Data1.Recordset!NamaInv, Format(Data1.Recordset!QTY, "##0"), Format(Data1.Recordset!harga, "###,###,##0"), Format(QTY * Data1.Recordset!harga, "###,###,##0")
    
    End If
    
    If (tombol = "DIP" Or tombol = "DIR") And DiscInternRp > 0 Then
     Prin "*** Disc. " & Format(DiscInternProsen, "#0.00") & "% " & Space(39 - Len("*** Disc. " & Format(DiscInternProsen, "#0.00") & "% " & "-" & Format((DiscInternRp) * Data1.Recordset!Qty, "###,##0"))) & "-" & Format((DiscInternRp) * Data1.Recordset!Qty, "###,##0") 'Data1.Recordset!KodeInv, Data1.Recordset!NamaInv, Format(Data1.Recordset!QTY, "##0"), Format(Data1.Recordset!harga, "###,###,##0"), Format(QTY * Data1.Recordset!harga, "###,###,##0")
    
    End If
    
    
    DisplayPesan NullToStr(Data1.Recordset!NamaInv), "Discount" & Space(Max(12 - Len(Format(DiscRp + DiscInternRp, "###,##0")), 0)) & Format(DiscRp + DiscInternRp, "###,##0") & ""
 'Jika discount dari Sendiri otomatis ambil nettnya
  'Data1.Recordset!harga = HargaJual
'  If Not IsDiscBySupplier And Data1.Recordset!Transaksi = "PLU" Then
'    Data1.Recordset.Edit
'    Data1.Recordset!HargaBruto = HargaJual
'    Data1.Recordset!harga = HargaJual
'    Data1.Recordset!DiscRp = 0 'DiscRp
'    Data1.Recordset!DiscProsen = 0 'DiscProsen
'    Data1.Recordset!DiscInternRp = 0 'DiscRp
'    Data1.Recordset!DiscInternProsen = 0 'DiscProsen
'    Data1.Recordset.Update
'    Data1.Recordset.Bookmark = Data1.Recordset.LastModified
'  ElseIf IsDiscBySupplier And Data1.Recordset!Transaksi = "VOD" Then
'    If Not DiscountPLUTampil(Data1.Recordset!BARCODE) Then
'        'Data1.Recordset.Edit
'        'Data1.Recordset!HargaBruto = HargaBruto
'        'Data1.Recordset!harga = HargaJual
'        'Data1.Recordset!DiscRp = 0 'DiscRp
'        'Data1.Recordset!DiscProsen = 0 'DiscProsen
'        'Data1.Recordset!DiscInternRp = 0 'DiscRp
'        'Data1.Recordset!DiscInternProsen = 0 'DiscProsen
'        'Data1.Recordset.Update
'        'Data1.Recordset.Bookmark = Data1.Recordset.LastModified
'    End If
' End If
    HitungSubTotal
        
End Sub

Sub UpdateFooter()
  'Footer
  If Dir(DirDatabase & "\Setting.mdb") = "" Then
     Footer1 = ""
     Footer2 = ""
    Footer3 = ""
  Else
    Dim dbs As Database
    Dim rs As Recordset
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
End Sub

Sub UpdateSales()
Err.Clear
On Error GoTo Trace
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Dim les
Dim rsSales As Recordset
Set rsSales = dbSale.OpenRecordset("Select * FROM MSales Where NoID=" & IDSales)
'If Time <= "15:30:00" Then
'    NamaShift = 1
'Else
'    NamaShift = 2
'End If
If rsSales.EOF And rsSales.BOF Then
  rsSales.AddNew
  rsSales!NoID = IDSales
  rsSales!kode = NoNota 'Format(IDSales, "0######")
  rsSales!Tanggal = Now
  rsSales!TotalBKP = JumlahBKP
  rsSales!DPP = JumlahDPP
  rsSales!PPN = JumlahPPN
  rsSales!JumDisc = TotalDiscBrg
  rsSales!SubTotal = Round(SubTotal) '- TotalDiscBrg ' - JumDiscInternRp
  rsSales!JumDiscInternRp = JumDiscInternRp
  rsSales!DiscNota = Disc
  rsSales!Pembulatan = PotonganPembulatan
  rsSales!DiscIntern = DiscINTERNNOTA
  rsSales!Hargatotal = Round(Total) + Round(BiayaCC)
  
  rsSales!UangMuka = Min(Dibayar, Round(Total)) + Round(BiayaCC)
  rsSales!Dibayar = Dibayar
  rsSales!IDUser = IDUser
  rsSales!Bank = Bank
  rsSales!IDBank = IDBankServer
  rsSales!IDJenisKartu = IDJenisKartu
  rsSales!NoAcc = NoAcc
  rsSales!KodeBank = KodeBank
  rsSales!NamaBank = NamaBank
  rsSales!NamaJenisKartu = NamaJenisKartu

  rsSales!Charge = ChargeBank
  rsSales!idcustomer = IDBank 'sementara pakai kolom IDCustomer, tapi ini artinya adalah NoID dari MBank (bukan IDBank yang server)
  rsSales!IDMember = IDMember
  rsSales!Voucher = Voucher
  rsSales!Shift = NamaShift
  rsSales!BarangPoin = BelanjaPoin
  rsSales!SisaPoin = BelanjaPoin - PoinNotaIni * defBelanjadapat1Poin
  rsSales!NilaiPoin = PoinNotaIni
  rsSales!Sopir = Agen
  rsSales!Komisi = KomisiProsen
  rsSales!KomisiRp = KomisiRp
  
  rsSales!IDReedemPoin = IDReedemPoin
  rsSales!ReedemPoin = ReedemPoin
  rsSales!NilaiReedemPoin = ReedemNilai
  
  rsSales.Update
Else
  les = rsSales!UangMuka
  rsSales.Edit
  rsSales!NoID = IDSales
'  rsSales!kode = NoNota'Format(IDSales, "0######")
  rsSales!Tanggal = Now
  rsSales!TotalBKP = JumlahBKP
  rsSales!DPP = JumlahDPP
  rsSales!PPN = JumlahPPN
  rsSales!JumDisc = TotalDiscBrg
  rsSales!SubTotal = Round(SubTotal) '- TotalDiscBrg ' - JumDiscInternRp
  rsSales!JumDiscInternRp = JumDiscInternRp
  rsSales!DiscNota = Disc
  rsSales!Pembulatan = PotonganPembulatan
  rsSales!DiscIntern = DiscINTERNNOTA
  rsSales!Hargatotal = Round(Total) + Round(BiayaCC)
  
  rsSales!UangMuka = Min(Dibayar, Round(Total)) + Round(BiayaCC) 'Min(Total + les, SubTotal)
  rsSales!Dibayar = Dibayar
  rsSales!IDUser = IDUser
  rsSales!Bank = Bank
  rsSales!IDBank = IDBankServer
  rsSales!IDJenisKartu = IDJenisKartu
  rsSales!NoAcc = NoAcc
  rsSales!KodeBank = KodeBank
  rsSales!NamaBank = NamaBank
  rsSales!NamaJenisKartu = NamaJenisKartu

  rsSales!Charge = ChargeBank
  rsSales!idcustomer = IDBank
  rsSales!IDMember = IDMember
  rsSales!Voucher = Voucher
  rsSales!Shift = NamaShift
  rsSales!BarangPoin = BelanjaPoin
  rsSales!SisaPoin = BelanjaPoin - PoinNotaIni * defBelanjadapat1Poin
  rsSales!NilaiPoin = PoinNotaIni
  rsSales!Sopir = Agen
  rsSales!Komisi = KomisiProsen
  rsSales!KomisiRp = KomisiRp
  
  rsSales!IDReedemPoin = IDReedemPoin
  rsSales!ReedemPoin = ReedemPoin
  rsSales!NilaiReedemPoin = ReedemNilai

  rsSales.Update
End If
rsSales.Bookmark = rsSales.LastModified
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True

Trace:
  If Err.Number <> 0 Then
    MsgBox Err.Description
    Err.Clear
  End If
End Sub
Sub BuatBaru()
Dim DB As Database
Dim rst As Recordset
IDMember = -1
NamaMember = ""
defDiscMember = 0
defMinMemberdapatDisc = 0
defDiscMemberBolehInput = False
KodeMember = ""
TipeHargaJual = 0
Agen = ""
KomisiRp = 0
KomisiProsen = 0
lbNama.Caption = ""
cmdReedemPoin.Visible = False
 IDSales = PakaiIDKosong("MSales", Date)
  If IDSales = -1 Then
    IDSales = GetNewID("MSales")
    NoNota = GetNewNota("MSales", Date, "ddMMyyyy")
  Else
    CekDataSales
  End If

'   Data1.RecordSource = "SELECT MSalesD.HargaBruto,MSalesD.DiscRp,MSalesD.DiscProsen,   MSalesD.NoID, MSalesD.IDSales, MSalesD.Transaksi, MSalesD.IDInventor, MSalesD.Qty, MSalesD.Harga, MSalesD.KodeInv, MSalesD.NamaInv, MSalesD.Satuan, MSalesD.Barcode, MSalesD.IDSatuan, MSalesD.Konversi, MSalesD.HargaPokok, MSalesD.IsSend, (MSalesD.Harga*MSalesD.Qty) as Total, (MSalesD.HargaBruto*MSalesD.Qty) as TotalBruto   " & _
'                        "FROM MSalesD WHERE MSalesD.IDSales=" & IDSales & " ORDER BY MSALESD.NOID"
'    Data1.RecordSource = "SELECT MSalesD.HargaBruto,MSalesD.DiscRp,MSalesD.DiscProsen,MSalesD.DiscInternRp,MSalesD.DiscInternProsen," & _
'                        "MSalesD.DiscRp+MSalesD.DiscInternRp as JumDiscRp,MSalesD.DiscProsen+MSalesD.DiscInternProsen as JumDiscProsen," & _
'                        "MSalesD.NoID, MSalesD.IDSales, MSalesD.Transaksi, MSalesD.IDInventor, MSalesD.Qty, MSalesD.Harga, MSalesD.KodeInv," & _
'                        "MSalesD.NamaInv, MSalesD.Satuan, MSalesD.Barcode, MSalesD.IDSatuan, MSalesD.Konversi, MSalesD.HargaPokok, MSalesD.IsSend," & _
'                        "(MSalesD.Harga*MSalesD.Qty) as Total, (MSalesD.HargaBruto*MSalesD.Qty) as TotalBruto,MSalesD.IsMember   " & _
'                        "FROM MSalesD WHERE MSalesD.IDSales=" & IDSales & " ORDER BY MSALESD.NOID"
    Data1.RecordSource = "SELECT MSalesD.*,MSalesD.Qty*MSalesD.Harga as JumlahNetto  " & _
                        " FROM MSalesD WHERE MSalesD.IDSales=" & IDSales & " ORDER BY MSalesD.NOID"

  Data1.Refresh
  IDBank = 0
  IDJenisKartu = 0
  IDBankServer = 0
  NamaBank = ""
  KodeBank = ""
  NamaJenisKartu = ""
  NoAcc = ""
  BelanjaPoin = 0
  Voucher = 0
  Dibayar = 0
  Disc = 0
  JumDiscInternRp = 0
  DiscINTERNNOTA = 0
  SubTotal = 0
  Total = 0
  Bank = 0
  ISCreditCard = False
  lbPDP.Visible = False
  BiayaCC = 0
  TotalDiscBrg = 0
  SaldoHutang = 0
  RoundingBawah = 0
  GetPoinMember
  TampilBawah
  If Dir(DirDatabase & "\TempDB" & Format(Now, "_yyyyMM") & ".mdb") = "" Then
    dbSale.Close
    FileCopy DirDatabase & "\TempDB.mdb", DirDatabase & "\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
  End If
  If Dir(DirUpdate & "\dbmaster.mdb") <> "" Then
    dbs.Close
    Set dbs = Nothing
'    FileCopy DirUpdate & "\DBMaster.mdb", DirDatabase & "\DBMaster.mdb"
    frmProses.Show 1
    If isOnline = False Then
      Set dbs = OpenDatabase(DirDatabase & "\DbMaster.mdb")
    Else
      Set dbs = OpenDatabase(DirDbServer & "\DbMaster.mdb")
    End If
    Set dbSale = OpenDatabase(DirDatabase & "\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
    Set rs = dbs.OpenRecordset("Tinv", dbOpenTable)
    rs.Index = "Kode"
  End If
 Qty = 1
 lbQty = Format(Qty, "##0") & " X"
'  DiscProsenBawah = 0
'  DiscRupiahBawah = 0
End Sub


Sub Pending()
Dim rsSales As Recordset
Set rsSales = dbSale.OpenRecordset("Select * FROM MSales Where NoID=" & IDSales & " ORDer by NOID")
If rsSales.EOF And rsSales.BOF Then
  rsSales.AddNew
'  rsSales!NoID = IDSales
'  rsSales!kode = NoNota 'Format(IDSales, "0######")
'  rsSales!tanggal = Now
'  rsSales!TotalPajak = 0
'  rsSales!JumDisc = TotalDiscBrg
'  rsSales!SubTotal = Round(SubTotal) - TotalDiscBrg
'  rsSales!JumDiscInternRp = JumDiscInternRp
'  rsSales!DiscNota = Disc
'  rsSales!Hargatotal = Round(Total) + Round(BiayaCC)
'  rsSales!IDPayment = 0
'  rsSales!UangMuka = Min(Dibayar, Round(Total)) + Round(BiayaCC)
'  rsSales!IDUser = IDUser
'  rsSales!Bank = Bank
'  rsSales!IDBank = IDBankServer
'  rsSales!NoAcc = NoAcc
'  rsSales!idcustomer = IDBank
'  rsSales!Voucher = Voucher
'  rsSales!Shift = NamaShift
'  rsSales!ispending = True
  rsSales.Update
Else
UpdateSales
  rsSales.Edit
'  rsSales!NoID = IDSales
'  rsSales!kode = Format(IDSales, "0######")
'  rsSales!tanggal = Now
'  rsSales!TotalPajak = 0
'  rsSales!JumDisc = TotalDiscBrg
'  rsSales!SubTotal = Round(SubTotal) - TotalDiscBrg
'  rsSales!JumDiscInternRp = JumDiscInternRp
'  rsSales!DiscNota = Disc
'  rsSales!Hargatotal = Round(Total) + Round(BiayaCC)
'  rsSales!IDPayment = 0
'  rsSales!UangMuka = Min(Dibayar, Round(Total)) + Round(BiayaCC)
'  rsSales!IDUser = IDUser
'  rsSales!Bank = Bank
'  rsSales!IDBank = IDBankServer
'  rsSales!NoAcc = NoAcc
'  rsSales!idcustomer = IDBank
'  rsSales!Voucher = Voucher
'  rsSales!Shift = NamaShift
  rsSales!ispending = True
  rsSales.Update
rsSales.Bookmark = rsSales.LastModified
End If
DisplayPesan "TRANSAKSI DI PENDING", "                     "
Prin "---------------------------------------" & Chr(13) & Chr(10) & _
                "Subtotal" & Space(31 - Len(Format(SubTotal, "###,###,##0"))) & Format(SubTotal, "###,###,##0") & Chr(13) & Chr(10) & _
                IIf(TotalDiscBrg + JumDiscInternRp = 0, "", "Discount Barang" & Space(24 - Len(Format(TotalDiscBrg + JumDiscInternRp, "###,###,##0"))) & Format(TotalDiscBrg + JumDiscInternRp, "###,###,##0") & Chr(13) & Chr(10)) & _
                IIf(Disc + DiscINTERNNOTA + PotonganPembulatan = 0, "", "Discount Nota" & Space(26 - Len(Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0"))) & Format(Disc + DiscINTERNNOTA + PotonganPembulatan, "###,###,##0") & Chr(13) & Chr(10)) & _
              "Voucher " & Space(31 - Len(Format(Voucher, "###,###,##0"))) & Format(Voucher, "###,###,##0") & Chr(13) & Chr(10) & _
              "Dibayar " & Space(31 - Len(Format(Dibayar - Voucher, "###,###,##0"))) & Format(Dibayar - Voucher, "###,###,##0") & Chr(13) & Chr(10) & _
              "Kembali " & Space(31 - Len(Format(Kembali, "###,###,##0"))) & Format(Kembali, "###,###,##0") & Chr(13) & Chr(10) & _
              IIf(ISSembunyikanFooterbarangKSB, "", "Belanja Barang POIN" & Space(20 - Len(Format(BelanjaPoin, "###,###,##0"))) & Format(BelanjaPoin, "###,###,##0") & Chr(13) & Chr(10)) & _
              "#" & Format(CLng(NoNota), "0#####") & "," & "Kasir" & NamaMesin & ":" & Mid(NamaKasir, 1, 6) & Space(6 - IIf(Len(NamaKasir) <= 6, Len(NamaKasir), 6)) & Format(Date, " dd/MM/yy") & Format(Time, "hh:mm:ss") & Chr(13) & Chr(10) & _
              "Macam Barang =" & Format(MacamItem, "##0") & " , Total Item =" & Format(Jumlahitem, "##0") & Chr(13) & Chr(10) & _
              IIf(KodeMember = "", "", "CUSTOMER: " & KodeMember & Chr(13) & Chr(10)) & _
              "----TERIMA KASIH ATAS KUNJUNGAN ANDA----" & Chr(13) & Chr(10) & _
              IIf(Footer1 = "", "", Footer1 & Chr(13) & Chr(10)) & _
              IIf(Footer2 = "", "", Footer2 & Chr(13) & Chr(10)) & _
              IIf(Footer3 = "", "", Footer3 & Chr(13) & Chr(10)) & _
              Chr(13) & Chr(10)
'        PrinBigChar Space(15 - Min(15, Len(NamaToko))) & UCase(NamaToko)
'        Prin Judulstruk & Chr(13) & Chr(10) & "---------------------------------------"
'        papercut

Prin "----------TRANSAKSI DI PENDING--------" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10)
    If IsHematKertas Then
        Prin Judulstruk
        Prin "---------------------------------------" & Chr(13) & Chr(10)
    End If
papercut
Ditutup = True
'CetakStruckPending IDSales, False
cmdHOME_Click

End Sub



Sub CekLastEdit()
Dim rss As Recordset
'Set rss = dbSale.OpenRecordset("Select MSales.* FRom MSalesD INNER JOIN MSales ON MSales.NoID=MSalesD.IDSales where MSales.NoID=" & IDSales)
Set rss = dbSale.OpenRecordset("Select MSales.* FRom MSales where MSales.NoID=" & IDSales)
If rss.BOF And rss.EOF Then
  IDSales = GetNewID("MSALES")
  NoNota = GetNewNota("MSales", Date, "ddMMyyyy")
Else
  If IsNull(rss!Hargatotal) Or IsNull(rss!UangMuka) Then
  Else
    'If ((rss!UangMuka + rss!Hutang) - rss!Hargatotal >= 0) And rss!Hargatotal > 0 Then
    If (rss!UangMuka > 0) And rss!Hargatotal > 0 And NullToBool(rss!IsSelesai) = vbTrue Then
    'If (rss!UangMuka > 0) And rss!Hargatotal > 0 Then
       'IDSales = IDSales + 1 'Jika ambil dari pendingan dg IDSales lama akan kacau
       '4 baris berikut di revisi per 25 Agustus 2005
      IDSales = GetNewID("MSALES")
      NoNota = GetNewNota("MSales", Date, "ddMMyyyy")
'       Data1.RecordSource = "SELECT MSalesD.HargaBruto,MSalesD.DiscRp,MSalesD.DiscProsen,   MSalesD.NoID, MSalesD.IDSales, MSalesD.Transaksi, MSalesD.IDInventor, MSalesD.Qty, MSalesD.Harga, MSalesD.KodeInv, MSalesD.NamaInv, MSalesD.Satuan, MSalesD.Barcode, MSalesD.IDSatuan, MSalesD.Konversi, MSalesD.HargaPokok, MSalesD.IsSend, (MSalesD.Harga*MSalesD.Qty) as Total, (MSalesD.HargaBruto*MSalesD.Qty) as TotalBruto   " & _
'                        "FROM MSalesD WHERE MSalesD.IDSales=" & IDSales & " ORDER BY MSALESD.NOID"
' Data1.RecordSource = "SELECT MSalesD.HargaBruto,MSalesD.DiscRp,MSalesD.DiscProsen,   MSalesD.NoID, MSalesD.IDSales, MSalesD.Transaksi, MSalesD.IDInventor, MSalesD.Qty, MSalesD.Harga, MSalesD.KodeInv, MSalesD.NamaInv, MSalesD.Satuan, MSalesD.Barcode, MSalesD.IDSatuan, MSalesD.Konversi, MSalesD.HargaPokok, MSalesD.IsSend, (MSalesD.Harga*MSalesD.Qty) as Total, (MSalesD.HargaBruto*MSalesD.Qty) as TotalBruto   " & _
'                        "FROM MSalesD WHERE MSalesD.IDSales=" & IDSales & " ORDER BY MSALESD.NOID"
'Data1.RecordSource = "SELECT MSalesD.HargaBruto,MSalesD.DiscRp,MSalesD.DiscProsen,MSalesD.DiscInternRp,MSalesD.DiscInternProsen," & _
'                        "MSalesD.DiscRp+MSalesD.DiscInternRp as JumDiscRp,MSalesD.DiscProsen+MSalesD.DiscInternProsen as JumDiscProsen," & _
'                        "MSalesD.NoID, MSalesD.IDSales, MSalesD.Transaksi, MSalesD.IDInventor, MSalesD.Qty, MSalesD.Harga, MSalesD.KodeInv, MSalesD.NamaInv, MSalesD.Satuan, MSalesD.Barcode, MSalesD.IDSatuan, MSalesD.Konversi, MSalesD.HargaPokok, MSalesD.IsSend, (MSalesD.Harga*MSalesD.Qty) as Total, (MSalesD.HargaBruto*MSalesD.Qty) as TotalBruto,MSalesD.IsMember   " & _
'                        "FROM MSalesD WHERE MSalesD.IDSales=" & IDSales & " ORDER BY MSALESD.NOID"
       Data1.RecordSource = "SELECT MSalesD.*,MSalesD.Qty*MSalesD.Harga as JumlahNetto  " & _
                            " FROM MSalesD WHERE MSalesD.IDSales=" & IDSales & " ORDER BY MSalesD.NOID"


       Data1.Refresh
       Agen = ""
       KomisiRp = 0
       KomisiProsen = 0
       Dibayar = 0
       IDBank = 0
       IDJenisKartu = 0
       IDBankServer = 0
       SubTotal = 0
       TotalDiscBrg = 0
       JumDiscInternRp = 0
       Total = 0
       Voucher = 0
       Disc = 0
       DiscINTERNNOTA = 0
        NoAcc = ""
        KodeMember = ""
        IDMember = -1
        NamaMember = ""
        defDiscMember = 0
        defMinMemberdapatDisc = 0
        defDiscMemberBolehInput = False
       PrinBigChar Space(15 - Min(15, Len(NamaToko))) & UCase(NamaToko)
       Prin Judulstruk & Chr(13) & Chr(10) & "---------------------------------------"
       Ditutup = True
    Else
       GetIDCUSTOMERONLINE NullToStr(rss!KodeMember)
       SubTotal = Round(rss!SubTotal, 0) 'Dibulatkan, karena di kasir koma tdk terlihat
       JumDiscInternRp = Round(rss!JumDiscInternRp, 0)
       DiscINTERNNOTA = Round(rss!DiscIntern, 0)
       Disc = rss!DiscNota
       NoNota = rss!kode
       Total = rss!Hargatotal
       Dibayar = rss!UangMuka
       Voucher = rss!Voucher
       Kembali = Dibayar - Total - Voucher
       Ditutup = False
        'Timer2.Enabled = False
        'bolehbergerak = False
    End If
  End If
End If
rss.Close
HitungSubTotal
End Sub
Sub CekDataSales()
Dim rss As Recordset
Set rss = dbSale.OpenRecordset("Select * FRom Msales where NoID=" & IDSales)
If rss.BOF And rss.EOF Then
Else
       IDReedemPoin = NullToNol(rss!IDReedemPoin)
       ReedemPoin = NullToNol(rss!ReedemPoin)
       ReedemNilai = NullToNol(rss!NilaiReedemPoin)
       Kembali = Dibayar - Total - Voucher
       
       GetIDCUSTOMERONLINE NullToStr(rss!KodeMember)
       NoNota = NullToStr(rss!kode)
       SubTotal = Round(rss!SubTotal, 0) 'Dibulatkan, karena di kasir koma tdk terlihat
       Disc = rss!DiscNota
       Total = rss!Hargatotal
       Dibayar = rss!UangMuka
       Voucher = rss!Voucher
       Agen = NullToStr(rss!Sopir)
       KomisiRp = NullToNol(rss!KomisiRp)
       KomisiProsen = NullToNol(rss!Komisi)
       Ditutup = False
 End If
rss.Close
Set rss = Nothing
End Sub
Public Sub GetIDCUSTOMERONLINE(ByVal KDCust As String)
  GetIDCUSTOMERLOKAL KDCust
End Sub
Public Sub GetIDCUSTOMERLOKAL(ByVal KodeCust As String)
On Error GoTo Trace
Dim rss As Recordset
Dim dbEmp As Database
Set dbEmp = OpenDatabase(DirDatabase & "\DBMaster.mdb")
Set rss = dbEmp.OpenRecordset("Select NoID,Kode,Nama,DiscountCustomer,SyaratMinimum,AllowInputDiscount from MCustomer where Nama='" & Replace(KodeCust, "'", "''") & "' OR Kode='" & Replace(KodeCust, "'", "''") & "' OR Barcode='" & Replace(KodeCust, "'", "''") & "'")
If rss.BOF And rss.EOF Then
  NamaMember = ""
  KodeMember = ""
  IDMember = -1
  LimitHutang = 0
  defDiscMember = 0
  defMinMemberdapatDisc = 0
  defDiscMemberBolehInput = False
Else
  NamaMember = NullToStr(rss!Nama)
  KodeMember = NullToStr(rss!kode)
'  LimitHutang = NullToNol(rss!LimitHutang)
  IDMember = rss!NoID
'  TipeHargaJual = NullToNol(rss!TipeHargaJual)
  
  defDiscMember = NullToNol(rss!DiscountCustomer)
  defMinMemberdapatDisc = NullToNol(rss!SyaratMinimum)
  defDiscMemberBolehInput = NullToStr(rss!AllowInputDiscount)
  rss.Close
  Set rss = Nothing
  dbEmp.Close
  Set dbEmp = Nothing
End If
lbNama.Caption = NamaMember
GetPoinMember

Trace:
If Err.Number <> 0 Then
  MsgBox "Error : " & Err.Number & vbCrLf & Err.Description, vbCritical, App.Title
  Err.Clear
End If
End Sub
Private Sub Timer2_Timer()
  'lbJam = Format(Time, "HH : mm : ss")
'posisi = ((posisi + 1) Mod 20)
 '   If bolehbergerak Then
 '       DisplayPesan Mid(NamaToko & "  " & NamaToko & "  " & "  " & NamaToko & "  ", 1 + posisi, 20), Mid("KASSA : " & NamaMesin & " , ^" & NamaKasir & "^   " & "KASSA : " & NamaMesin & " , ^" & NamaKasir & "^   " & "KASSA : " & NamaMesin & " , ^" & NamaKasir & "^   ", 1 + posisi, 20)
 '   End If
'Timer2.Enabled = False
     '   Text1.SetFocus
     
lbJam = Format(Time, "HH:mm:ss")
' Timer1.Enabled = False
    
    If bolehbergerak Then
    If posisi <= 0 Then
      skala2 = 1
    ElseIf posisi >= 6 Then
      skala2 = -1
    End If
      posisi = posisi + skala2
       ' DisplayPesan Mid(NamaToko & "  " & NamaToko & "  " & "  " & NamaToko & "  ", 1 + posisi, 20), Mid("KASSA : " & NamaMesin & " , ^" & NamaKasir & "^   " & "KASSA : " & NamaMesin & " , ^" & NamaKasir & "^   " & "KASSA : " & NamaMesin & " , ^" & NamaKasir & "^   ", 1 + posisi, 20)
       DisplayPesan Space(posisi) & "SELAMAT DATANG", "                    "
    End If
    If dxLabel1.Left <= 1400 Then
      skala1 = 50
    ElseIf dxLabel1.Left >= 2600 Then
     skala1 = -50
    End If
    dxLabel1.Move dxLabel1.Left + skala1, dxLabel1.Top
    dxLabel2.Move dxLabel2.Left + skala1, dxLabel2.Top
End Sub
Sub HitungItem()
Dim rsitem As Recordset
Set rsitem = dbSale.OpenRecordset("SELECT Count(MSalesD.IDInventor) AS MacamItem, Sum(MSalesD.Qty) AS JumItem From MSalesD Where (((MSalesD.IDSales) = " & IDSales & "))")
If rsitem.EOF And rsitem.BOF Then
    MacamItem = 0
    Jumlahitem = 0
Else
    If IsNull(rsitem!MacamItem) Then
        MacamItem = 0
    Else
        MacamItem = rsitem!MacamItem
    End If
    If IsNull(rsitem!Jumitem) Then
        Jumlahitem = 0
    Else
        Jumlahitem = rsitem!Jumitem
    End If
End If
End Sub

'Private Sub Timer3_Timer()
''lbStatus = "Status  : " & GetStatusNetwork
'' Text1.SetFocus
'End Sub

Sub BukaCommBarcode()
On Error Resume Next
BarcodeIn = ""
MSComm1.PortOpen = True
End Sub

Sub TutupCommBarcode()
On Error Resume Next
MSComm1.PortOpen = False
End Sub

Function GetLastTime() As Long
Dim dbs As Database
Dim rs As Recordset
Set dbs = OpenDatabase(App.path & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
Set rs = dbs.OpenRecordset("select Top 1 NOID from MSAles where Day(tanggal)=" & Day(Date) & " and Month(tanggal)=" & Month(Date) & " and year(tanggal)=" & Year(Date) & " and IIF(IsNull(IsSelesai),0,IsSelesai)=0 Order by Tanggal DESC")
If rs.EOF And rs.BOF Then
GetLastTime = 1
Else
GetLastTime = rs!NoID
End If
End Function
Public Function JumlahBolehHutang(ByVal LimitHutang As Double, ByVal idcustomer As Long) As Double
    Dim JumBayar As Double
    Dim JumBelanja As Double
    JumBelanja = NullToNol(ExecuteSkalarSQL("Select Sum(Total-UangMuka-diskonnota) as Hasil From Mjual where IDSupplier=" & idcustomer))
    JumBayar = NullToNol(ExecuteSkalarSQL("Select Sum(MBayarHutangD.Bayar+MBayarHutangD.Retur) as Hasil From MBayarHutangD inner join Mjual On MBayarHutangD.IDBeli=MJual.ID where MJual.IDSupplier=" & idcustomer))
    
    JumlahBolehHutang = LimitHutang - (JumBelanja - JumBayar)
End Function

Private Sub CetakStruckTKP(ByVal IDTKP As Long, ByVal Reprint As String)
'openDrawer
Dim cn As New ADODB.Connection
Dim rst As New ADODB.Recordset
If TipeCetakan = None Then Exit Sub
If TipeCetakan = Optional_ Then
  If MsgBox("Mau Cetak Struk Tukar Poin?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
End If
Dim i As Integer
On Error GoTo Err
DoEvents
cn.ConnectionString = Cnstr
cn.Open
Set rst = cn.Execute("SELECT MTukarPoin.*, MAlamat.Kode AS KdMember, MAlamat.Nama AS NamaMember, MAlamat.Alamat, MUser.Nama AS Kasir " & vbCrLf & _
                     " From MTukarPoin " & vbCrLf & _
                     " LEFT JOIN MAlamat ON MAlamat.NoID=MTukarPoin.IDMember " & vbCrLf & _
                     " LEFT JOIN MUser ON MUser.NoID=MTukarPoin.IDKasir WHERE MTukarPoin.NoID=" & IDTKP)
If Not (rst.BOF Or rst.EOF) Then
  CrystalReport1.Reset
  CrystalReport1.ReportFileName = App.path & "\Report\TukarPoin.rpt"
  CrystalReport1.Formulas(0) = "NoID=" & IDTKP
  CrystalReport1.Formulas(1) = "NamaKasir='" & NullToStr(rst!Kasir) & "'"
  CrystalReport1.Formulas(2) = "NamaPerusahaan='" & Trim(NamaToko) & "'"
  CrystalReport1.Formulas(3) = "AlamatPerusahaan='" & Trim(Replace(Judulstruk, vbCrLf, "'+ Chr(13) +'")) & "'"
  CrystalReport1.Formulas(4) = "Kassa='" & Format(IDPOSDef, "##") & "'"
  CrystalReport1.Formulas(5) = "NoMember='" & Trim(NullToStr(rst!kdMember)) & "'"
  CrystalReport1.Formulas(6) = "NamaMember='" & Trim(NullToStr(rst!NamaMember)) & "'"
  CrystalReport1.Formulas(7) = "AlamatMember='" & Trim(NullToStr(rst!Alamat)) & "'"
  CrystalReport1.Formulas(8) = "Poin='" & Trim(NullToNol(rst!JumlahPoin)) & "'"
  CrystalReport1.Formulas(9) = "TukarPoin='" & Trim(NullToNol(rst!Kredit)) & "'"
  CrystalReport1.Formulas(10) = "SisaPoin='" & Trim(NullToNol(rst!Saldo)) & "'"
  CrystalReport1.Formulas(11) = "Catatan='" & Trim(NullToStr(rst!Keterangan)) & "'"
  CrystalReport1.Formulas(12) = "Reprint='" & Trim(Reprint) & "'"
  CrystalReport1.WindowShowPrintSetupBtn = True
  CrystalReport1.WindowShowRefreshBtn = True
  CrystalReport1.WindowState = crptMaximized
  If TipeCetakan = Preview_ Then
    CrystalReport1.Destination = crptToWindow
  Else
    CrystalReport1.Destination = crptToPrinter
    CrystalReport1.ProgressDialog = False
  End If
  CrystalReport1.Action = 1
End If
Err:
If Err.Number <> 0 Then
  MsgBox "Error : " & Err.Number & "-" & Err.Description, vbCritical, App.Title
  Err.Clear
End If
If cn.State = adStateOpen Then
  cn.Close
End If
Set cn = Nothing
If rst.State = adStateOpen Then
  rst.Close
End If
Set rst = Nothing
End Sub

Private Sub CetakStruckReedem2(ByVal IDSales As Long, ByVal Reprint As Boolean)
'openDrawer
Dim dbs As Database
Dim rs As Recordset
Set dbs = OpenDatabase(App.path & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
Set rs = dbs.OpenRecordset("select Top 1 * from MSales where NoID=" & IDSales)
If rs.EOF And rs.BOF Then
Else
  On Error GoTo Err
  Dim iRedemPoin As Double, SQL As String
  Dim TotalPoin As Double
  Dim Tgl As Date
  Dim cn As New ADODB.Connection
  Dim com As New ADODB.Command
  Dim rst As New ADODB.Recordset
  If TipeCetakan = None Then Exit Sub
  
  iRedemPoin = NullToNol(rs!ReedemPoin)
  Tgl = NullToDate(rs!Tanggal)
  If iRedemPoin >= 1 Then
    'If TipeCetakan = Optional_ Then
      If MsgBox("Mau Cetak Struk Reedem Poin?", vbQuestion + vbYesNo) = vbNo Then GoTo grr
    'End If
    DoEvents
    cn.ConnectionString = Cnstr
    cn.Open
'    'Set com.ActiveConnection = cn
'    With com
'      .CommandType = adCmdStoredProc
'      .CommandText = "sp_MengambilDataRedeemPoinPenjualan" '& FixKoma(IDTKP) & ", " & IDPOSDef & ", '" & Format(Tanggal, "yyyy-MM-dd") & "'"
'    '  .Parameters.Refresh
'      .Parameters.Append .CreateParameter("NoID", adInteger, adParamInput, , IDSales)
'      .Parameters.Append .CreateParameter("IDKassa", adInteger, adParamInput, , IDPOSDef)
'      .Parameters.Append .CreateParameter("Tanggal", adVarChar, adParamInput, 20, Format(Tgl, "yyyy-MM-dd"))
'      rst.Open "sp_MengambilDataRedeemPoinPenjualan" & IDSales & ", " & IDPOSDef & ", '" & Format(Tgl, "yyyy-MM-dd") & "'", com
'    End With
    SQL = "SELECT SUM(Debet-Kredit) FROM MCustomerPoin (NOLOCK) WHERE YEAR(Tanggal)=YEAR(GETDATE()) AND Tanggal<'" & Format(Tgl, "yyyy-MM-dd HH:mm:ss") & "' AND IDCustomer=" & NullToNol(rs!IDMember)
    Set rst = cn.Execute(SQL)
    If rst.EOF Or rst.BOF Then
      TotalPoin = 0
    Else
      TotalPoin = NullToNol(rst(0).Value)
    End If
    Set rst = Nothing
    SQL = "SELECT A.Tanggal, A.KodeReff Kode, A.IDCustomer IDMember, B.Kode KodeMember," & vbCrLf & _
          "B.Nama NamaMember, B.Alamat AlamatMember, CONVERT(NUMERIC(18,2), " & FixKoma(TotalPoin) & ") Poin," & vbCrLf & _
          "A.ReedemPoin RedeemPoin, A.ReedemNilai NilaiRedeem, CONVERT(NUMERIC(18,2), " & FixKoma(TotalPoin) & ")-ISNULL(A.ReedemPoin,0) SisaPoin, A.NamaKasir Kasir" & vbCrLf & _
          "FROM MJual (NOLOCK) A" & vbCrLf & _
          "LEFT JOIN MAlamat (NOLOCK) B ON A.IDCustomer = B.NoID" & vbCrLf & _
          "WHERE A.NoIDPos = " & IDSales & " " & vbCrLf & _
          "AND A.IDPos = " & IDPOSDef & " " & vbCrLf & _
          "AND (CONVERT(DATE, A.Tanggal) BETWEEN '" & Format(Tgl, "yyyy-MM-dd") & "' AND '" & Format(Tgl, "yyyy-MM-dd") & "')"
    Set rst = cn.Execute(SQL)
    If Not (rst.BOF Or rst.EOF) Then
      CrystalReport1.Reset
      CrystalReport1.ReportFileName = App.path & "\Report\RedeemPoin.rpt"
      CrystalReport1.Formulas(0) = "NoID=" & IDSales
      CrystalReport1.Formulas(1) = "NamaKasir='" & NullToStr(rst!Kasir) & "'"
      CrystalReport1.Formulas(2) = "NamaPerusahaan='" & Trim(NamaToko) & "'"
      CrystalReport1.Formulas(3) = "AlamatPerusahaan='" & Trim(Replace(Judulstruk, vbCrLf, "'+ Chr(13) +'")) & "'"
      CrystalReport1.Formulas(4) = "Kassa='" & Format(IDPOSDef, "##") & "'"
      CrystalReport1.Formulas(5) = "NoMember='" & Trim(NullToStr(rst!KodeMember)) & "'"
      CrystalReport1.Formulas(6) = "NamaMember='" & Trim(NullToStr(rst!NamaMember)) & "'"
      CrystalReport1.Formulas(7) = "AlamatMember='" & Trim(NullToStr(rst!AlamatMember)) & "'"
      CrystalReport1.Formulas(8) = "Poin=" & NullToNol(rst!Poin) + NullToNol(rs!NilaiPoin)
      CrystalReport1.Formulas(9) = "TukarPoin=" & NullToNol(rst!RedeemPoin)
      CrystalReport1.Formulas(10) = "SisaPoin=" & NullToNol(rst!SisaPoin) + NullToNol(rs!NilaiPoin)
      CrystalReport1.Formulas(11) = "Catatan='" & Trim(NullToStr(rst!kode)) & "'"
      CrystalReport1.Formulas(12) = "Reprint='" & IIf(Reprint, "Reprint", "") & "'"
      CrystalReport1.Formulas(13) = "NilaiRedeem=" & NullToNol(rst!NilaiRedeem)
      CrystalReport1.WindowShowPrintSetupBtn = True
      CrystalReport1.WindowShowRefreshBtn = True
      CrystalReport1.WindowState = crptMaximized
      If TipeCetakan = Preview_ Then
        CrystalReport1.Destination = crptToWindow
      Else
        CrystalReport1.Destination = crptToPrinter
        CrystalReport1.ProgressDialog = False
      End If
      CrystalReport1.Action = 1
    End If
Err:
    If Err.Number <> 0 Then
      MsgBox "Error : " & Err.Number & "-" & Err.Description, vbCritical, App.Title
      Err.Clear
    End If
    If cn.State = adStateOpen Then
      cn.Close
    End If
    Set cn = Nothing
    If rst.State = adStateOpen Then
      rst.Close
    End If
    Set com = Nothing
    Set rst = Nothing
  End If
End If

grr:
dbs.Close
Set dbs = Nothing
Set rs = Nothing
End Sub

Private Sub CetakStruck(ByVal NoID As Long, ByVal IsRePrint As Boolean, Optional ByVal ReprintKe As Integer = 0, Optional ByVal IsAllVoid As Boolean = False)
'openDrawer

If TipeCetakan = None Then Exit Sub
If TipeCetakan = Optional_ Then
  If MsgBox("Mau Cetak Struk", vbQuestion + vbYesNo) = vbNo Then Exit Sub
End If
Dim i As Integer
Dim dbstmp As Database
Dim dbMaster As Database
Dim rst As Recordset
'On Error GoTo Err
Set dbstmp = OpenDatabase(App.path & "\database\tempdb" & Format(Now, "_yyyyMM") & ".mdb")
Set dbMaster = OpenDatabase(App.path & "\database\DBMaster.mdb")
'dbstmp.Execute "Delete * From MCetakSales"
'dbstmp.Execute "Insert into MCetakSales(NoID) Values(" & NoID & ")"
DoEvents
Set rst = dbstmp.OpenRecordset("SELECT * FROM MSales WHERE NoID=" & NoID)
  If Not (rst.BOF Or rst.EOF) Then
    Set rst = dbMaster.OpenRecordset("SELECT * FROM MCustomer WHERE NoID=" & (rst!IDMember))
CrystalReport1.Reset
' If NullToNol(rst!TipeHargaJual) = 0 Then
      CrystalReport1.ReportFileName = App.path & "\Report\STRUCK.rpt"
    'ElseIf NullToNol(rst!TipeHargaJual) = 2 Then
     ' CrystalReport1.ReportFileName = App.Path & "\Report\STRUCK1.rpt"
    'Else
     ' CrystalReport1.ReportFileName = App.Path & "\Report\STRUCK.rpt"
    'End If
    CrystalReport1.DataFiles(0) = App.path & "\Database\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
    CrystalReport1.DataFiles(1) = App.path & "\Database\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
    
    'CrystalReport1.re .DataFiles(2) = App.path & "\Database\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
'    CrystalReport1.DataFiles(2) = App.path & "\Database\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
 
'    CrystalReport1.SelectionFormula = "{MSales.NoID}=" & NoID
    CrystalReport1.Formulas(1) = "IDSales=" & NoID
    CrystalReport1.Formulas(2) = "NamaKasir='" & NamaKasir & "'"
    CrystalReport1.Formulas(3) = "NamaPerusahaan='" & Trim(NamaToko) & "'"
    CrystalReport1.Formulas(4) = "AlamatPerusahaan='" & Trim(Replace(Judulstruk, vbCrLf, "'+ Chr(13) +'")) & "'"
    CrystalReport1.Formulas(5) = "Kassa='" & Format(IDPOSDef, "##") & "'"
    If IsRePrint Then
        If IsAllVoid Then
          CrystalReport1.Formulas(6) = "RePrint='COPY (ALL VOID)'"
        Else
          CrystalReport1.Formulas(6) = "RePrint='COPY " & IIf(ReprintKe >= 2, ReprintKe, "") & "'"
        End If
    Else
      If IsAllVoid Then
          CrystalReport1.Formulas(6) = "RePrint='ALL-VOID'"
        Else
          CrystalReport1.Formulas(6) = "RePrint=''"
        End If
    End If
    
    If IsNotaDariPending Then
      CrystalReport1.Formulas(7) = "IsPending=True"
    Else
      CrystalReport1.Formulas(7) = "IsPending=False"
    End If
    CrystalReport1.SubreportToChange = "Voucher"
    CrystalReport1.DataFiles(0) = App.path & "\Database\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
    
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowState = crptMaximized
    If TipeCetakan = Preview_ Then
      CrystalReport1.Destination = crptToWindow
    Else
      CrystalReport1.Destination = crptToPrinter
      CrystalReport1.ProgressDialog = False
    End If
    CrystalReport1.Action = 1
    dbstmp.Execute "Update MSales Set IsSelesai=-1 where NoID=" & NoID
  End If
  rst.Close
  Set rst = Nothing
  dbstmp.Close
  Set dbstmp = Nothing
  dbMaster.Close
  Set dbMaster = Nothing
  Exit Sub
Err:
If Err.Number <> 0 Then
  MsgBox "Error : " & Err.Number & "-" & Err.Description, vbCritical, App.Title
  Err.Clear
End If
End Sub

Sub openDrawer()
If NoPortDrawer <> -2 Then
openDrawerbyDos
Else


On Error GoTo Err
'Set rst = dbstmp.OpenRecordset("SELECT * FROM MSales WHERE NoID=" & NoID)
'    Set rst = dbMaster.OpenRecordset("SELECT * FROM MCustomer WHERE NoID=" & (rst!IDMember))
'    If NullToNol(rst!TipeHargaJual) = 0 Then
      CrystalReport1.ReportFileName = App.path & "\Report\DRAWER.rpt"
'    ElseIf NullToNol(rst!TipeHargaJual) = 2 Then
'      CrystalReport1.ReportFileName = App.Path & "\Report\STRUCK1.rpt"
'    Else
'      CrystalReport1.ReportFileName = App.Path & "\Report\STRUCK.rpt"
'    End If
    CrystalReport1.DataFiles(0) = App.path & "\Database\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
    CrystalReport1.DataFiles(1) = App.path & "\Database\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
    '    CrystalReport1.WindowShowPrintSetupBtn = True
'    CrystalReport1.WindowShowRefreshBtn = True
'    CrystalReport1.WindowState = crptMaximized
'    If TipeCetakan = Preview_ Then
'      CrystalReport1.Destination = crptToWindow
'    Else
      CrystalReport1.Destination = crptToPrinter
'    End If
    CrystalReport1.Action = 1
End If
  Exit Sub
Err:
If Err.Number <> 0 Then
  MsgBox "Error : " & Err.Number & "-" & Err.Description, vbCritical, App.Title
  Err.Clear
End If

End Sub
Private Sub CetakStruckPending(ByVal NoID As Long, ByVal IsRePrint As Boolean)
If TipeCetakan = None Then Exit Sub
Dim i As Integer
On Error GoTo Err
  CrystalReport1.ReportFileName = App.path & "\Report\STRUCKPending.rpt"
  CrystalReport1.DataFiles(0) = App.path & "\Database\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
  CrystalReport1.DataFiles(1) = App.path & "\Database\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
  'CrystalReport1.DataFiles(2) = App.Path & "\Database\TempDB"& format(now,"_yyyyMM") &".mdb"
  CrystalReport1.Formulas(1) = "IDSales=" & NoID
  CrystalReport1.Formulas(2) = "NamaKasir='" & NamaKasir & "'"
  CrystalReport1.Formulas(3) = "NamaPerusahaan='" & Trim(NamaToko) & "'"
  CrystalReport1.Formulas(4) = "AlamatPerusahaan='" & Trim(Replace(Judulstruk, vbCrLf, "'+ Chr(13) +'")) & "'"
  If IsRePrint Then
    CrystalReport1.Formulas(5) = "RePrint='RE-PRINT'"
  Else
    CrystalReport1.Formulas(5) = "RePrint=''"
  End If
  CrystalReport1.WindowShowPrintSetupBtn = True
  CrystalReport1.WindowShowRefreshBtn = True
  CrystalReport1.WindowState = crptMaximized
  If TipeCetakan = Preview_ Then
    CrystalReport1.Destination = crptToWindow
  Else
    CrystalReport1.Destination = crptToPrinter
  End If
  CrystalReport1.Action = 1
  Exit Sub
Err:
If Err.Number <> 0 Then
  MsgBox "Error : " & Err.Number & "-" & Err.Description, vbCritical, App.Title
  Err.Clear
End If
End Sub

Private Sub Timer4_Timer()
'  Dim rst As New ADODB.Recordset
'  dim
End Sub
