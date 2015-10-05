VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmDaftarJual 
   Caption         =   "Daftar Penjualan"
   ClientHeight    =   8025
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14685
   Icon            =   "frmDaftarPenjualan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   14685
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtJumlah 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   11760
      TabIndex        =   12
      Top             =   7200
      Width           =   2655
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   7920
      TabIndex        =   11
      Top             =   600
      Width           =   6495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Tutup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   9
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   8
      Top             =   120
      Width           =   2055
   End
   Begin TDBDate6Ctl.TDBDate TDBDate1 
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   120
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   661
      Calendar        =   "frmDaftarPenjualan.frx":146AA
      Caption         =   "frmDaftarPenjualan.frx":147C2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmDaftarPenjualan.frx":1482E
      Keys            =   "frmDaftarPenjualan.frx":1484C
      Spin            =   "frmDaftarPenjualan.frx":148AA
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd/mm/yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "dd/mm/yyyy"
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
      Text            =   "23/07/2012"
      ValidateMode    =   0
      ValueVT         =   1380253703
      Value           =   41113
      CenturyMode     =   0
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   720
      Width           =   2055
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   735
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6840
      Width           =   2655
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Bindings        =   "frmDaftarPenjualan.frx":148D2
      Height          =   5535
      Left            =   240
      OleObjectBlob   =   "frmDaftarPenjualan.frx":148E6
      TabIndex        =   0
      Top             =   1560
      Width           =   14175
   End
   Begin TDBDate6Ctl.TDBDate TDBDate2 
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   120
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   661
      Calendar        =   "frmDaftarPenjualan.frx":198D0
      Caption         =   "frmDaftarPenjualan.frx":199E8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmDaftarPenjualan.frx":19A54
      Keys            =   "frmDaftarPenjualan.frx":19A72
      Spin            =   "frmDaftarPenjualan.frx":19AD0
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd/mm/yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "dd/mm/yyyy"
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
      Text            =   "23/07/2012"
      ValidateMode    =   0
      ValueVT         =   1380253703
      Value           =   41113
      CenturyMode     =   0
   End
   Begin VB.Label Label5 
      Caption         =   "Jumlah"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10680
      TabIndex        =   10
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "s/d"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Tanggal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Kode Grup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      TabIndex        =   1
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "frmDaftarJual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Month(TDBDate2) <> Month(Date) And Year(TDBDate2) <> Year(Date) Then
            MsgBox "Tanggal harus pada Bulan Ini!", vbInformation
    Exit Sub
    ElseIf Month(TDBDate2) <> Month(Date) And Year(TDBDate2) <> Year(Date) Then
            MsgBox "Tanggal harus pada Bulan Ini!", vbInformation
    Exit Sub
    End If
   Data1.DatabaseName = App.path & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
 
'SQL = "SELECT TInv.[NoID], TInv.[IDInventor], TInv.[Kode], TInv.[Barcode], TInv.[Nama], TInv.[IDSatuan], TInv.[KodeSat], TInv.[Konversi], Tinv.[HargaJual] , TInv.[HargaPokok], TInv.[DiscExpired], TInv.[DiscMulai], TInv.[DiscProsen], TInv.[DiscRupiah], TInv.[IsPoin], TInv.[BKP],TInv.[IsPoinSupplier],TInv.[IDPoinSupplier],TInv.[IsOperator], TInv.[HargaMin],TInv.[DiscMulai],TInv.[DiscExpired], TInv.[Harga1], TInv.[Harga2], TInv.[Harga3] FROM TInv  ORDER BY " & NamaField
SQL = "SELECT * from MSales where Tanggal>=#" & Format(TDBDate1, "MM/dd/yyyy") & "# and Tanggal<#" & Format(TDBDate2 + 1, "MM/dd/yyyy") & "# "
If Text2.Text <> "" Then
SQL = SQL & " and Sopir='" & Replace(Text2.Text, "'", "''") & "'"
End If
SQL = SQL & " ORDER BY Tanggal"
Data1.RecordSource = SQL
Data1.Refresh
SQL = "SELECT Sum(hargaTotal) as Total,Sum(KomisiRp) as Jumlah from MSales where Tanggal>=#" & Format(TDBDate1, "MM/dd/yyyy") & "# and Tanggal<#" & Format(TDBDate2 + 1, "MM/dd/yyyy") & "# "
If Text2.Text <> "" Then
SQL = SQL & " and Sopir='" & Replace(Text2.Text, "'", "''") & "'"
End If
    Dim dbs As Database
    Dim rs As Recordset
    Set dbs = OpenDatabase(App.path & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
    Set rs = dbs.OpenRecordset(SQL)
    If rs.EOF And rs.BOF Then
    txtJumlah.Text = "0.00"
    txtTotal.Text = "0.00"
    Else
    txtJumlah.Text = Format(NullToNol(rs!Jumlah), "#,##0.00")
    txtTotal.Text = Format(NullToNol(rs!Total), "#,##0.00")
    End If
    rs.Close
    Set rs = Nothing
    dbs.Close
    Set dbs = Nothing
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Data1.Database.Close
End Sub

Private Sub Text1_Change()

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
   'Text1.SetFocus
Case "SCK"
    frmCustomer.NamaField = "Kode"
    frmCustomer.lbCari = "CARI KODE"
    frmCustomer.TampilMarketing isambil, kdMember, discMarketing
    If isambil Then
         
        Text2.Text = kdMember
       ' Text1.Text = Format(discMarketing, "##0.##")
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
End Select
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
