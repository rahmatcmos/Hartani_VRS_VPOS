VERSION 5.00
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dxctrls.dll"
Begin VB.Form frmSetDeviceKASIR 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setting Device"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8370
   Icon            =   "frmSetDeVice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   8370
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
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
      Height          =   360
      ItemData        =   "frmSetDeVice.frx":628A
      Left            =   2655
      List            =   "frmSetDeVice.frx":6294
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   6570
      Width           =   1170
   End
   Begin VB.TextBox txtDrawer 
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
      Left            =   2640
      TabIndex        =   9
      Top             =   2205
      Width           =   1275
   End
   Begin VB.TextBox txtDatabase 
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
      IMEMode         =   3  'DISABLE
      Left            =   2640
      TabIndex        =   21
      Top             =   4815
      Width           =   5715
   End
   Begin VB.ComboBox cbTipeCetakan 
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
      Height          =   360
      ItemData        =   "frmSetDeVice.frx":62A3
      Left            =   2640
      List            =   "frmSetDeVice.frx":62A5
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   6165
      Width           =   5715
   End
   Begin VB.TextBox Text4 
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
      Left            =   2640
      TabIndex        =   25
      Top             =   5745
      Width           =   5715
   End
   Begin VB.TextBox Text3 
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
      Left            =   2640
      TabIndex        =   23
      Top             =   5325
      Width           =   5715
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
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   19
      Top             =   4365
      Width           =   5715
   End
   Begin VB.TextBox Text1 
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
      Left            =   2640
      TabIndex        =   17
      Top             =   3945
      Width           =   5715
   End
   Begin VB.TextBox txtFooter 
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
      Left            =   2640
      TabIndex        =   15
      Top             =   3465
      Width           =   1275
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Tutup"
      Height          =   555
      Left            =   6975
      TabIndex        =   34
      Top             =   7935
      Width           =   1395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Simpan"
      Height          =   555
      Left            =   5535
      TabIndex        =   33
      Top             =   7935
      Width           =   1395
   End
   Begin VB.TextBox txtjudul 
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
      Height          =   825
      Left            =   2640
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   510
      Width           =   5715
   End
   Begin VB.TextBox txtKassa 
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
      Left            =   2640
      TabIndex        =   5
      Top             =   1380
      Width           =   1275
   End
   Begin VB.TextBox txtPerusahaan 
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
      Left            =   2640
      TabIndex        =   1
      Top             =   90
      Width           =   5715
   End
   Begin VB.TextBox txtCustomerDisplay 
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
      Left            =   2640
      TabIndex        =   13
      Top             =   3045
      Width           =   1275
   End
   Begin VB.TextBox txtBarcode 
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
      Left            =   2640
      TabIndex        =   11
      Top             =   2625
      Width           =   1275
   End
   Begin VB.TextBox txtPrinter 
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
      Left            =   2640
      TabIndex        =   7
      Top             =   1800
      Width           =   1275
   End
   Begin CONTROLSLibCtl.dxCheckBox dxCheckBox1 
      Height          =   270
      Left            =   2670
      TabIndex        =   35
      Top             =   8040
      Width           =   2640
      _Version        =   65536
      _cx             =   4657
      _cy             =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Tampil Saldo Poin Member"
      Enabled         =   -1  'True
      AutoSize        =   -1  'True
      BackStyle       =   1
      BackColor       =   14215660
      ForeColor       =   0
      ViewStyle       =   1
      Checked         =   0   'False
      GroupIndex      =   -1
      TextLayout      =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   12632256
   End
   Begin CONTROLSLibCtl.dxCheckBox ckResetPerKasir 
      Height          =   270
      Left            =   2670
      TabIndex        =   32
      Top             =   7680
      Width           =   2430
      _Version        =   65536
      _cx             =   4286
      _cy             =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Reset Per Kasir Per Shift"
      Enabled         =   -1  'True
      AutoSize        =   -1  'True
      BackStyle       =   1
      BackColor       =   14215660
      ForeColor       =   0
      ViewStyle       =   1
      Checked         =   0   'False
      GroupIndex      =   -1
      TextLayout      =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   12632256
   End
   Begin CONTROLSLibCtl.dxCheckBox ckKantong 
      Height          =   270
      Left            =   2670
      TabIndex        =   31
      Top             =   7350
      Width           =   2130
      _Version        =   65536
      _cx             =   3757
      _cy             =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Pakai Kantong Plastik"
      Enabled         =   -1  'True
      AutoSize        =   -1  'True
      BackStyle       =   1
      BackColor       =   14215660
      ForeColor       =   0
      ViewStyle       =   1
      Checked         =   0   'False
      GroupIndex      =   -1
      TextLayout      =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   12632256
   End
   Begin CONTROLSLibCtl.dxCheckBox ckHargaBeli 
      Height          =   270
      Left            =   2670
      TabIndex        =   30
      Top             =   7020
      Width           =   2505
      _Version        =   65536
      _cx             =   4419
      _cy             =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Kunci Harga Di Harga Beli"
      Enabled         =   -1  'True
      AutoSize        =   -1  'True
      BackStyle       =   1
      BackColor       =   14215660
      ForeColor       =   0
      ViewStyle       =   1
      Checked         =   0   'False
      GroupIndex      =   -1
      TextLayout      =   1
      UseMaskColor    =   -1  'True
      MaskColor       =   12632256
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo Stock Tampil"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   330
      TabIndex        =   28
      Top             =   6660
      Width           =   1785
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Port Drawer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   315
      TabIndex        =   8
      Top             =   2250
      Width           =   1200
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Database"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   315
      TabIndex        =   20
      Top             =   4860
      Width           =   930
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipe Cetakan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   315
      TabIndex        =   26
      Top             =   6225
      Width           =   1245
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IDPOS Default"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   315
      TabIndex        =   24
      Top             =   5775
      Width           =   1350
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NoID Gudang yg Dipakai"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   315
      TabIndex        =   22
      Top             =   5355
      Width           =   2280
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password Server"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   315
      TabIndex        =   18
      Top             =   4395
      Width           =   1665
   End
   Begin VB.Line Line1 
      X1              =   2640
      X2              =   11040
      Y1              =   3885
      Y2              =   3885
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Server"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   315
      TabIndex        =   16
      Top             =   3975
      Width           =   1245
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Spasi Struk Bawah"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   315
      TabIndex        =   14
      Top             =   3495
      Width           =   1800
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   315
      TabIndex        =   2
      Top             =   540
      Width           =   690
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. Kassa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   315
      TabIndex        =   4
      Top             =   1410
      Width           =   930
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Perusahaan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   315
      TabIndex        =   0
      Top             =   120
      Width           =   1740
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Port Customer Display"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   315
      TabIndex        =   12
      Top             =   3075
      Width           =   2160
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Port Scanner"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   315
      TabIndex        =   10
      Top             =   2655
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Port Printer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   315
      TabIndex        =   6
      Top             =   1830
      Width           =   1155
   End
End
Attribute VB_Name = "frmSetDeviceKASIR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbs As Database
Dim rs As Recordset

Private Sub Command1_Click()
On Error GoTo pesan
Set dbs = OpenDatabase(DirDatabase & "\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
      Set rs = dbs.OpenRecordset("Umum")
      If rs.EOF And rs.BOF Then
        rs.AddNew
      Else
      rs.Edit
      End If
        rs!Perusahaan = txtPerusahaan.Text
        rs!Judul = txtjudul.Text
        rs!kassa = txtKassa.Text
        rs!namaBarcode = txtBarcode.Text
        rs!NamaPrinter = txtPrinter.Text
        rs!NamaDrawer = txtDrawer.Text
        rs!NamaCustomerDisplay = txtCustomerDisplay.Text
        rs!SpasiFooter = NullToNol(txtFooter.Text)
        rs!IsCekStock = NullToNol(Combo1.ListIndex)
        rs.Update
    dbs.Close
    
Set dbs = OpenDatabase(DirDatabase & "\TempDB.mdb")
      Set rs = dbs.OpenRecordset("Umum")
      If rs.EOF And rs.BOF Then
        rs.AddNew
      Else
      rs.Edit
      End If
        rs!Perusahaan = txtPerusahaan.Text
        rs!Judul = txtjudul.Text
        rs!kassa = txtKassa.Text
        rs!namaBarcode = txtBarcode.Text
        rs!NamaPrinter = txtPrinter.Text
        rs!NamaDrawer = txtDrawer.Text
        rs!NamaCustomerDisplay = txtCustomerDisplay.Text
        rs!SpasiFooter = NullToNol(txtFooter.Text)
        rs!IsCekStock = NullToNol(Combo1.ListIndex)
        rs.Update
    dbs.Close
    
    Open App.path & "\database\SettingServer.dat" For Output As #1  ' Open file for output.
    Print #1, Text1.Text
    Print #1, "sa"
    Print #1, Text2.Text
    Print #1, txtDatabase.Text
    Print #1, Text4.Text
    Print #1, Text3.Text
    Print #1, Text1.Text
    Print #1, "sa"
    Print #1, Text2.Text
    Print #1, txtDatabase
    Print #1, ""
    Print #1, cbTipeCetakan.ListIndex
    Print #1, ckHargaBeli.Checked
    Print #1, ckKantong.Checked
    Print #1, ckResetPerKasir.Checked
    Close #1
    
    MsgBox "Data tersimpan!"
    Unload Me
Exit Sub
pesan:
MsgBox "Ada kesalahan!" & vbCrLf & Err.Description, vbCritical + vbOKOnly
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
  cbTipeCetakan.AddItem "(none)", 0
  cbTipeCetakan.AddItem "Preview", 1
  cbTipeCetakan.AddItem "Print", 2
 cbTipeCetakan.AddItem "Optional", 2
  cbTipeCetakan.ListIndex = 0
Dim kal() As String 'serversales, uid,pwd,dbs_sales, IDPOS,IDGUDANG,serverPoint,uidpoint,pwd,dbs
                    '  0            1  2   3           4     5       6              7     8,  9
                    
    Set dbs = OpenDatabase(DirDatabase & "\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
      Set rs = dbs.OpenRecordset("Umum")
      If rs.EOF And rs.BOF Then
        txtPerusahaan.Text = ""
        txtjudul.Text = ""
        txtKassa.Text = ""
        txtBarcode.Text = ""
        txtPrinter.Text = ""
        txtDrawer.Text = ""
        txtCustomerDisplay.Text = ""
        txtFooter.Text = 0
        Text3.Text = 1
        Text4.Text = 1
        Text1.Text = ".\SQLEXPRESS"
        Text2.Text = "sahasystem"
Combo1.ListIndex = 0
      Else
        txtPerusahaan.Text = NullToStr(rs!Perusahaan)
        txtjudul.Text = NullToStr(rs!Judul)
        txtKassa.Text = NullToStr(rs!kassa)
        txtBarcode.Text = NullToStr(rs!namaBarcode)
        txtPrinter.Text = NullToStr(rs!NamaPrinter)
        txtDrawer.Text = NullToStr(rs!NamaDrawer)
        txtCustomerDisplay.Text = NullToStr(rs!NamaCustomerDisplay)
        txtFooter.Text = NullToNol(rs!SpasiFooter)
        Combo1.ListIndex = NullToNol(rs!IsCekStock)
        Open App.path & "\database\SettingServer.dat" For Input As #1
        i = 0
        While Not EOF(1)
          ReDim Preserve kal(i + 1)
          Input #1, kal(i)
          i = i + 1
        Wend
        Close #1
        
        Text3.Text = kal(5)
        Text4.Text = kal(4)
        Text1.Text = kal(0)
        Text2.Text = kal(2)
        txtDatabase.Text = kal(3)
        cbTipeCetakan.ListIndex = NullToNol(kal(11))
        ckHargaBeli.Checked = NullToBool(kal(12))
        ckKantong.Checked = NullToBool(kal(13))
        ckResetPerKasir.Checked = NullToBool(kal(14))
      End If
    dbs.Close
End Sub

Private Sub txtKassa_Change()
On Error Resume Next
  Text4.Text = CLng(txtKassa.Text)
End Sub
