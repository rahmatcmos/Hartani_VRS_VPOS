VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1710
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1395
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   585
      TabIndex        =   0
      Top             =   315
      Width           =   1680
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    On Error Resume Next
    Dim dbs As Database
    Set dbs = OpenDatabase(App.Path & "\database\tempdb_" & Format(Date, "yyyyMM") & ".mdb")
    'dbs.Execute "update msalesd set Jumlah=Qty*Harga where IDSales<(select Min(NoID) from MSales where (Tanggal>=#03/07/2012#));"
    'dbs.Execute "update msalesd set Jumlah=int(Qty*Harga+0.5) where IDSales>=(select Min(NoID) from MSales where Tanggal>=#03/07/2012#)"
    dbs.Execute "ALTER TABLE MSales ADD Sopir Text(30)"
    dbs.Execute "ALTER TABLE MSales ADD Komisi double default 0.0"
    dbs.Execute "ALTER TABLE MSales ADD KomisiRp money default 0.0"

    dbs.Execute "alter table MSalesDVoucher add Qty Integer"
    dbs.Execute "update MSalesDVoucher  set Qty =1 where IsNUll(Qty)=True"
        
    dbs.Execute "alter table mreset add Charge currency"
    dbs.Execute "alter table mreset add Retur currency"
    dbs.Execute "alter table mreset add CountDebetMandiri Int"
    dbs.Execute "alter table mreset add CountDebetBCA Int"
    dbs.Execute "alter table mreset add CountDebetLain Int"
    dbs.Execute "alter table mreset add CountCreditVisa Int"
    dbs.Execute "alter table mreset add CountCreditBCA Int"
    dbs.Execute "alter table mreset add CountCreditMaster Int"
    dbs.Execute "alter table mreset add CountCreditLain Int"
    dbs.Execute "alter table mreset add JumlahKasKeluar currency"
    dbs.Execute "alter table mreset add JumlahKomisi currency"
    dbs.Execute "alter table MSalesD add SaldoStock double"
    dbs.Execute "alter table Umum add IsCekStock int"
    dbs.Execute "Create Table MCetakSales(NoID Long CONSTRAINT NoIDCetakConstraint Primary Key);"
    dbs.Execute "Create Table MSalesDVoucher(NoID Long CONSTRAINT MyFieldConstraint Primary Key,IDSales Long,IDPenerbit int,NamaPenerbit text(50),IDVoucher long,NoVoucher text (30),Nominal currency);"
    dbs.Execute "Create INDEX idxIDSales ON MSalesDVoucher(IDSales);"
    dbs.Execute "Create INDEX idxIDPenerbit ON MSalesDVoucher(IDPenerbit);"
    dbs.Execute "Create INDEX idxIDVoucher ON MSalesDVoucher(IDVoucher);"
    dbs.Execute "Create INDEX IDSales ON MSalesD(IDSales);"
    dbs.Execute "Create Table MKasKeluar(NOID int,Tanggal Date,Shift int,Jumlah Currency,KodeKasir text(30),IDKasir int, NamaKasir text(50),KodePengawas text(30),IDPengawas int, NamaPengawas text(50))"
    dbs.Close
    Set dbs = OpenDatabase(App.Path & "\database\tempdb.mdb")
    dbs.Execute "ALTER TABLE MSales ADD Sopir Text(30)"
    dbs.Execute "ALTER TABLE MSales ADD Komisi double default 0.0"
    dbs.Execute "ALTER TABLE MSales ADD KomisiRp money default 0.0"

    dbs.Execute "alter table MSalesDVoucher add Qty Integer"
    dbs.Execute "update MSalesDVoucher  set Qty =1 where IsNUll(Qty)=True"
        
    'dbs.Execute "update msalesd set Jumlah=int(Qty*Harga+0.5)"
    dbs.Execute "alter table mreset add Charge currency"
    dbs.Execute "alter table mreset add Retur currency"
    dbs.Execute "alter table mreset add CountDebetMandiri Int"
    dbs.Execute "alter table mreset add CountDebetBCA Int"
    dbs.Execute "alter table mreset add CountDebetLain Int"
    dbs.Execute "alter table mreset add CountCreditVisa Int"
    dbs.Execute "alter table mreset add CountCreditBCA Int"
    dbs.Execute "alter table mreset add CountCreditMaster Int"
    dbs.Execute "alter table mreset add CountCreditLain Int"
    dbs.Execute "alter table mreset add JumlahKasKeluar currency"
    dbs.Execute "alter table mreset add JumlahKomisi currency"
    dbs.Execute "alter table MSalesD add SaldoStock double"
    dbs.Execute "alter table Umum add IsCekStock int"
    dbs.Execute "Create Table MCetakSales(NoID Long CONSTRAINT NoIDCetakConstraint Primary Key);"
    dbs.Execute "Create Table MSalesDVoucher(NoID Long CONSTRAINT MyFieldConstraint Primary Key,IDSales Long,IDPenerbit int,NamaPenerbit text(50),IDVoucher long,NoVoucher text (30),Nominal currency);"
    dbs.Execute "Create INDEX idxIDSales ON MSalesDVoucher(IDSales);"
    dbs.Execute "Create INDEX idxIDPenerbit ON MSalesDVoucher(IDPenerbit);"
    dbs.Execute "Create INDEX idxIDVoucher ON MSalesDVoucher(IDVoucher);"
    dbs.Execute "Create INDEX IDSales ON MSalesD(IDSales);"
    dbs.Execute "Create Table MKasKeluar(NOID int,Tanggal Date,Shift int,Jumlah Currency,KodeKasir text(30),IDKasir int, NamaKasir text(50),KodePengawas text(30),IDPengawas int, NamaPengawas text(50))"
    dbs.Close
    Set dbs = Nothing
    MsgBox "Proses Selesai!"
    Exit Sub
pesan:
    MsgBox "Ada salah: " & Err.Description, vbCritical
End Sub
'083857225218


