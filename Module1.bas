Attribute VB_Name = "Module1"
Option Explicit
Public KodeUserLogin As String
Public isRemcomendedOnline As Boolean
Public isTampilSaldoStock As Boolean
Public isOnline As Boolean
Public Const NMRESETGAGAL = "D:\Master\Temp\Rst"
Public Const NMBACKUPGAGAL = "D:\Master\Temp\Bck"
Public Const DefPembulatan = 1
Public IDMember As Long
Public TipeHargaJual As Integer
Public Const ISSembunyikanFooterbarangKSB As Boolean = True
Public NamaMember As String
Public KodeMember As String
Public defDiscMember As Double
Public defMinMemberdapatDisc As Double
Public defDiscMemberBolehInput As Boolean

Public IDUser As Long
Public NamaKasir As String
Public PersenBiayaKartuKredit As Double
Public NamaMesin As String
Public NamaToko As String
Public SpasiFooter As Integer
Public DirDatabase As String
Public DirReset As String
Public DirBackup As String
Public DirReport As String
Public isSupervisor As Boolean
Public IsNotaSelesai As Boolean
Public IDNotaTerakhir As Long
Public SendByCode(0 To 255) As String * 3
Public Judulstruk As String
Public comDisplay As Object
Public comPrinter As Object
Public comDrawer As Object
Public KeyKode As Integer
Public isHasilKonversi As Boolean
Public isRun As Boolean
Public bolehbergerak As Boolean
Public DirUpdate As String
Public NoPortPrinter As Integer
Public NoPortDisplay As Integer
Public NoPortBarcode As Integer
Public NoPortDrawer As Integer
Public KodeKasir As String
Public IsPajak As String
Public PersenLap As Double
Public KodeUserDua As String
Public NamaShift As Integer
Public DirDbServer As String
Public PathVoucher As String
Public Footer1 As String
Public Footer2 As String
Public Footer3 As String
Public Const IsHematKertas As Boolean = False
Public LimitHutang As Double
Public x As New clsSerial
Public IDPengawas_ As Long
Public KodePengawas_ As String
Public NamaPengawas_ As String

Public Function DirSaja(ByVal nmfile As String) As String
'\\SERVER\XEON\A\asdf
Dim i, pos As Long
i = InStr(1, nmfile, "\")
pos = i
    Do While i < Len(nmfile)
    i = InStr(pos + 1, nmfile, "\")
    If i = 0 Then
        i = Len(nmfile)
    Else
        pos = i
    End If
    Loop
    DirSaja = Left(nmfile, pos)
End Function
Sub BacaFileVcr()
Open App.path & "\database\voucher.txt" For Input As #1 ' Open file for input.
Do While Not EOF(1) ' Loop until end of file.
    Input #1, PathVoucher
Loop
Close #1    ' Close file.
End Sub

Sub Main()
'Aktifkan Modul Serial
If x.CekCPUTrial("SGT56VPOS32001WEM789FOH87") Then
  Mainkan
Else
  If x.HasilX <> HarusDitutup Then
    Mainkan
  Else
    End
  End If
End If
'Tanpa Serial
'  Mainkan
End Sub
Public Sub Mainkan()
    DirDbServer = App.path & "\Database" 'jika ONLINE
    BacaFileVcr
    DirUpdate = App.path & "\Update"
    DirDatabase = App.path & "\Database"
    DirReset = App.path & "\Reset"
    DirBackup = App.path & "\Backup"
    DirReport = App.path & "\Report"
    BuatLogApp ("Aplikasi  Start...")
  DoEvents
  BuatLogApp ("Update Struktur database...")
  DoEvents
    TambahPembulatan
  BuatLogApp ("Update Struktur Pembulatan...OK.")
  DoEvents
    TambahDiscountIntern
    TambahFieldIsUpload
    TambahFieldDiscMember2
    TambahFieldResetDiscMember2
    TambahFieldIsPDP
    TambahFieldIsDisc2
  TambahFieldDiskon
  TambahFieldReedem
  TambahFieldReedemReset
  TambahFieldIsMember
  BuatLogApp ("Update Struktur database...Sukses!")
  DoEvents
  bacaSettingServer
  BuatLogApp ("Form Login...Inisialisasi!")
  
    Form2.Show
    Form2.SetFocus
  BuatLogApp ("Form Login...Show")
  
  'GetTombol
  'frmBarang.NamaField = "Nama"
  'frmBarang.lbCari = "CARI NAMA"
  '
  'frmBarang.Show
  'frmBarang.SetFocus
End Sub
Sub TambahPembulatan()
On Error GoTo akhir
Dim dbs As Database
'Dim TblDef As TableDef
'    Dim Fld As Field
'    Dim Idx As Index
Dim NamaFiledb As String
NamaFiledb = App.path & "\database\TempDB.mdb"
Set dbs = OpenDatabase(NamaFiledb)

dbs.Execute "ALTER TABLE MSales ADD Pembulatan Currency "
dbs.Execute "UPDATE MSales Set Pembulatan=0 WHERE IsNull(Pembulatan)"
If Dir(App.path & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb") <> "" Then
  NamaFiledb = App.path & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
  Set dbs = OpenDatabase(NamaFiledb)
  dbs.Execute "ALTER TABLE MSales ADD Pembulatan Currency "
  dbs.Execute "UPDATE MSales Set Pembulatan=0 WHERE IsNull(Pembulatan)"
End If
akhir:
  If Not dbs Is Nothing Then
    dbs.Close
  End If
End Sub
Sub TambahFieldIsUpload()
On Error GoTo akhir
Dim dbs As Database
'Dim TblDef As TableDef
'    Dim Fld As Field
'    Dim Idx As Index
Dim NamaFiledb As String
NamaFiledb = App.path & "\database\TempDB.mdb"
Set dbs = OpenDatabase(NamaFiledb)
dbs.Execute "ALTER TABLE MSales ADD IsUpload Bit NOT NULL "
If Dir(App.path & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb") <> "" Then
  NamaFiledb = App.path & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
  Set dbs = OpenDatabase(NamaFiledb)
  dbs.Execute "ALTER TABLE MSales ADD IsUpload Bit NOT NULL "
End If
akhir:
If Not dbs Is Nothing Then
  dbs.Close
End If
End Sub
Sub TambahDiscountIntern()
On Error GoTo akhir
Dim dbs As Database
'Dim TblDef As TableDef
'    Dim Fld As Field
'    Dim Idx As Index
Dim NamaFiledb As String
NamaFiledb = App.path & "\database\TempDB.mdb"
Set dbs = OpenDatabase(NamaFiledb)
'dbs.Execute "ALTER TABLE tblcategories ADD email1 text(60) "
'dbs.Execute "ALTER TABLE tblcategories ADD email2 text(60) NULL"
'dbs.Execute "ALTER TABLE tblcategories ADD isUpdateInfo BIT "
dbs.Execute "ALTER TABLE MSales ADD IDMember INT "

dbs.Execute "ALTER TABLE MSales ADD BarangKSB Currency "
dbs.Execute "ALTER TABLE MSales ADD SisaKSB Currency "

dbs.Execute "ALTER TABLE MSales ADD DiscIntern Currency "
dbs.Execute "ALTER TABLE MSalesD ADD DiscInternRp Currency"
dbs.Execute "ALTER TABLE MSalesD ADD DiscInternProsen Double"
dbs.Execute "ALTER TABLE MSales ADD JumDiscInternRp Currency "
dbs.Execute "UPDATE MSales SET DiscIntern=0 WHERE IsNull(DiscIntern)"
dbs.Execute "UPDATE MSalesD SET DiscInternRp=0 WHERE IsNull(DiscInternRp)"
dbs.Execute "UPDATE MSalesD SET DiscInternProsen=0 WHERE IsNull(DiscInternProsen)"
dbs.Execute "UPDATE MSales Set JumDiscInternRp=0 WHERE IsNull(JumDiscInternRp)"

If Dir(App.path & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb") <> "" Then
  NamaFiledb = App.path & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
  Set dbs = OpenDatabase(NamaFiledb)
  'dbs.Execute "ALTER TABLE tblcategories ADD email1 text(60) "
  'dbs.Execute "ALTER TABLE tblcategories ADD email2 text(60) NULL"
  'dbs.Execute "ALTER TABLE tblcategories ADD isUpdateInfo BIT "
  dbs.Execute "ALTER TABLE MSales ADD IDMember INT "
  
  dbs.Execute "ALTER TABLE MSales ADD BarangKSB Currency "
  dbs.Execute "ALTER TABLE MSales ADD SisaKSB Currency "
  
  dbs.Execute "ALTER TABLE MSales ADD DiscIntern Currency "
  dbs.Execute "ALTER TABLE MSalesD ADD DiscInternRp Currency"
  dbs.Execute "ALTER TABLE MSalesD ADD DiscInternProsen Double"
  dbs.Execute "ALTER TABLE MSales ADD JumDiscInternRp Currency "
  dbs.Execute "UPDATE MSales SET DiscIntern=0 WHERE IsNull(DiscIntern)"
  dbs.Execute "UPDATE MSalesD SET DiscInternRp=0 WHERE IsNull(DiscInternRp)"
  dbs.Execute "UPDATE MSalesD SET DiscInternProsen=0 WHERE IsNull(DiscInternProsen)"
  dbs.Execute "UPDATE MSales Set JumDiscInternRp=0 WHERE IsNull(JumDiscInternRp)"
End If

'    'CREATE TABLE: MSales
'    '=============================
'    Set TblDef = dbs.TableDefs.Append("MSales")  '.CreateTableDef("MSales")
'    With TblDef
'        .Attributes = 0
'        .Connect = ""
'        .SourceTableName = ""
'        .ValidationRule = ""
'        .ValidationText = ""
'
'        'CREATE FIELD: NoID
'        '=============================
'        Set Fld = TblDef.CreateField("coba", 4, 4)
'        With Fld
'            .Attributes = 1
'            .DefaultValue = 0
'            .OrdinalPosition = 0
'            .Required = True
'            .ValidationRule = ""
'            .ValidationText = ""
'        End With
'        .Fields.Append Fld
'       End With
akhir:
  If Not dbs Is Nothing Then
    dbs.Close
  End If
End Sub
Sub TambahFieldReedem()
On Error Resume Next
Dim dbs As Database
Dim NamaFiledb As String
NamaFiledb = App.path & "\database\TempDB.mdb"
Set dbs = OpenDatabase(NamaFiledb)
dbs.Execute "ALTER TABLE MSales ADD IDReedemPoin Double"
dbs.Execute "ALTER TABLE MSales ADD ReedemPoin Double"
dbs.Execute "ALTER TABLE MSales ADD NilaiReedemPoin Currency"
dbs.Close
If Dir(App.path & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb") <> "" Then
  NamaFiledb = App.path & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
  Set dbs = OpenDatabase(NamaFiledb)
  dbs.Execute "ALTER TABLE MSales ADD IDReedemPoin Double"
  dbs.Execute "ALTER TABLE MSales ADD ReedemPoin Double"
  dbs.Execute "ALTER TABLE MSales ADD NilaiReedemPoin Currency"
End If
akhir:
On Error Resume Next
If Not dbs Is Nothing Then
  dbs.Close
End If
Err.Clear
End Sub

Sub TambahFieldReedemReset()
On Error Resume Next
Dim dbs As Database
Dim NamaFiledb As String
NamaFiledb = App.path & "\database\TempDB.mdb"
Set dbs = OpenDatabase(NamaFiledb)
dbs.Execute "ALTER TABLE MReset ADD ReedemPoin Double"
dbs.Execute "ALTER TABLE MReset ADD NilaiReedemPoin Currency"
dbs.Close
If Dir(App.path & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb") <> "" Then
  NamaFiledb = App.path & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
  Set dbs = OpenDatabase(NamaFiledb)
  dbs.Execute "ALTER TABLE MReset ADD ReedemPoin Double"
  dbs.Execute "ALTER TABLE MReset ADD NilaiReedemPoin Currency"
End If
akhir:
On Error Resume Next
If Not dbs Is Nothing Then
  dbs.Close
End If
Err.Clear
End Sub
Sub TambahFieldDiskon()
On Error GoTo akhir
Dim dbs As Database
'Dim TblDef As TableDef
'    Dim Fld As Field
'    Dim Idx As Index
Dim NamaFiledb As String
NamaFiledb = App.path & "\database\TempDB.mdb"
Set dbs = OpenDatabase(NamaFiledb)
'dbs.Execute "ALTER TABLE tblcategories ADD email1 text(60) "
'dbs.Execute "ALTER TABLE tblcategories ADD email2 text(60) NULL"
'dbs.Execute "ALTER TABLE tblcategories ADD isUpdateInfo BIT "
dbs.Execute "ALTER TABLE MSalesD ADD IsMember BIT "
dbs.Execute "ALTER TABLE MSalesD ADD HargaBruto Currency "
dbs.Execute "ALTER TABLE MSalesD ADD DiscRp Currency"
dbs.Execute "ALTER TABLE MSalesD ADD DiscProsen Double"
If Dir(App.path & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb") <> "" Then
  NamaFiledb = App.path & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
  Set dbs = OpenDatabase(NamaFiledb)
  'dbs.Execute "ALTER TABLE tblcategories ADD email1 text(60) "
  'dbs.Execute "ALTER TABLE tblcategories ADD email2 text(60) NULL"
  'dbs.Execute "ALTER TABLE tblcategories ADD isUpdateInfo BIT "
  dbs.Execute "ALTER TABLE MSalesD ADD IsMember BIT "
  dbs.Execute "ALTER TABLE MSalesD ADD HargaBruto Currency "
  dbs.Execute "ALTER TABLE MSalesD ADD DiscRp Currency"
  dbs.Execute "ALTER TABLE MSalesD ADD DiscProsen Double"
End If

'    'CREATE TABLE: MSales
'    '=============================
'    Set TblDef = dbs.TableDefs.Append("MSales")  '.CreateTableDef("MSales")
'    With TblDef
'        .Attributes = 0
'        .Connect = ""
'        .SourceTableName = ""
'        .ValidationRule = ""
'        .ValidationText = ""
'
'        'CREATE FIELD: NoID
'        '=============================
'        Set Fld = TblDef.CreateField("coba", 4, 4)
'        With Fld
'            .Attributes = 1
'            .DefaultValue = 0
'            .OrdinalPosition = 0
'            .Required = True
'            .ValidationRule = ""
'            .ValidationText = ""
'        End With
'        .Fields.Append Fld
'       End With
akhir:
If Not dbs Is Nothing Then
  dbs.Close
End If
End Sub

Sub TambahFieldIsDisc2()
On Error Resume Next
Dim dbs As Database
Dim NamaFiledb As String
If Dir(App.path & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb") <> "" Then
  NamaFiledb = App.path & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
  Set dbs = OpenDatabase(NamaFiledb)
  dbs.Execute "ALTER TABLE MSalesD ADD HargaNormal Money "
  dbs.Execute "ALTER TABLE MSalesD ADD IsDisc2 BIT "
  dbs.Close
End If
If Dir(App.path & "\database\TempDB.mdb") <> "" Then
  NamaFiledb = App.path & "\database\TempDB.mdb"
  Set dbs = OpenDatabase(NamaFiledb)
  dbs.Execute "ALTER TABLE MSalesD ADD HargaNormal Money "
  dbs.Execute "ALTER TABLE MSalesD ADD IsDisc2 BIT "
  dbs.Close
End If
Set dbs = Nothing
akhir:
On Error Resume Next
  dbs.Close
  Set dbs = Nothing
End Sub

Sub TambahFieldIsPDP()
On Error GoTo akhir
Dim dbs As Database
Dim NamaFiledb As String
If Dir(App.path & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb") <> "" Then
  NamaFiledb = App.path & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
  Set dbs = OpenDatabase(NamaFiledb)
  dbs.Execute "ALTER TABLE MSalesD ADD IsPDP BIT "
  dbs.Close
End If
If Dir(App.path & "\database\TempDB.mdb") <> "" Then
  NamaFiledb = App.path & "\database\TempDB.mdb"
  Set dbs = OpenDatabase(NamaFiledb)
  dbs.Execute "ALTER TABLE MSalesD ADD IsPDP BIT "
  dbs.Close
End If
Set dbs = Nothing
akhir:
On Error Resume Next
  dbs.Close
  Set dbs = Nothing
End Sub

Sub TambahFieldDiscMember2()
On Error GoTo akhir
Dim dbs As Database
Dim NamaFiledb As String
If Dir(App.path & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb") <> "" Then
  NamaFiledb = App.path & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
  Set dbs = OpenDatabase(NamaFiledb)
  dbs.Execute "ALTER TABLE MSalesD ADD DiscRp2 Currency "
  dbs.Execute "ALTER TABLE MSalesD ADD DiscProsen2 Double "
  dbs.Close
End If
If Dir(App.path & "\database\TempDB.mdb") <> "" Then
  NamaFiledb = App.path & "\database\TempDB.mdb"
  Set dbs = OpenDatabase(NamaFiledb)
  dbs.Execute "ALTER TABLE MSalesD ADD DiscRp2 Currency "
  dbs.Execute "ALTER TABLE MSalesD ADD DiscProsen2 Double "
  dbs.Close
End If
Set dbs = Nothing
akhir:
On Error Resume Next
  dbs.Close
  Set dbs = Nothing
End Sub

Sub TambahFieldResetDiscMember2()
On Error GoTo akhir
Dim dbs As Database
Dim NamaFiledb As String
If Dir(App.path & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb") <> "" Then
  NamaFiledb = App.path & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
  Set dbs = OpenDatabase(NamaFiledb)
  dbs.Execute "ALTER TABLE MReset ADD DiscMember2 Currency "
  dbs.Close
End If
If Dir(App.path & "\database\TempDB.mdb") <> "" Then
  NamaFiledb = App.path & "\database\TempDB.mdb"
  Set dbs = OpenDatabase(NamaFiledb)
  dbs.Execute "ALTER TABLE MReset ADD DiscMember2 Currency "
  dbs.Close
End If
Set dbs = Nothing
akhir:
On Error Resume Next
  dbs.Close
  Set dbs = Nothing
End Sub

Sub TambahFieldIsMember()
On Error GoTo akhir
Dim dbs As Database
'Dim TblDef As TableDef
'    Dim Fld As Field
'    Dim Idx As Index
Dim NamaFiledb As String
NamaFiledb = App.path & "\database\dbmaster.mdb"
Set dbs = OpenDatabase(NamaFiledb)

'dbs.Execute "ALTER TABLE tblcategories ADD email1 text(60) "
'dbs.Execute "ALTER TABLE tblcategories ADD email2 text(60) NULL"
'dbs.Execute "ALTER TABLE tblcategories ADD isUpdateInfo BIT "
'dbs.Execute "ALTER TABLE TInv ADD IsMember BIT "
dbs.Execute "Update TInv Set IsMember=0 where (nama  like '*DANCOW*' or nama  like '*MINYAK*' or nama like '*BERAS*') "
dbs.Close
Set dbs = Nothing
akhir:
On Error Resume Next
  dbs.Close
  Set dbs = Nothing
End Sub
Public Function BacaAngka(ByVal Angka As Currency) As String
10000     On Error GoTo MS_BooBoo
10010     'Do Not Modify These 2 Lines!!!
10020 Dim strAngka As String
10040 Dim Koma As Long
        strAngka = Format(Angka, "####.#0")
        Koma = CInt(Right(strAngka, 2))
10070 If Koma > 0 Then
10080 BacaAngka = Trim(BacaAngkaSaja(CLng(Angka - Koma / 100))) & " rupiah " & Trim(BacaAngkaSaja(Koma)) & " sen"
10090 Else
10100 BacaAngka = Trim(BacaAngkaSaja(CLng(Angka - Koma / 100))) & " rupiah "
10110 End If
10120     Exit Function
10130
10140 MS_BooBoo:
10150     If Err.Number > 0 Then Resume Next
10170     'Do Not Modify These 4 Lines!!!
10180
End Function
Public Function BacaAngkaSaja(ByVal Angka As Currency) As String
10000     On Error GoTo MS_BooBoo
10010     'Do Not Modify These 2 Lines!!!
10020
10030 Dim strAngka, Satuan(1 To 12) As String
10040 Dim baca(1 To 9) As String
10050 Dim x As Variant
      Dim Y
10060 Dim panjang, posisi, i, j As Integer
10070 BacaAngkaSaja = ""
10080 baca(1) = "se": baca(2) = "dua ": baca(3) = "tiga ": baca(4) = "empat ": baca(5) = "lima "
10090 baca(6) = "enam ": baca(7) = "tujuh ": baca(8) = "delapan ": baca(9) = "sembilan ":
10100 Satuan(1) = "": Satuan(2) = "puluh ": Satuan(3) = "ratus ": Satuan(4) = "ribu "
10110 Satuan(5) = "puluh ": Satuan(6) = "ratus ": Satuan(7) = "juta ": Satuan(8) = "puluh "
10120 Satuan(9) = "ratus ": Satuan(10) = "milyar ": Satuan(11) = "puluh ": Satuan(12) = "ratus ":
10130 strAngka = str(Angka)
10140 panjang = Len(strAngka)
10150 For j = 2 To panjang
10160     x = Val(Mid(strAngka, j, 1))
10170     posisi = panjang + 1 - j
10180     If x <> 0 Then
10190         'Kasus antara 11-19
10200         If (posisi = 2 Or posisi = 5 Or posisi = 8 Or posisi = 11) And x = 1 Then
10210             Y = Val(Mid(strAngka, j + 1, 1))
10220             If (Y <> 0) Then
10230                 BacaAngkaSaja = BacaAngkaSaja & baca(Y) & "belas " & Satuan(posisi - 1)
10240             Else
10250                 BacaAngkaSaja = BacaAngkaSaja & "sepuluh " & Satuan(posisi - 1)
10260             End If
10270             j = j + 1
10280         ElseIf x = 1 Then
10290             If posisi = 1 Then
10300                 BacaAngkaSaja = BacaAngkaSaja & "satu"
10310             ElseIf (posisi = 4 And panjang > 5) Or posisi = 10 Or posisi = 7 Then
10320                 BacaAngkaSaja = BacaAngkaSaja & "satu " & Satuan(posisi)
10330             Else
10340                 BacaAngkaSaja = BacaAngkaSaja & baca(x) & Satuan(posisi)
10350             End If
10360         Else
10370             BacaAngkaSaja = BacaAngkaSaja & baca(x) & Satuan(posisi)
10380         End If
10390     ElseIf panjang > 5 Then
10400         If posisi = 4 And (Val(Mid(strAngka, j - 1, 1)) <> 0 Or Val(Mid(strAngka, j - 2, 1)) <> 0) Then
10410             BacaAngkaSaja = BacaAngkaSaja & " ribu "
10420         End If
10430         If panjang > 8 Then
10440             If posisi = 7 And (Val(Mid(strAngka, j - 1, 1)) <> 0 Or Val(Mid(strAngka, j - 2, 1)) <> 0) Then
10450                 BacaAngkaSaja = BacaAngkaSaja & " juta "
10460             End If
10470         End If
10480         If panjang > 11 Then
10490             If posisi = 10 And (Val(Mid(strAngka, j - 1, 1)) <> 0 Or Val(Mid(strAngka, j - 2, 1)) <> 0) Then
10500                 BacaAngkaSaja = BacaAngkaSaja & " milyar "
10510             End If
10520         End If
10530    End If
10540 Next
10550     Exit Function
10560
10570 MS_BooBoo:
10580     If Err > 0 Then Resume Next
10600     'Do Not Modify These 4 Lines!!!
10610
End Function
'Sub GetTombol()
'Dim dbs As Database
'Dim rs As Recordset
' Set dbs = OpenDatabase(DirDatabase & "\TempDB"& format(now,"_yyyyMM") &".mdb")
'
'  Set rs = dbs.OpenRecordset("SELECT * FROM Tombol WHERE Terpakai=true")
'  If rs.BOF And rs.EOF Then
'  Else
'    rs.MoveFirst
'    Do While Not rs.EOF
'      SendByCode(rs!Code) = rs!tombol
'      rs.MoveNext
'    Loop
'  End If
'  Set rs = Nothing
'  dbs.Close
'End Sub
Function NullToNol(x) As Double
If IsNull(x) Then
  NullToNol = 0
Else
  If IsNumeric(x) Then
    NullToNol = x
  Else
    NullToNol = 0
  End If
End If
End Function
Function NullToDate(x) As Date
If IsNull(x) Then
  NullToDate = CDate(#1/1/1990#)
Else
  If IsDate(x) Then
    NullToDate = x
  Else
    NullToDate = CDate(#1/1/1990#)
  End If
End If
End Function
Function NullToBool(x) As Boolean
If IsNull(x) Then
  NullToBool = False
ElseIf x = "" Then
  NullToBool = False
Else
  NullToBool = CBool(x)
End If
End Function

Function NullToStr(x) As String
If IsNull(x) Then
  NullToStr = " "
Else
  NullToStr = x
End If
End Function

Public Function GetNewID(ByVal nmTabel As String) As Long
  Dim dbs As Database
  Dim rs As Recordset
  Set dbs = OpenDatabase(DirDatabase & "\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
  Set rs = dbs.OpenRecordset("SELECT MAX(NoID) as ID FROM " & nmTabel)
  If rs.EOF And rs.BOF Then
    GetNewID = 1
  Else
    If IsNull(rs!ID) Then
      GetNewID = 1
    Else
      GetNewID = rs!ID + 1
    End If
  End If
  rs.Close
  dbs.Close
  Set rs = Nothing
  Set dbs = Nothing
End Function
Public Function GetNewNota(ByVal nmTabel As String, ByVal TANGGAL As Date, ByVal formatdate As String) As String
  Dim dbs As Database
  Dim rs As Recordset
  Set dbs = OpenDatabase(DirDatabase & "\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
  Set rs = dbs.OpenRecordset("SELECT MAX(Kode) as NoNota FROM " & nmTabel & " where year(tanggal) =" & Year(TANGGAL) & " and month(tanggal)=" & Month(TANGGAL) & " and day(tanggal)=" & Day(TANGGAL))
  If rs.EOF And rs.BOF Then
    GetNewNota = Format(1, "00000")
  Else
    GetNewNota = Format(NullToNol(rs!NoNota) + 1, "00000")
  End If
  rs.Close
  dbs.Close
  Set rs = Nothing
  Set dbs = Nothing
End Function
Public Function PakaiIDKosong(ByVal nmTabel As String, ByVal TANGGAL As Date) As Long
  Dim dbs As Database
  Dim rs As Recordset
  Set dbs = OpenDatabase(DirDatabase & "\TempDB" & Format(Now, "_yyyyMM") & ".mdb")
  Set rs = dbs.OpenRecordset("SELECT Min(MSales.NoID) as ID FROM MSales Left Join MSalesD on MSales.NoID=MSalesD.IDSales where year(MSales.tanggal) =" & Year(TANGGAL) & " and month(MSales.tanggal)=" & Month(TANGGAL) & " and day(MSales.tanggal)=" & Day(TANGGAL) & " group by msales.NoID,MSales.SubTotal having MSales.SubTotal=0 and count(MSalesD.NoID)=0 ")
  If rs.EOF And rs.BOF Then
    PakaiIDKosong = -1
  Else
    PakaiIDKosong = NullToNol(rs!ID)
  End If
  rs.Close
  dbs.Close
  Set rs = Nothing
  Set dbs = Nothing
End Function
Public Sub cetakdetil(ByVal kode As String, ByVal Nama As String, ByVal Qty As String, ByVal hargaSatuan As String, ByVal Jumlah As String)
  Prin Nama & Chr(13) & Chr(10) & kode & Space(16 - Min(Len(Left(kode, 13)) + Len(Qty), 16)) & Qty & " X " & Space(8 - Min(Len(hargaSatuan), 8)) & hargaSatuan & "=" & Space(11 - Len(Jumlah)) & Jumlah
End Sub

Public Sub CetakFooter(ByVal pesan As String, ByVal Jumlah As String)
    Prin pesan & Space(39 - Max(Len(pesan) + Len(Jumlah), 39)) & Jumlah
End Sub

Public Sub DisplayPesan(ByVal Pesan1 As String, ByVal pesan2 As String)

  Form3.Text2.Text = Pesan1
  Form3.Text3.Text = pesan2
On Error GoTo buka:
If comDisplay < 1 Then Exit Sub
  Dim errctr
  errctr = 0
  BukaPortDisplay
    comDisplay.Output = Chr(27) & "[2J" & Chr(13)
    comDisplay.Output = Chr(27) & "[1;1H" & Pesan1 & Chr(13)
    comDisplay.Output = Chr(27) & "[2;1H" & pesan2 & Chr(13)
    Exit Sub
buka:
    errctr = errctr + 1
    BukaPortDisplay
    If errctr > 1 Then Resume Next Else Resume
End Sub
Sub BukaPortDisplay()
On Error Resume Next
 If NoPortDisplay < 1 Or NoPortDisplay > 10 Then Exit Sub
    comDisplay.CommPort = NoPortDisplay
    comDisplay.InputLen = 0
    comDisplay.Settings = "9600,O,8,1"
    comDisplay.PortOpen = True
End Sub
Sub BuatLog(ByVal pesan As String)
If Dir(DirReport & "\" & Format(Date, "yyyyMMdd") & ".txt") <> "" Then
  Open DirReport & "\" & Format(Date, "yyyyMMdd") & ".txt" For Append As #1 ' Open file for output.
Else
  Open DirReport & "\" & Format(Date, "yyyyMMdd") & ".txt" For Output As #1 ' Open file for output.
End If
Print #1, pesan
Close #1  ' Close file.

End Sub
Sub BuatLogApp(ByVal pesan As String)
If Dir(DirReport & "\App" & Format(Date, "yyyyMMdd") & ".txt") <> "" Then
  Open DirReport & "\App" & Format(Date, "yyyyMMdd") & ".txt" For Append As #1 ' Open file for output.
Else
  Open DirReport & "\App" & Format(Date, "yyyyMMdd") & ".txt" For Output As #1 ' Open file for output.
End If
Print #1, Format(Now, "HH:mm:ss ") & pesan
Close #1  ' Close file.

End Sub

Sub Prin(ByVal pesan As String)
On Error GoTo buka:
  Dim errctr
  errctr = 0
  BuatLog pesan
If NoPortPrinter = 0 Then
  PrinLPT Chr(27) & Chr(33) & Chr(1) & pesan
ElseIf NoPortPrinter = -2 Then
  Exit Sub
ElseIf NoPortPrinter = -1 Then
  printFile pesan
Else
comPrinter.PortOpen = True
 comPrinter.Output = Chr(27) & Chr(33) & Chr(1) & pesan & Chr(13) & Chr(10)

End If
       Exit Sub
buka:
    errctr = errctr + 1
    BukaPortPrinter
    If errctr > 1 Then Resume Next Else Resume
End Sub
Sub BukaPortPrinter()
 On Error Resume Next
  If NoPortPrinter < 1 Or NoPortPrinter > 10 Then Exit Sub
  comPrinter.CommPort = NoPortPrinter
  comPrinter.InputLen = 0
  comPrinter.Settings = "9600,N,8,1"
  comPrinter.PortOpen = True
End Sub
Sub BukaPortDrawer()
 On Error Resume Next
  If NoPortDrawer < 1 Or NoPortDrawer > 99 Then Exit Sub
  comDrawer.CommPort = NoPortDrawer
  comDrawer.InputLen = 0
  comDrawer.Settings = "9600,N,8,1"
  comDrawer.PortOpen = True
End Sub
Sub PrinLPT(ByVal pesan As String)
On Error GoTo Handle
Open "PRN" For Output As #1  ' Open file for output.
Print #1, pesan
Close #1  ' Close file.
Exit Sub
Handle:
  If MsgBox("Printer belum siap", vbRetryCancel + vbQuestion) = vbRetry Then
    Resume
  Else
    Resume Next
  End If
  
End Sub

Sub printFile(ByVal pesan As String)
Open DirReport & "\tempPrin.txt" For Output As #1  ' Open file for output.
Print #1, pesan
Close #1  ' Close file.
Shell DirReport & "\prindos.bat", vbNormalFocus
End Sub

Sub openDrawerbyDos()
On Error GoTo buka:
  Dim errctr
  errctr = 0
 If NoPortDrawer = 0 Then
'  PrinLPT Chr(27) + "p" + Chr(0) + Chr(8) + Chr(16) ' open drawer
    PrinLPT Chr(27) + Chr(7) + Chr(11) + Chr(55) + Chr(7) ' open drawer'STAR
ElseIf NoPortDrawer > 0 Then
   comDrawer.Output = Chr(27) + "p" + Chr(0) + Chr(8) + Chr(16) ' open drawer
Else
  'tak ada action
End If
   Exit Sub
buka:
    errctr = errctr + 1
    BukaPortDrawer
    If errctr > 1 Then Resume Next Else Resume
End Sub

'Procedure PDWidth(cChar)
'IF PCOUNT()=0
'   cChar:=CHR(27)+CHR(33)+CHR(32)
'ELSEIF PCOUNT()>0.AND.VALTYPE(cChar)#'L'
'   cChar:=CHR(27)+CHR(33)+CHR(32)+cChar+CHR(13)+CHR(27)+CHR(33)+CHR(0)
'Else
'   cChar:=CHR(13)+CHR(27)+CHR(33)+CHR(0)
'End If
'RETURN cChar

Sub PrinBigChar(ByVal pesan As String)
On Error GoTo buka:
  Dim errctr
  errctr = 0
    BuatLog pesan
If NoPortPrinter = 0 Then
  PrinLPT Chr(27) & Chr(33) & Chr(32) & pesan & Chr(27) & Chr(33) & Chr(0)
ElseIf NoPortPrinter = -2 Then
  Exit Sub
ElseIf NoPortPrinter = -1 Then
  printFile pesan
Else
 comPrinter.Output = Chr(27) & Chr(33) & Chr(32) & pesan & Chr(27) & Chr(33) & Chr(0) & Chr(13) & Chr(10)
 'pesan & Chr(13) & Chr(10)

End If
       Exit Sub
buka:
    errctr = errctr + 1
    BukaPortPrinter
    If errctr > 1 Then Resume Next Else Resume

End Sub
Sub papercut()
On Error GoTo buka:
  Dim errctr
  errctr = 0
'  Prin Chr(13) & Chr(10)
   comPrinter.Output = Chr(27) + "@" + Chr(27) + "m"
   Exit Sub
buka:
    errctr = errctr + 1
    BukaPortPrinter
    If errctr > 1 Then Resume Next Else Resume
End Sub

Public Function Max(A, B) As Double
  If A > B Then
      Max = A
  Else
      Max = B
  End If
End Function
Public Function Min(A, B) As Double
  If A > B Then
      Min = B
  Else
      Min = A
  End If
End Function

Public Function GetStatusNetwork() As String
  On Error GoTo pesan
  If isRemcomendedOnline Then
        Dim sqlcon As ADODB.Connection
        Set sqlcon = New ADODB.Connection
        sqlcon.ConnectionString = Cnstr
        sqlcon.Open
        sqlcon.Close
        Set sqlcon = Nothing
        GetStatusNetwork = "ONLINE"
        isOnline = True
    Else
        isOnline = False
        GetStatusNetwork = "Local"
    End If
    Exit Function
pesan:
    BuatLogApp "SQL Server " & Err.Number & " : " & Err.Description & " : CN1 " & Cnstr & " : CN2 " & CNSTRPOINT
    On Error Resume Next
    Set sqlcon = Nothing
    isOnline = False
    GetStatusNetwork = "Local"
End Function
'Public Function Replace(x As String, asal As String, pengganti As String) As String
'Replace = x
'
''Replace = ""
''i = 1
''z = InStr(i, x, "'")
''Replace = Mid(x, i, z - i)
''i = z
''While z
''  z = InStr(i, x, "'")
''  Replace = Replace & Mid(x, i + 1, z - i)
''Wend
'End Function

Public Sub ApplyTombol(ByRef TXT As Object, ByVal ScanKode As Integer)
Dim Hasil As String
Hasil = SendByCode(ScanKode)
Select Case Hasil
Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "*", "`", "!", "@", "#", "$", "%", "^", "&", "<", "(", ")", "-", "+", "{", "}", "[", "]", "/", "?"
  TXT.Text = TXT.Text & Hasil
Case "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"
  TXT.Text = TXT.Text & Hasil
Case "PLU"
Case "HLD"
Case "CRC"
Case "VOD"
Case "AVD"
Case "RTN"
Case "CLR"
Case "ENT"
Case "ESC"
Case "STL"
Case "CSH"
Case "BK1"
Case "BK2"
Case "BK3"
Case "BK4"
Case "BK5"
Case "BK6"
Case "BK7"
Case "BK8"
Case "BK9"
Case "VCR"
Case "DBR"
Case "DBP"
Case "DTR"
Case "DTP"
Case "RPT"
Case "RST"
Case "CSO"
Case "NS"
Case "AMT"
Case "SPC"
Case "LFT"
Case "RGT"
Case "PUP"
Case "PDN"
End Select
End Sub
