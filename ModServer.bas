Attribute VB_Name = "Module5"
Public TerakhirCariKodeCust As String
Public TerakhirCariNamaCust As String

Public Cnstr As String
Public CNSTRPOINT As String
Public IDPOSDef As Long
Public IDGudangDef As Long
Public TipeCetakan As TipeCetak
Public IsLockDiHargaBeli As Boolean
Public IsPakaiKantong As Boolean
Public IsResetPerKasir As Boolean
Public Enum TipeCetak
  None = 0
  Preview_ = 1
  Optional_ = 2
  Print_ = 3
End Enum
Public DefCetakVoidSaatReset As Integer
Dim kal() As String 'serversales, uid,pwd,dbs_sales, IDPOS,IDGUDANG,serverPoint,uidpoint,pwd,dbs
                    '  0            1  2   3           4     5       6              7     8,  9
Public Sub bacaSettingServer()
  Dim i As Integer
  namafile = App.path & "\database\SettingServer.dat"
  Open namafile For Input As #1
  i = 0
  While Not EOF(1)
    ReDim Preserve kal(i + 1)
    Input #1, kal(i)
    i = i + 1
  Wend
  Close #1
  'Provider=SQLNCLI.1;Password=elliteserv;Persist Security Info=True;User ID=sa;Initial Catalog=DBCITYTOYS;Data Source=.
  Cnstr = "Provider=SQLOLEDB.1;Password=" & kal(2) & ";Persist Security Info=False;User ID=" & kal(1) & ";Initial Catalog=" & kal(3) & ";Data Source=" & kal(0) & ""
  CNSTRPOINT = "Provider=SQLOLEDB.1;Password=" & kal(8) & ";Persist Security Info=False;User ID=" & kal(7) & ";Initial Catalog=" & kal(9) & ";Data Source=" & kal(6) & ""
'  Cnstr = "Provider=SQLOLEDB;Data Source=" & kal(0) & ";initial Catalog=" & kal(3) & ";User ID=" & kal(1) & ";Password=" & kal(2) & ";Integrated Security=True;Connect Timeout=15"
'  CNSTRPOINT = "Provider=SQLOLEDB;Data Source=" & kal(6) & ";initial Catalog=" & kal(9) & ";User ID=" & kal(7) & ";Password=" & kal(8) & ";Integrated Security=True;Connect Timeout=15"
IDPOSDef = kal(4)
IDGudangDef = kal(5)
TipeCetakan = NullToNol(kal(11))
IsLockDiHargaBeli = NullToBool(kal(12))
IsPakaiKantong = NullToBool(kal(13))
IsResetPerKasir = NullToBool(kal(14))
End Sub
Public Function BolehReset(ByVal Shift As Integer, ByVal TANGGAL As String, ByVal tgl As Date) As Boolean
If Shift = 1 Or Shift = 2 Then
  Dim i As Integer
  Dim x As Date
  Dim SettingJam() As String
  Dim TglSettingReset As Date
  namafile = App.path & "\database\Setting.dat"
  Open namafile For Input As #1
  i = 0
  While Not EOF(1)
    ReDim Preserve SettingJam(i + 1)
    Input #1, SettingJam(i)
    i = i + 1
  Wend
    Close #1
    TglSettingReset = CDate(Mid(TANGGAL, 5, 4) & "/" & Mid(TANGGAL, 3, 2) & "/" & Mid(TANGGAL, 1, 2) & " " & SettingJam(Shift - 1)) 'BATAS RESET
  DefCetakVoidSaatReset = CLng(SettingJam(2))
    x = CDate(Format(Date, "yyyy/MM/dd") & " " & Format(tgl, "HH:mm"))
    If TglSettingReset <= x Then
        BolehReset = True
    Else
        BolehReset = False
    End If
  Else
    BolehReset = False
  End If
  End Function

Public Function BoolToInt(ByVal x As Boolean) As Integer
  If IsNull(x) Then
    BoolToInt = 0
  Else
    If x Then
      BoolToInt = 1
    Else
      BoolToInt = 0
    End If
  End If
End Function

Public Function FixKoma(ByVal x As Double) As String
  FixKoma = Replace(Format(x, "##0.000"), ",", ".")
End Function

Function GetNilaiPoinMember(ByVal IDMember As Long) As Long
  Dim Hasil As String
  On Error GoTo pesan
  bacaSettingServer
  Dim isOnline As Boolean
  Dim sqlcon As New ADODB.Connection
  Dim rs As New ADODB.Recordset
  If isRemcomendedOnline = True Then
    Set sqlcon = New ADODB.Connection
    sqlcon.ConnectionString = Cnstr
    sqlcon.Open
    Set rs = sqlcon.Execute("SELECT vSaldoPoin.SaldoPoin FROM vSaldoPoin WHERE vSaldoPoin.IDCustomer=" & IDMember)
    If rs.EOF Or rs.BOF Then
      Hasil = ""
    Else
      Hasil = NullToNol(rs(0).Value)
    End If
    rs.Close
    sqlcon.Close
    Set sqlcon = Nothing
    GetNilaiPoinMember = Hasil
    Exit Function
  End If
pesan:
  On Error Resume Next
  Set sqlcon = Nothing
  GetNilaiPoinMember = 0
End Function

Function ExecuteSQL(ByVal SQL As String) As String
    On Error GoTo pesan
    'On Error GoTo 0
    bacaSettingServer
    Dim isOnline As Boolean
    Dim sqlcon As New ADODB.Connection
    Set sqlcon = New ADODB.Connection
    sqlcon.ConnectionString = Cnstr
    sqlcon.Open
    sqlcon.Execute SQL
    sqlcon.Close
    Set sqlcon = Nothing
    ExecuteSQL = "ONLINE"
    Exit Function
pesan:
BuatLogApp ("error: " & Err.Description & vbCrLf & "SQl: " & SQL)
    On Error Resume Next
    On Error GoTo 0
    Set sqlcon = Nothing
    ExecuteSQL = "Local"
End Function

Function ExecuteSQLMEMBER(ByVal SQL As String) As String
    On Error GoTo pesan
    On Error GoTo 0
    bacaSettingServer
    Dim isOnline As Boolean
    Dim sqlcon As New ADODB.Connection
    Set sqlcon = New ADODB.Connection
    sqlcon.ConnectionString = CNSTRPOINT
    sqlcon.Open
    sqlcon.Execute SQL
    sqlcon.Close
    Set sqlcon = Nothing
    ExecuteSQLMEMBER = "ONLINE"
    Exit Function
pesan:
    On Error Resume Next
    On Error GoTo 0
    Set sqlcon = Nothing
    ExecuteSQLMEMBER = "Local"
End Function

Function ExecuteSkalarSQLMEMBER(ByVal SQL As String) As Long
Dim Hasil As Long
' CnstrMember = "Provider=SQLOLEDB.1;Password=" & "sahasystem" & ";Persist Security Info=True;User ID=" & "sa" & ";Initial Catalog=" & "Retail" & ";Data Source=" & "Xeon" & ""

    On Error GoTo pesan
    On Error GoTo 0
    bacaSettingServer
    Dim isOnline As Boolean
    Dim sqlcon As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Set sqlcon = New ADODB.Connection
    sqlcon.ConnectionString = CNSTRPOINT
    sqlcon.Open
    Set rs = sqlcon.Execute(SQL)
    If rs.EOF Or rs.BOF Then
      Hasil = 0
    Else
      Hasil = NullToNol(rs!Hasil)
    End If
    rs.Close
    sqlcon.Close
    Set sqlcon = Nothing
    ExecuteSkalarSQLMEMBER = Hasil
    Exit Function
pesan:
    On Error Resume Next
    On Error GoTo 0
    Set sqlcon = Nothing
    ExecuteSkalarSQLMEMBER = 0
End Function


Function ExecuteSkalarSQL(ByVal SQL As String) As String
Dim Hasil As String
' CnstrMember = "Provider=SQLOLEDB.1;Password=" & "sahasystem" & ";Persist Security Info=True;User ID=" & "sa" & ";Initial Catalog=" & "Retail" & ";Data Source=" & "Xeon" & ""

    On Error GoTo pesan
'    On Error GoTo 0
    bacaSettingServer
    Dim isOnline As Boolean
    Dim sqlcon As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Set sqlcon = New ADODB.Connection
    sqlcon.ConnectionString = Cnstr
    sqlcon.Open
    Set rs = sqlcon.Execute(SQL)
    If rs.EOF Or rs.BOF Then
      Hasil = ""
    Else
      Hasil = NullToStr(rs(0).Value)
    End If
    rs.Close
    sqlcon.Close
    Set sqlcon = Nothing
    ExecuteSkalarSQL = Hasil
    Exit Function
pesan:
    On Error Resume Next
'    On Error GoTo 0
    Set sqlcon = Nothing
    ExecuteSkalarSQL = ""
End Function

Public Sub KirimKeServer(ByVal IDSales As Long, ByVal tgl As Date)
    Dim ispending As Boolean
    ispending = False
    Dim penambahanPointSebelumnya As Long
    Dim SaldoKSBNotaIni As Long
    Dim NamaTabelSales As String
    Dim DiscNotaProsen As Double
    Dim DiscPersen1 As Double
    Dim IDWilayah As Long
    Dim i As Integer
    Dim DataLokal As Database
    Dim rstLokal As Recordset
'    Dim rstJual As Recordset
  On Error GoTo Trace
    If isRemcomendedOnline = True Then
      Dim nmfile As String
      Dim SQL As String
      Dim jumrec As Long
      Dim TANGGAL As String
      nmfile = App.path & "\database\TempDB" & Format(tgl, "_yyyyMM") & ".mdb"
      
      Set DataLokal = OpenDatabase(nmfile)
      Set rstLokal = DataLokal.OpenRecordset("Select MSales.* from MSales INNER JOIN MSalesD ON MSalesD.IDSales=MSales.NoID where UCASE(MSalesD.Transaksi)<>'AVD' AND MSales.NoID=" & IDSales)
      
      With rstLokal
        If rstLokal.EOF Or rstLokal.BOF Then
        Else
        TANGGAL = "convert(datetime,'" & Format(rstLokal!TANGGAL, "MM/dd/yyyy") & "',101)"
        'BO pakai AHS
        NamaTabelSales = "MJual"
        If NamaTabelSales = "MJual" Then
                If (rstLokal!SubTotal) <> 0 Then
                    DiscNotaProsen = rstLokal!DiscNota * 100 / rstLokal!SubTotal
                Else
                    DiscNotaProsen = 0
                End If
                IDSalesAHS = NullToNol(ExecuteSkalarSQL("Select Max(NoID) as Hasil From MJual")) + 1
                IDWilayah = NullToNol(ExecuteSkalarSQL("SELECT IDWilayah FROM MGudang WHERE NoID=" & IDGudangDef))
                SQL = "INSERT INTO MJual (IDGudang,IDWilayah,IsPOS,NoID,Kode,KodeReff,Tanggal,TanggalStock,JatuhTempo,"
                SQL = SQL & " IDCustomer,TanggalSJ,NoSJ,SubTotal,DiskonNotaProsen,DiskonNotaRp,DiskonNotaTotal,"
                SQL = SQL & " Biaya, Total, Bayar, Sisa,IDAdmin,IDPacking,Shift,NamaKasir,Pembulatan,IDBank,NoAcc,IDPOS,NoIDPOS,Kas,Bank,Voucher,Charge,TotalBKP," & _
                "DPP, NilaiPPN,IDJenisKartu,TasKresekA,TasKresekB,TasKresekC,TasKresekD," & _
                "BarangPoin, NilaiPoin , SisaPoin, IDReedemPoin, ReedemPoin, ReedemNilai)  VALUES (" & vbCrLf
                SQL = SQL & IDGudangDef & "," & IDWilayah & ","
                SQL = SQL & 1 & ","
                SQL = SQL & IDSalesAHS & ","
                SQL = SQL & "'" & Replace(rstLokal!kode & "/" & Format(IDPOSDef, "00") & "/" & Format(tgl, "yyMM"), "'", "''") & "',"
                SQL = SQL & "'" & Replace(rstLokal!kode, "'", "''") & "',"
                SQL = SQL & "'" & Format(rstLokal!TANGGAL, "yyyy-MM-dd HH:nn:ss") & "',"
                SQL = SQL & "'" & Format(rstLokal!TANGGAL, "yyyy-MM-dd HH:nn:ss") & "',"
                SQL = SQL & "'" & Format(DateAdd("d", 30, CDate(rstLokal!TANGGAL)), "yyyy-MM-dd HH:nn:ss") & "',"
                SQL = SQL & NullToNol(rstLokal!IDMember) & ","
                SQL = SQL & "'" & Format(rstLokal!TANGGAL, "yyyy-MM-dd HH:nn:ss") & "',"
                SQL = SQL & "'',"
                SQL = SQL & FixKoma(rstLokal!SubTotal) & ","
                SQL = SQL & FixKoma(0) & ","
                SQL = SQL & FixKoma(rstLokal!DiscIntern) & ","
                SQL = SQL & FixKoma(rstLokal!DiscIntern + NullToNol(rstLokal!Pembulatan)) & ","
                SQL = SQL & FixKoma(rstLokal!Hargatotal - rstLokal!SubTotal - rstLokal!Pembulatan - rstLokal!DiscIntern) & ","
                SQL = SQL & FixKoma(rstLokal!SubTotal - rstLokal!Pembulatan - rstLokal!DiscIntern) & ","
                SQL = SQL & FixKoma(rstLokal!UangMuka) & ","
                SQL = SQL & FixKoma(rstLokal!Hargatotal - rstLokal!UangMuka) & ","
                SQL = SQL & IDUser & ","
                SQL = SQL & -1 & "," & Replace(rstLokal!Shift, "'", "''") & ", '" & Replace(NamaKasir, "'", "''") & "'," & FixKoma(NullToNol(rstLokal!Pembulatan)) & "," & NullToNol(rstLokal!IDBank) & ",'" & Replace(NullToStr(rstLokal!NoAcc), "'", "''") & "'," & IDPOSDef & "," & NullToNol(rstLokal!NoID) & "," & _
                FixKoma(rstLokal!UangMuka - rstLokal!Voucher - rstLokal!Bank) & "," & FixKoma(rstLokal!Bank) & "," & FixKoma(rstLokal!Voucher) & "," & FixKoma(rstLokal!Charge) & "," & FixKoma(rstLokal!TotalBKP) & "," & _
                FixKoma(rstLokal!DPP) & "," & FixKoma(rstLokal!PPN) & "," & NullToNol(rstLokal!IDJenisKartu) & "," & NullToNol(rstLokal!TasKresekA) & "," & NullToNol(rstLokal!TasKresekB) & "," & NullToNol(rstLokal!TasKresekC) & "," & NullToNol(rstLokal!TasKresekD) & "," & _
                NullToNol(rstLokal!BarangPoin) & "," & NullToNol(rstLokal!NilaiPoin) & "," & NullToNol(rstLokal!SisaPoin) & ", " & FixKoma(NullToNol(rstLokal!IDReedemPoin)) & ", " & FixKoma(NullToNol(rstLokal!ReedemPoin)) & ", " & FixKoma(NullToNol(rstLokal!NilaiReedemPoin)) & ")"
       End If
          '999999: cara lama masih ada kemungkinan record sales dengan detil kosong kekirim yang berakibat fatal diserver
          If BoolToInt(rstLokal!ispending) = 1 Then
            ispending = True
          Else
            ispending = False
          End If
'            lbStatus.Caption = "Status : " & ExecuteSQL(sql)
'          End If
          DoEvents
        End If
        End With

        If ispending = False Then
        Dim IDJualD As Long
              Set rstLokal = DataLokal.OpenRecordset("Select * FROM MSalesd WHERE IDSales=" & IDSales & " ORDER BY NoID")
              With DataLokal
                  If rstLokal.EOF Or rstLokal.BOF Then
                  Else
                  '999999: Pindah disini
                   
                    Form3.lbStatus.Caption = "Status : " & ExecuteSQL(SQL)
                    rstLokal.MoveFirst
                    i = 1
                    Do While Not rstLokal.EOF
                         If NamaTabelSales = "MJual" Then
                          IDJualD = NullToNol(ExecuteSkalarSQL("Select Max(NoID) as Hasil From MJualD")) + 1
                          If rstLokal!HargaBruto <> 0 Then
                             DiscPersen1 = (rstLokal!DiscRp + rstLokal!DiscInternRp) * 100 / rstLokal!HargaBruto
                          Else
                             DiscPersen1 = 0
                          End If
                          
                          SQL = "INSERT INTO MJualD (NoID,IDJual,IDPackingD,NoUrut,Tgl,Jam,IDBarang,IDBarangD,IDSatuan,Qty,QtyPcs," & _
                          "Harga,HargaPcs,CTN,DiscPersen1,DiscPersen2,DiscPersen3,Disc1,Disc2,Disc3,Jumlah,Catatan,IDGudang," & _
                          "Konversi,IsPoin,IsPoinSupplier,IDPoinSupplier, BKP,Transaksi ) VALUES ("
                          SQL = SQL & IDJualD & ","
                          SQL = SQL & IDSalesAHS & ","
                          SQL = SQL & -1 & ","
                          SQL = SQL & i & ","
                          SQL = SQL & "GetDate(),"
                          SQL = SQL & "GetDate(),"
                          SQL = SQL & rstLokal!IdInventor & ","
                           SQL = SQL & rstLokal!IDInvSat & ","
                         SQL = SQL & rstLokal!idSatuan & ","
                          SQL = SQL & FixKoma(rstLokal!Qty) & ","
                          SQL = SQL & FixKoma(rstLokal!Qty * rstLokal!Konversi) & ","
                          SQL = SQL & FixKoma(rstLokal!HargaBruto) & ","
                          SQL = SQL & FixKoma(rstLokal!Jumlah / IIf(rstLokal!Qty = 0, 1, rstLokal!Qty) / IIf(rstLokal!Konversi = 0, 1, rstLokal!Konversi)) & ","
                          SQL = SQL & FixKoma(NullToNol(ExecuteSkalarSQL("Select " & rstLokal!Qty & "/MBarang.Ctn_Pcs*" & rstLokal!Konversi & " AS Hasil FROM MBarang WHERE NoID=" & rstLokal!IdInventor))) & ","
                          SQL = SQL & FixKoma(rstLokal!DiscInternProsen) & ","
                          SQL = SQL & FixKoma(0) & ","
                          SQL = SQL & FixKoma(0) & ","
                          SQL = SQL & FixKoma(rstLokal!DiscInternRp) & ","
                          SQL = SQL & FixKoma(0) & ","
                          SQL = SQL & FixKoma(0) & ","
                          SQL = SQL & FixKoma(rstLokal!Jumlah) & ","
                          If rstLokal!Transaksi = "PLU" Then
                            SQL = SQL & "'Penjualan POS',"
                          ElseIf rstLokal!Transaksi = "VOD" Then
                            SQL = SQL & "'Void POS',"
                          Else
                            SQL = SQL & "'Returan POS',"
                          End If
                          SQL = SQL & IDGudangDef & ","
                          SQL = SQL & FixKoma(NullToNol(rstLokal!Konversi)) & ","
                          SQL = SQL & NullToNol(rstLokal!IsPoin) & ","
                          SQL = SQL & NullToNol(rstLokal!IsPoinSupplier) & ","
                          SQL = SQL & NullToNol(rstLokal!IDPoinSupplier) & ","
                          SQL = SQL & NullToNol(rstLokal!BKP) & ",'"
                          SQL = SQL & NullToStr(rstLokal!Transaksi) & "'"
                          SQL = SQL & ")"
                          
                        End If
DoEvents
                    ExecuteSQL (SQL)
                    i = i + 1
                    DoEvents
DoEvents
                    rstLokal.MoveNext
                    Loop
                    DoEvents
                  End If
                End With
                DataLokal.Execute "Update MSales Set IsUpload=1, IsSelesai=1 where NoID=" & IDSales
          End If
    If IDMember > 0 Then
'        penambahanPointSebelumnya = ExecuteSkalarSQLMEMBER("select Sum(SaldoNotaIni) as hasil " & _
'        "from MCustomerPoint where IDCustomer=" & IDMember & " AND Tanggal=convert(datetime,'" & Format(Date, "MM/dd/yyyy") & "',101)")
'
'        SaldoKSBNotaIni = (BelanjaPoin - Voucher) - ((BelanjaPoin - Voucher + penambahanPointSebelumnya) \ 100000) * 100000
'
'        ExecuteSQLMEMBER "Insert Into MCustomerPoint(IDCustomer,KodeCustomer,Kode,Tanggal,Kassa,Bruto,Netto,Debet,IsKueBasah,PenambahanSebelumnya,SaldoNotaIni) Values(" & _
'                IDMember & ",'" & KodeMember & "','" & Format(IDSales, "0000000") & "'," & TANGGAL & ",'" & NamaMesin & "'," & FixKoma(Total) & "," & _
'                FixKoma(BelanjaPoin - Voucher) & "," & (BelanjaPoin - Voucher + penambahanPointSebelumnya) \ 100000 & "," & IIf(NamaMesin = "12" Or NamaMesin = "13", 1, 0) & "," & FixKoma(penambahanPointSebelumnya) & "," & FixKoma(SaldoKSBNotaIni) & ")"
    End If
   End If
Trace:
  If Err.Number <> 0 Then
    MsgBox Err.Number & "  " & Err.Description, vbCritical, App.Title
    Err.Clear
     BuatLog "Kesalahan saat eksekusi SQl:" & SQL & vbCrLf & "ERROR:" & Err.Description
  End If
End Sub

Public Sub KirimKeServerBeginTrans(ByVal IDSales As Long)
    Dim ispending As Boolean
    ispending = False
    Dim penambahanPointSebelumnya As Long
    Dim SaldoKSBNotaIni As Long
    Dim NamaTabelSales As String
    Dim DiscNotaProsen As Double
    Dim DiscPersen1 As Double
    Dim IDWilayah As Long
    Dim i As Integer
    Dim DataLokal As Database
    Dim rstLokal As Recordset
'    Dim rstJual As Recordset

  Dim con As ADODB.Connection
  Dim rs As ADODB.Recordset
  Dim Sukses As Boolean, Hasil As Integer
  Sukses = False
  Hasil = 0
  On Error GoTo Trace
    If isRemcomendedOnline = True Then
      Dim nmfile As String
      Dim SQL As String
      Dim jumrec As Long
      Dim TANGGAL As String
      nmfile = App.path & "\database\TempDB" & Format(Now, "_yyyyMM") & ".mdb"
      
      Set DataLokal = OpenDatabase(nmfile)
      Set rstLokal = DataLokal.OpenRecordset("Select MSales.* from MSales INNER JOIN MSalesD ON MSalesD.IDSales=MSales.NoID where UCASE(MSalesD.Transaksi)<>'AVD' AND MSales.NoID=" & IDSales)
      
      With rstLokal
        If rstLokal.EOF Or rstLokal.BOF Then
        Else
        TANGGAL = "convert(datetime,'" & Format(rstLokal!TANGGAL, "MM/dd/yyyy") & "',101)"
        'BO pakai AHS
        Set con = New ADODB.Connection
        con.ConnectionString = Cnstr
        con.Open 'Start Begin Tran
        Set rs = New ADODB.Recordset
        con.BeginTrans
'        Set rs = sqlcon.Execute(SQL)
'        If rs.EOF Or rs.BOF Then
'          hasil = ""
'        Else
'          hasil = NullToStr(rs(0).Value)
'        End If
    
        NamaTabelSales = "MJual"
        If NamaTabelSales = "MJual" Then
                If (rstLokal!SubTotal) <> 0 Then
                    DiscNotaProsen = rstLokal!DiscNota * 100 / rstLokal!SubTotal
                Else
                    DiscNotaProsen = 0
                End If
                Set rs = con.Execute("SELECT IDWilayah FROM MGudang WHERE NoID=" & IDGudangDef)
                If rs.EOF Or rs.BOF Then
                  IDWilayah = -1
                Else
                  IDWilayah = NullToNol(rs(0).Value)
                End If
                Set rs = con.Execute("Select Max(NoID) as Hasil From MJual")
                If rs.EOF Or rs.BOF Then
                  IDSalesAHS = 1
                Else
                  IDSalesAHS = NullToNol(rs(0).Value) + 1
                End If
                
                SQL = "INSERT INTO MJual (IDGudang,IDWilayah,IsPOS,NoID,Kode,KodeReff,Tanggal,TanggalStock,JatuhTempo,"
                SQL = SQL & " IDCustomer,TanggalSJ,NoSJ,SubTotal,DiskonNotaProsen,DiskonNotaRp,DiskonNotaTotal,"
                SQL = SQL & " Biaya, Total, Bayar, Sisa,IDAdmin,IDPacking,Shift,NamaKasir,Pembulatan,IDBank,NoAcc,IDPOS,NoIDPOS,Kas,Bank,Voucher,Charge,TotalBKP," & _
                     "DPP, NilaiPPN,IDJenisKartu,TasKresekA,TasKresekB,TasKresekC,TasKresekD," & _
                     "BarangPoin, NilaiPoin , SisaPoin,KodeMarketing,FeeMarketing,FeeMarketingRp, IDReedemPoin, ReedemPoin, ReedemNilai)  VALUES (" & vbCrLf
                SQL = SQL & IDGudangDef & "," & IDWilayah & ","
                SQL = SQL & 1 & ","
                SQL = SQL & IDSalesAHS & ","
                SQL = SQL & "'" & Replace(rstLokal!kode & "/" & Format(IDPOSDef, "00") & "/" & Format(tgl, "yyMM"), "'", "''") & "',"
                SQL = SQL & "'" & Replace(rstLokal!kode, "'", "''") & "',"
                SQL = SQL & "'" & Format(rstLokal!TANGGAL, "yyyy-MM-dd HH:nn:ss") & "',"
                SQL = SQL & "'" & Format(rstLokal!TANGGAL, "yyyy-MM-dd HH:nn:ss") & "',"
                SQL = SQL & "'" & Format(DateAdd("d", 30, CDate(rstLokal!TANGGAL)), "yyyy-MM-dd HH:nn:ss") & "',"
                SQL = SQL & NullToNol(rstLokal!IDMember) & ","
                SQL = SQL & "'" & Format(rstLokal!TANGGAL, "yyyy-MM-dd HH:nn:ss") & "',"
                SQL = SQL & "'',"
                SQL = SQL & FixKoma(rstLokal!SubTotal) & ","
                SQL = SQL & FixKoma(0) & ","
                SQL = SQL & FixKoma(rstLokal!DiscIntern) & ","
                SQL = SQL & FixKoma(rstLokal!DiscIntern + NullToNol(rstLokal!Pembulatan)) & ","
                SQL = SQL & FixKoma(rstLokal!Hargatotal - rstLokal!SubTotal - rstLokal!Pembulatan - rstLokal!DiscIntern) & ","
                SQL = SQL & FixKoma(rstLokal!SubTotal - rstLokal!Pembulatan - rstLokal!DiscIntern) & ","
                SQL = SQL & FixKoma(rstLokal!UangMuka) & ","
                SQL = SQL & FixKoma(rstLokal!Hargatotal - rstLokal!UangMuka) & ","
                SQL = SQL & IDUser & ","
                SQL = SQL & -1 & "," & Replace(rstLokal!Shift, "'", "''") & ", '" & Replace(NamaKasir, "'", "''") & "'," & FixKoma(NullToNol(rstLokal!Pembulatan)) & "," & NullToNol(rstLokal!IDBank) & ",'" & Replace(NullToStr(rstLokal!NoAcc), "'", "''") & "'," & IDPOSDef & "," & NullToNol(rstLokal!NoID) & "," & _
                FixKoma(rstLokal!UangMuka - rstLokal!Voucher - rstLokal!Bank) & "," & FixKoma(rstLokal!Bank) & "," & FixKoma(rstLokal!Voucher) & "," & FixKoma(rstLokal!Charge) & "," & FixKoma(rstLokal!TotalBKP) & "," & _
                FixKoma(rstLokal!DPP) & "," & FixKoma(rstLokal!PPN) & "," & NullToNol(rstLokal!IDJenisKartu) & "," & NullToNol(rstLokal!TasKresekA) & "," & NullToNol(rstLokal!TasKresekB) & "," & NullToNol(rstLokal!TasKresekC) & "," & NullToNol(rstLokal!TasKresekD) & "," & _
                NullToNol(rstLokal!BarangPoin) & "," & NullToNol(rstLokal!NilaiPoin) & "," & NullToNol(rstLokal!SisaPoin) & ",'" & _
                Replace(NullToStr(rstLokal!Sopir), "'", "''") & "'," & NullToNol(rstLokal!Komisi) & "," & NullToNol(rstLokal!KomisiRp) & ", " & FixKoma(NullToNol(rstLokal!IDReedemPoin)) & ", " & FixKoma(NullToNol(rstLokal!ReedemPoin)) & ", " & FixKoma(NullToNol(rstLokal!NilaiReedemPoin)) & ")"
       End If
          '999999: cara lama masih ada kemungkinan record sales dengan detil kosong kekirim yang berakibat fatal diserver
          If BoolToInt(rstLokal!ispending) = 1 Then
            ispending = True
          Else
            ispending = False
          End If
'            lbStatus.Caption = "Status : " & ExecuteSQL(sql)
'          End If
          DoEvents
        End If
        End With

        If ispending = False Then
        Dim IDJualD As Long
              Set rstLokal = DataLokal.OpenRecordset("Select * FROM MSalesd WHERE IDSales=" & IDSales & " ORDER BY NoID")
              With DataLokal
                  If rstLokal.EOF Or rstLokal.BOF Then
                  Else
                  '999999: Pindah disini
                    Sukses = False
                    con.Execute SQL, Hasil
                    If Hasil >= 1 Then 'Sukses Insertkan Header
                      Sukses = True
                      Form3.lbStatus.Caption = "Status : ONLINE"
                    Else
                      Sukses = False
                      Form3.lbStatus.Caption = "Status : LOCAL"
                    End If
                    If Sukses Then
                      rstLokal.MoveFirst
                      i = 1
                      Do While Not rstLokal.EOF
                           If NamaTabelSales = "MJual" Then
                            If rstLokal!HargaBruto <> 0 Then
                               DiscPersen1 = (rstLokal!DiscRp + rstLokal!DiscInternRp) * 100 / rstLokal!HargaBruto
                            Else
                               DiscPersen1 = 0
                            End If
                            Set rs = con.Execute("Select Max(NoID) as Hasil From MJualD")
                            If rs.EOF Or rs.BOF Then
                              IDJualD = 1
                            Else
                              IDJualD = NullToNol(rs(0).Value) + 1
                            End If
                            
                            SQL = "INSERT INTO MJualD (NoID,IDJual,IDPackingD,NoUrut,Tgl,Jam,IDBarang,IDBarangD,IDSatuan,Qty,QtyPcs," & _
                            "Harga,HargaPcs,HargaPokok,CTN,DiscPersen1,DiscPersen2,DiscPersen3,Disc1,Disc2,Disc3,Jumlah,Catatan,IDGudang," & _
                            "Konversi,IsPoin,IsPoinSupplier,IDPoinSupplier, BKP,Transaksi,IsPDP,IsDisc2,HargaNormal ) VALUES ("
                            SQL = SQL & IDJualD & ","
                            SQL = SQL & IDSalesAHS & ","
                            SQL = SQL & -1 & ","
                            SQL = SQL & i & ","
                            SQL = SQL & "GetDate(),"
                            SQL = SQL & "GetDate(),"
                            SQL = SQL & rstLokal!IdInventor & ","
                            SQL = SQL & rstLokal!IDInvSat & ","
                            SQL = SQL & rstLokal!idSatuan & ","
                            SQL = SQL & FixKoma(rstLokal!Qty) & ","
                            SQL = SQL & FixKoma(rstLokal!Qty * rstLokal!Konversi) & ","
                            SQL = SQL & FixKoma(rstLokal!HargaBruto) & ","
                            SQL = SQL & FixKoma(rstLokal!Jumlah / IIf(rstLokal!Qty = 0, 1, rstLokal!Qty) / IIf(rstLokal!Konversi = 0, 1, rstLokal!Konversi)) & ","
                            SQL = SQL & FixKoma(rstLokal!HargaPokok) & ","
                            SQL = SQL & FixKoma(0) & ","
                            SQL = SQL & FixKoma(rstLokal!DiscInternProsen) & ","
                            SQL = SQL & FixKoma(rstLokal!DiscProsen) & ","
                            SQL = SQL & FixKoma(0) & ","
                            SQL = SQL & FixKoma(rstLokal!DiscInternRp) & ","
                            SQL = SQL & FixKoma(rstLokal!DiscRp) & ","
                            SQL = SQL & FixKoma(0) & ","
                            SQL = SQL & FixKoma(rstLokal!Jumlah) & ","
                            If rstLokal!Transaksi = "PLU" Then
                              SQL = SQL & "'Penjualan POS',"
                            ElseIf rstLokal!Transaksi = "VOD" Then
                              SQL = SQL & "'Void POS',"
                            Else
                              SQL = SQL & "'Returan POS',"
                            End If
                            SQL = SQL & IDGudangDef & ","
                            SQL = SQL & FixKoma(NullToNol(rstLokal!Konversi)) & ","
                            SQL = SQL & NullToNol(rstLokal!IsPoin) & ","
                            SQL = SQL & NullToNol(rstLokal!IsPoinSupplier) & ","
                            SQL = SQL & NullToNol(rstLokal!IDPoinSupplier) & ","
                            SQL = SQL & NullToNol(rstLokal!BKP) & ",'"
                            SQL = SQL & NullToStr(rstLokal!Transaksi) & "',"
                            SQL = SQL & BoolToInt(NullToBool(rstLokal!IsPDP)) & ","
                            SQL = SQL & BoolToInt(NullToBool(rstLokal!IsDisc2)) & ","
                            SQL = SQL & FixKoma(NullToNol(rstLokal!HargaNormal)) & " "
                            SQL = SQL & ")"
                            
                          End If
                        DoEvents
                        Sukses = False
                        con.Execute SQL, Hasil
                        If Hasil >= 1 Then 'Sukses Insertkan Detil
                          Sukses = True
                          Form3.lbStatus.Caption = "Status : ONLINE"
                        Else
                          Sukses = False
                          Form3.lbStatus.Caption = "Status : LOCAL"
                          Exit Do
                        End If
                        DoEvents
                        i = i + 1
                        rstLokal.MoveNext
                      Loop
                    End If
                    DoEvents
                  End If
                End With
                If Sukses Then
                  con.CommitTrans
                  DataLokal.Execute "Update MSales Set IsUpload=1, IsSelesai=1 where NoID=" & IDSales
                Else
                  con.RollbackTrans
                End If
          End If
    If IDMember > 0 Then
'        penambahanPointSebelumnya = ExecuteSkalarSQLMEMBER("select Sum(SaldoNotaIni) as hasil " & _
'        "from MCustomerPoint where IDCustomer=" & IDMember & " AND Tanggal=convert(datetime,'" & Format(Date, "MM/dd/yyyy") & "',101)")
'
'        SaldoKSBNotaIni = (BelanjaPoin - Voucher) - ((BelanjaPoin - Voucher + penambahanPointSebelumnya) \ 100000) * 100000
'
'        ExecuteSQLMEMBER "Insert Into MCustomerPoint(IDCustomer,KodeCustomer,Kode,Tanggal,Kassa,Bruto,Netto,Debet,IsKueBasah,PenambahanSebelumnya,SaldoNotaIni) Values(" & _
'                IDMember & ",'" & KodeMember & "','" & Format(IDSales, "0000000") & "'," & TANGGAL & ",'" & NamaMesin & "'," & FixKoma(Total) & "," & _
'                FixKoma(BelanjaPoin - Voucher) & "," & (BelanjaPoin - Voucher + penambahanPointSebelumnya) \ 100000 & "," & IIf(NamaMesin = "12" Or NamaMesin = "13", 1, 0) & "," & FixKoma(penambahanPointSebelumnya) & "," & FixKoma(SaldoKSBNotaIni) & ")"
    End If
   End If
   Exit Sub
Trace:
  If Err.Number <> 0 Then
'    MsgBox Err.Number & "  " & Err.Description, vbCritical, App.Title
    BuatLog "Kesalahan saat eksekusi SQl:" & SQL & vbCrLf & "ERROR:" & Err.Description
    Err.Clear
    If Not con Is Nothing Then
      If con.State = adStateOpen Then
        con.RollbackTrans
      End If
    End If
  End If
  If Not rs Is Nothing Then
    rs.Close
    Set rs = Nothing
  End If
  If Not con Is Nothing Then
    con.Close
    Set con = Nothing
  End If
End Sub


