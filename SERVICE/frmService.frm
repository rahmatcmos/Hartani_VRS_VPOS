VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmService 
   Caption         =   "Service Database"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10605
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5310
   ScaleWidth      =   10605
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1Elastic 
      Height          =   5310
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10605
      _cx             =   18706
      _cy             =   9366
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   6
      MousePointer    =   0
      Version         =   801
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   0
      ChildSpacing    =   0
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   4
      GridCols        =   4
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   0
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmService.frx":0000
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.Timer Timer1 
         Left            =   2670
         Top             =   1020
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   5310
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   10605
         _ExtentX        =   18706
         _ExtentY        =   9366
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"frmService.frx":007C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LamaLoading As Integer
Dim JumlahData As Integer
Dim SQL As String
Dim rstList As New ADODB.Recordset
Dim m_con As New ADODB.Connection
Dim m_conMSSQL As New ADODB.Connection
Dim IsOnline As Boolean
Private Sub Form_Load()
  Timer1.Interval = 2000
  LamaLoading = 20
  Me.Caption = Me.Caption & " on " & getstringinifiles("dbconfig", "server", "", app_ini)
  Timer1.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If MsgBox("Ingin menutup aplikasi service ?", vbYesNo, App.Title) = vbNo Then
    Cancel = 1
  Else
    RichTextBox1.SaveFile App.Path & "\Report_" & Format(Now, "ddMMyyyy") & ".rtf"
  End If
End Sub

Private Function KirimKeKassa(ByVal NoID As Long) As Boolean
  On Error GoTo Trace
  Dim rstData As New ADODB.Recordset
  Dim rstCek As New ADODB.Recordset
  Dim Hasil As Boolean
  Hasil = False
  Dim i As Integer
  SQL = "SELECT * FROM MPOS"
  If NullToString(rstList.Fields("Tabel").Value) = "MBarang" Then
    'TINV
    SQL = "SELECT MBarangD.*, MBarang.Kode AS KodeBarang, MBarang.Nama AS NamaBarang, MSatuan.Kode AS Satuan, MBarang.Barcode, MBarang.HPP AS HargaPokok, " & vbCrLf & _
          " MBarang.DiscA1, MBarang.DiscA2, MBarang.DiscB1, MBarang.DiscB2, MBarang.DiscC1, MBarang.DiscC2, MBarang.DiscD1, MBarang.DiscD2, MBarang.DiscE1, MBarang.DiscE2, MBarang.DiscF1, MBarang.DiscF2 " & vbCrLf & _
          " FROM (MBarangD INNER JOIN MBarang ON MBarang.NoID=MBarangD.IDBarang) LEFT JOIN MSatuan ON MSatuan.NoID=MBarangD.IDSatuan" & vbCrLf & _
          " WHERE MBarangD.NoID=" & NullToLong(rstList.Fields("IDTabel").Value) & " AND MBarangD.IsActive=1 AND MBarang.IsActive=1 AND MBarangD.IsJualPOS=1"
    Set rstData = cOra.ExecuteQueryrstAdd(SQL)
    If Not (rstData.EOF Or rstData.BOF) Then
      With rstData
        Set rstCek = ExecuteQueryrstAddDBMaster("SELECT * FROM TINV WHERE NoID=" & NullToLong(rstData.Fields("NoID").Value), NullToString(rstList.Fields("PathDBTemp").Value) & "\Database\DBMaster.mdb")
        If (rstCek.BOF Or rstCek.EOF) Then
          SQL = "INSERT INTO TINV (NoID,IDInventor,Kode,Barcode,Nama,IDSatuan,KodeSat,Konversi,HargaJual,HargaPokok,HargaA,HargaB,HargaC,HargaD,HargaE,HargaF,HargaMinA,HargaMinB,HargaMinC,HargaMinD,HargaMinE,HargaMinF,DiscProsen1A,DiscProsen2A,DiscProsen1B,DiscProsen2B,DiscProsen1C,DiscProsen2C,DiscProsen1D,DiscProsen2D,DiscProsen1E,DiscProsen2E,DiscProsen1F,DiscProsen2F,DiscProsen,DiscRupiah,IsMember,IsOperator,IsOperator1,HargaMin) VALUES (" & vbCrLf
          SQL = SQL & NullToLong(.Fields("NoID").Value) & "," & vbCrLf
          SQL = SQL & NullToLong(.Fields("IDBarang").Value) & "," & vbCrLf
          SQL = SQL & "'" & FixApostropi(NullToString(.Fields("KodeBarang").Value)) & "'," & vbCrLf
          SQL = SQL & "'" & FixApostropi(NullToString(.Fields("Barcode").Value)) & "'," & vbCrLf
          SQL = SQL & "'" & FixApostropi(NullToString(.Fields("NamaBarang").Value)) & "'," & vbCrLf
          SQL = SQL & NullToLong(.Fields("IDSatuan").Value) & "," & vbCrLf
          SQL = SQL & "'" & FixApostropi(NullToString(.Fields("Satuan").Value)) & "'," & vbCrLf
          SQL = SQL & FixKoma(NullToDouble(.Fields("Konversi").Value)) & "," & vbCrLf
          SQL = SQL & FixKoma(NullToDouble(.Fields("HargaJualB").Value)) & "," & vbCrLf
          SQL = SQL & FixKoma(NullToDouble(.Fields("HargaPokok").Value)) & "," & vbCrLf
          SQL = SQL & FixKoma(NullToDouble(.Fields("HargaJualA").Value)) & "," & vbCrLf
          SQL = SQL & FixKoma(NullToDouble(.Fields("HargaJualB").Value)) & "," & vbCrLf
          SQL = SQL & FixKoma(NullToDouble(.Fields("HargaJualC").Value)) & "," & vbCrLf
          SQL = SQL & FixKoma(NullToDouble(.Fields("HargaJualD").Value)) & "," & vbCrLf
          SQL = SQL & FixKoma(NullToDouble(.Fields("HargaJualE").Value)) & "," & vbCrLf
          SQL = SQL & FixKoma(NullToDouble(.Fields("HargaJualF").Value)) & "," & vbCrLf
          SQL = SQL & FixKoma(NullToDouble(.Fields("HargaJualA").Value)) & "," & vbCrLf
          SQL = SQL & FixKoma(NullToDouble(.Fields("HargaJualB").Value)) & "," & vbCrLf
          SQL = SQL & FixKoma(NullToDouble(.Fields("HargaJualC").Value)) & "," & vbCrLf
          SQL = SQL & FixKoma(NullToDouble(.Fields("HargaJualD").Value)) & "," & vbCrLf
          SQL = SQL & FixKoma(NullToDouble(.Fields("HargaJualE").Value)) & "," & vbCrLf
          SQL = SQL & FixKoma(NullToDouble(.Fields("HargaJualF").Value)) & "," & vbCrLf
          SQL = SQL & FixKoma(NullToDouble(.Fields("DiscA1").Value)) & "," & vbCrLf
          SQL = SQL & FixKoma(NullToDouble(.Fields("DiscA2").Value)) & "," & vbCrLf
          SQL = SQL & FixKoma(NullToDouble(.Fields("DiscB1").Value)) & "," & vbCrLf
          SQL = SQL & FixKoma(NullToDouble(.Fields("DiscB2").Value)) & "," & vbCrLf
          SQL = SQL & FixKoma(NullToDouble(.Fields("DiscC1").Value)) & "," & vbCrLf
          SQL = SQL & FixKoma(NullToDouble(.Fields("DiscC2").Value)) & "," & vbCrLf
          SQL = SQL & FixKoma(NullToDouble(.Fields("DiscD1").Value)) & "," & vbCrLf
          SQL = SQL & FixKoma(NullToDouble(.Fields("DiscD2").Value)) & "," & vbCrLf
          SQL = SQL & FixKoma(NullToDouble(.Fields("DiscE1").Value)) & "," & vbCrLf
          SQL = SQL & FixKoma(NullToDouble(.Fields("DiscE2").Value)) & "," & vbCrLf
          SQL = SQL & FixKoma(NullToDouble(.Fields("DiscF1").Value)) & "," & vbCrLf
          SQL = SQL & FixKoma(NullToDouble(.Fields("DiscF2").Value)) & "," & vbCrLf
          SQL = SQL & FixKoma(0) & "," & vbCrLf
          SQL = SQL & FixKoma(0) & "," & vbCrLf
          SQL = SQL & FixKoma(0) & "," & vbCrLf
          SQL = SQL & FixKoma(0) & "," & vbCrLf
          SQL = SQL & FixKoma(0) & "," & vbCrLf
          SQL = SQL & FixKoma(NullToDouble(.Fields("HargaJualB").Value)) & ")"
        Else
          SQL = "UPDATE TINV SET " & vbCrLf
          SQL = SQL & " Kode='" & FixApostropi(NullToString(.Fields("KodeBarang").Value)) & "'," & vbCrLf
          SQL = SQL & " Barcode='" & FixApostropi(NullToString(.Fields("Barcode").Value)) & "'," & vbCrLf
          SQL = SQL & " Nama='" & FixApostropi(NullToString(.Fields("NamaBarang").Value)) & "'," & vbCrLf
          SQL = SQL & " IDSatuan=" & NullToLong(.Fields("IDSatuan").Value) & "," & vbCrLf
          SQL = SQL & " KodeSat='" & FixApostropi(NullToString(.Fields("Satuan").Value)) & "'," & vbCrLf
          SQL = SQL & " Konversi=" & FixKoma(NullToDouble(.Fields("Konversi").Value)) & "," & vbCrLf
          SQL = SQL & " HargaJual=" & FixKoma(NullToDouble(.Fields("HargaJualB").Value)) & "," & vbCrLf
          SQL = SQL & " HargaPokok=" & FixKoma(NullToDouble(.Fields("HargaPokok").Value)) & "," & vbCrLf
          SQL = SQL & " HargaA=" & FixKoma(NullToDouble(.Fields("HargaJualA").Value)) & "," & vbCrLf
          SQL = SQL & " HargaB=" & FixKoma(NullToDouble(.Fields("HargaJualB").Value)) & "," & vbCrLf
          SQL = SQL & " HargaC=" & FixKoma(NullToDouble(.Fields("HargaJualC").Value)) & "," & vbCrLf
          SQL = SQL & " HargaD=" & FixKoma(NullToDouble(.Fields("HargaJualD").Value)) & "," & vbCrLf
          SQL = SQL & " HargaE=" & FixKoma(NullToDouble(.Fields("HargaJualE").Value)) & "," & vbCrLf
          SQL = SQL & " HargaF=" & FixKoma(NullToDouble(.Fields("HargaJualF").Value)) & "," & vbCrLf
          SQL = SQL & " HargaMinA=" & FixKoma(NullToDouble(.Fields("HargaJualA").Value)) & "," & vbCrLf
          SQL = SQL & " HargaMinB=" & FixKoma(NullToDouble(.Fields("HargaJualB").Value)) & "," & vbCrLf
          SQL = SQL & " HargaMinC=" & FixKoma(NullToDouble(.Fields("HargaJualC").Value)) & "," & vbCrLf
          SQL = SQL & " HargaMinD=" & FixKoma(NullToDouble(.Fields("HargaJualD").Value)) & "," & vbCrLf
          SQL = SQL & " HargaMinE=" & FixKoma(NullToDouble(.Fields("HargaJualE").Value)) & "," & vbCrLf
          SQL = SQL & " HargaMinF=" & FixKoma(NullToDouble(.Fields("HargaJualF").Value)) & "," & vbCrLf
          SQL = SQL & " DiscProsen1A=" & FixKoma(NullToDouble(.Fields("DiscA1").Value)) & "," & vbCrLf
          SQL = SQL & " DiscProsen2A=" & FixKoma(NullToDouble(.Fields("DiscA2").Value)) & "," & vbCrLf
          SQL = SQL & " DiscProsen1B=" & FixKoma(NullToDouble(.Fields("DiscB1").Value)) & "," & vbCrLf
          SQL = SQL & " DiscProsen2B=" & FixKoma(NullToDouble(.Fields("DiscB2").Value)) & "," & vbCrLf
          SQL = SQL & " DiscProsen1C=" & FixKoma(NullToDouble(.Fields("DiscC1").Value)) & "," & vbCrLf
          SQL = SQL & " DiscProsen2C=" & FixKoma(NullToDouble(.Fields("DiscC2").Value)) & "," & vbCrLf
          SQL = SQL & " DiscProsen1D=" & FixKoma(NullToDouble(.Fields("DiscD1").Value)) & "," & vbCrLf
          SQL = SQL & " DiscProsen2D=" & FixKoma(NullToDouble(.Fields("DiscD2").Value)) & "," & vbCrLf
          SQL = SQL & " DiscProsen1E=" & FixKoma(NullToDouble(.Fields("DiscE1").Value)) & "," & vbCrLf
          SQL = SQL & " DiscProsen2E=" & FixKoma(NullToDouble(.Fields("DiscE2").Value)) & "," & vbCrLf
          SQL = SQL & " DiscProsen1F=" & FixKoma(NullToDouble(.Fields("DiscF1").Value)) & "," & vbCrLf
          SQL = SQL & " DiscProsen2F=" & FixKoma(NullToDouble(.Fields("DiscF2").Value)) & "," & vbCrLf
          SQL = SQL & " HargaMin=" & FixKoma(NullToDouble(.Fields("HargaJualB").Value)) & vbCrLf
          SQL = SQL & " WHERE NoID=" & NullToLong(.Fields("NoID").Value)
        End If
        TulisPesan "Memulai Kirim item barang Kode " & NullToString(.Fields("KodeBarang").Value) & " ke " & NullToString(rstList.Fields("NamaKassa").Value)
        If ExecuteUpdateDBMaster(SQL, NullToString(rstList.Fields("PathDBTemp").Value) & "\Database\DBMaster.mdb") Then
          TulisPesan "Berhasil Kirim item barang Kode " & NullToString(.Fields("KodeBarang").Value) & " ke " & NullToString(rstList.Fields("NamaKassa").Value)
          Hasil = True
        Else
          TulisPesan "Gagal Kirim item barang Kode " & NullToString(.Fields("KodeBarang").Value) & " ke " & NullToString(rstList.Fields("NamaKassa").Value)
        End If
        .MoveNext
      End With
    Else
      TulisPesan "Memulai Menghapus item barang NoID = " & NullToLong(rstList.Fields("IDTabel").Value) & " on " & NullToString(rstList.Fields("NamaKassa").Value)
      SQL = "DELETE FROM TINV WHERE NoID=" & NullToLong(rstList.Fields("IDTabel").Value)
      ExecuteUpdateDBMaster SQL, NullToString(rstList.Fields("PathDBTemp").Value) & "\Database\DBMaster.mdb"
      TulisPesan "Berhasil Menghapus item barang NoID = " & NullToLong(rstList.Fields("IDTabel").Value) & " on " & NullToString(rstList.Fields("NamaKassa").Value)
    End If
  ElseIf NullToString(rstList.Fields("Tabel").Value) = "MCustomer" Then
    'MCustomer
    SQL = "SELECT MAlamat.*, MJenisHarga.Kode AS Gol " & _
          " FROM MAlamat LEFT JOIN MJenisHarga ON MJenisHarga.NoID=MAlamat.DefaultTipeHarga " & _
          " WHERE MAlamat.NoID=" & NullToLong(rstList.Fields("IDTabel").Value) & " AND MAlamat.IsCustomer=1 And MAlamat.IsActive = 1 "
      Set rstData = cOra.ExecuteQueryrstAdd(SQL)
      If (rstData.EOF Or rstData.BOF) Then
        With rstData
          Set rstCek = ExecuteQueryrstAddDBMaster("SELECT * FROM MCustomer WHERE NoID=" & NullToLong(rstData.Fields("NoID").Value), NullToString(rstList.Fields("PathDBTemp").Value) & "\Database\DBMaster.mdb")
          If Not (rstCek.BOF Or rstCek.EOF) Then
            SQL = "INSERT INTO MCustomer ([NoID],[Kode],[Barcode],[Nama],[DOJ],[LimitHutang],[TipeHargaJual],[Alamat],[GolonganHarga]) VALUES (" & vbCrLf & _
                  NullToLong(.Fields("NoID").Value) & "," & vbCrLf & _
                  "'" & FixApostropi(NullToString(.Fields("Kode").Value)) & "'," & vbCrLf & _
                  "'" & FixApostropi(NullToLong(.Fields("NoID").Value)) & "'," & vbCrLf & _
                  "'" & FixApostropi(NullToString(.Fields("Nama").Value)) & "'," & vbCrLf & _
                  "#" & Format(NullToDate(.Fields("DOJ").Value), "yyyy/MM/dd") & "#," & vbCrLf & _
                  NullToDouble(.Fields("LimitPiutang").Value) & "," & NullToLong(.Fields("DefaultTipeHarga").Value) & ",'" & FixApostropi(NullToString(.Fields("Alamat").Value)) & "','" & FixApostropi(NullToString(.Fields("Gol").Value)) & "')"
          Else
            SQL = "UPDATE MCustomer SET " & vbCrLf & _
                  " [Kode]='" & FixApostropi(NullToString(.Fields("Kode").Value)) & "'," & vbCrLf & _
                  " [Barcode]='" & FixApostropi(NullToLong(.Fields("NoID").Value)) & "'," & vbCrLf & _
                  " [Nama]='" & FixApostropi(NullToString(.Fields("Nama").Value)) & "'," & vbCrLf & _
                  " [DOJ]=#" & Format(NullToDate(.Fields("DOJ").Value), "yyyy/MM/dd") & "#," & vbCrLf & _
                  " [LimitHutang]=" & NullToDouble(.Fields("LimitPiutang").Value) & "," & _
                  " [TipeHargaJual]=" & NullToLong(.Fields("DefaultTipeHarga").Value) & "," & _
                  " [Alamat]='" & FixApostropi(NullToString(.Fields("Alamat").Value)) & "'," & _
                  " [GolonganHarga]='" & FixApostropi(NullToString(.Fields("Gol").Value)) & "' " & vbCrLf & _
                  " WHERE NoID=" & NullToLong(.Fields("NoID").Value)
          End If
          TulisPesan "Memulai Kirim Data Customer Kode " & NullToString(.Fields("Kode").Value) & " ke " & NullToString(rstList.Fields("NamaKassa").Value)
          If ExecuteUpdateDBMaster(SQL, NullToString(rstList.Fields("PathDBTemp").Value) & "\Database\DBMaster.mdb") Then
            TulisPesan "Berhasil Kirim Data Customer Kode " & NullToString(.Fields("Kode").Value) & " ke " & NullToString(rstList.Fields("NamaKassa").Value)
            Hasil = True
          Else
            TulisPesan "Gagal Kirim Data Customer Kode " & NullToString(.Fields("Kode").Value) & " ke " & NullToString(rstList.Fields("NamaKassa").Value)
          End If
          .MoveNext
        End With
      Else
        TulisPesan "Memulai Menghapus Data Customer NoID = " & NullToLong(rstList.Fields("IDTabel").Value) & " on " & NullToString(rstList.Fields("NamaKassa").Value)
        SQL = "DELETE FROM MCustomer WHERE NoID=" & NullToLong(rstList.Fields("IDTabel").Value)
        ExecuteUpdateDBMaster SQL, NullToString(rstList.Fields("PathDBTemp").Value) & "\Database\DBMaster.mdb"
        TulisPesan "Berhasil Menghapus Data Customer NoID = " & NullToLong(rstList.Fields("IDTabel").Value) & " on " & NullToString(rstList.Fields("NamaKassa").Value)
      End If
  End If
  rstList.MoveNext
  JumlahData = JumlahData - 1
Trace:
  If Err.Number <> 0 Then
    TulisPesan "Kesalahan : " & Err.Number & " " & Err.Description
    Err.Clear
  End If
  
On Error Resume Next
  If rstCek.State = adStateOpen Then
    rstCek.Close
    Set rstCek = Nothing
  End If
  If rstData.State = adStateOpen Then
    rstData.Close
    Set rstData = Nothing
  End If
  KirimKeKassa = Hasil
End Function
Private Function ExecuteQueryrstAddDBMaster(ByVal SQL As String, ByVal Path As String) As ADODB.Recordset
  Dim rstadd As New ADODB.Recordset
  On Error GoTo errExec
  
  If m_con.State = adStateOpen Then
  Else
    m_con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Path
    m_con.Open
  End If
  rstadd.CursorLocation = adUseClient
  rstadd.Open SQL, m_con, adOpenDynamic, adLockOptimistic, adCmdText
 
  Set ExecuteQueryrstAddDBMaster = rstadd
  
'  m_con.Close
'  Set m_con = Nothing
errExec:
    If Err <> 0 Then
        TulisPesan "[" & Err.Number & "] " & Err.Description
        Err.Clear
    End If
End Function
Private Function ExecuteUpdateDBMaster(ByVal SQL As String, ByVal Path As String) As Boolean
  Dim m_com As New ADODB.Command
  Dim Hasil As Boolean
  On Error GoTo errExec
  
  If m_con.State = adStateOpen Then
  Else
    m_con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Path
    m_con.Open
  End If
  m_com.ActiveConnection = m_con
  m_com.CommandText = SQL
  m_com.CommandType = adCmdText
  m_com.Execute
  Hasil = True
errExec:
    If Err <> 0 Then
        TulisPesan "[" & Err.Number & "] " & Err.Description
        Err.Clear
        Hasil = False
    End If
ExecuteUpdateDBMaster = Hasil
End Function

Private Sub Timer1_Timer()
On Error GoTo Trace
  Dim NoID As Long
  If LamaLoading >= 1 And LamaLoading < 15 Then
    If LamaLoading = 0 And JumlahData = 0 Then
      TulisPesan "Tutup Koneksi End Transaction"
    Else
      If JumlahData > 0 Then
        NoID = NullToLong(rstList.Fields("NoID").Value)
        If KirimKeKassa(JumlahData) Then
          cOra.ExecuteUpdate "DELETE FROM MDBMasterUpdate WHERE NoID=" & NoID
        End If
      End If
    End If
    LamaLoading = LamaLoading + 1
  ElseIf LamaLoading >= 15 Then
    TulisPesan "Membaca Server"
    SQL = "SELECT MDBMasterUpdate.*, MPOS.PathDBTemp, MPOS.Nama AS NamaKassa FROM MDBMasterUpdate LEFT JOIN MPOS ON MPOS.NoID=MDBMasterUpdate.IDKassa ORDER BY TanggalModified DESC"
    Set rstList = ExecuteQueryrstAddSQLServer(SQL)
    If Not (rstList.BOF Or rstList.EOF) Then
      JumlahData = rstList.RecordCount
    Else
      JumlahData = 0
    End If
    TulisPesan "Jumlah data update " & JumlahData & ""
    LamaLoading = 0
  Else
    LamaLoading = LamaLoading + 1
  End If
Trace:
  If Err.Number <> 0 Then
    TulisPesan "Kesalahan : " & Err.Number & " " & Err.Description
    Err.Clear
  End If
End Sub

Public Function ExecuteQueryrstAddSQLServer(ByVal SQL As String) As ADODB.Recordset
  On Error GoTo errExec
  Dim rstadd As New ADODB.Recordset
  DoEvents
  If m_conMSSQL.State = adStateClosed Then
    m_conMSSQL.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Decrypt(getstringinifiles("dbconfig", "user", "", app_ini)) & ";pwd=" & Decrypt(getstringinifiles("dbconfig", "pwd", "", app_ini)) & ";Initial Catalog=" & getstringinifiles("dbconfig", "dbname", "", app_ini) & ";Data Source=" & getstringinifiles("dbconfig", "server", "", app_ini) & ";" & _
                                  "Port=" & getstringinifiles("dbconfig", "port", "1433", app_ini)
    m_conMSSQL.Open
    IsOnline = True
  End If
    
  If rstadd.State = adStateOpen Then
    rstadd.Close
    Set rstadd = Nothing
  End If
    
  rstadd.CursorLocation = adUseClient
  rstadd.Open SQL, m_conMSSQL, adOpenDynamic, adLockOptimistic, adCmdText
 
  Set ExecuteQueryrstAddSQLServer = rstadd
errExec:
    If Err <> 0 Then
        TulisPesan "[" & Err.Number & "] " & Err.Description
        Err.Clear
        m_conMSSQL.Close
        IsOnline = False
    End If
End Function
