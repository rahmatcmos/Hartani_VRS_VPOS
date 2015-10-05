VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAdo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/*instance fields declaration*/
'/*Author by Yanto Hariyono*/

'/*Support database type*/
'/* 0 = MySQL
'/* 1 = Oracle
'/* 2 = Access
'/* 3 = PostgreSQL
'/* 4 = SQLServer
'/* 5 = ODBC

Public Enum dbType
    MySQL = 0
    Oracle = 1
    MsAccess = 2
    PostgreSQL = 3
    SQLServer = 4
    ODBC = 5
End Enum

Private m_con As New ADODB.Connection
Private m_com As New ADODB.Command
Private m_cons As New ADODB.Connection
Private m_coms As New ADODB.Command
Private rst As New ADODB.Recordset
Private rst2 As New ADODB.Recordset
Private rst3 As New ADODB.Recordset
Private SQL As String, SQLS As String
Public strconnection As String

Private Const ODBC_ADD_DSN = 1        ' Add data source
Private Const ODBC_CONFIG_DSN = 2     ' Configure (edit) data source
Private Const ODBC_REMOVE_DSN = 3     ' Remove data source
Private Const vbAPINull As Long = 0   ' NULL Pointer

'Function Declare
#If Win32 Then
    Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" _
    (ByVal hWndParent As Long, ByVal fRequest As Long, _
    ByVal lpszDriver As String, ByVal lpszAttributes As String) _
    As Long
#Else
    Private Declare Function SQLConfigDataSource Lib "ODBCINST.DLL" _
    (ByVal hWndParent As Integer, ByVal fRequest As Integer, ByVal _
    lpszDriver As String, ByVal lpszAttributes As String) As Integer
#End If

'Encrypt text
Public Function EncryptText(strText As String, ByVal strPwd As String)
    Dim i As Integer, c As Integer
    Dim strBuff As String

#If Not CASE_SENSITIVE_PASSWORD Then

    'Convert password to upper case
    'if not case-sensitive
    strPwd = UCase$(strPwd)

#End If

    'Encrypt string
    If Len(strPwd) Then
        For i = 1 To Len(strText)
            c = Asc(Mid$(strText, i, 1))
            c = c + Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    EncryptText = strBuff
End Function

'Decrypt text encrypted with EncryptText
Public Function DecryptText(strText As String, ByVal strPwd As String)
    Dim i As Integer, c As Integer
    Dim strBuff As String

#If Not CASE_SENSITIVE_PASSWORD Then

    'Convert password to upper case
    'if not case-sensitive
    strPwd = UCase$(strPwd)

#End If

    'Decrypt string
    If Len(strPwd) Then
        For i = 1 To Len(strText)
            c = Asc(Mid$(strText, i, 1))
            c = c - Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    DecryptText = strBuff
End Function

'/*MySQL*/
'/*--------------------------------------------------------------------------*/
Public Sub createmyodbc(ByVal pserver As String, ByVal pdatabase As String, _
                        ByVal puser As String, ByVal ppasswd As String, ByVal pport As Integer)
      #If Win32 Then
          Dim intRet As Long
      #Else
          Dim intRet As Integer
      #End If
      Dim strDriver As String
      Dim strAttributes As String
      
      strDriver = "MySQL ODBC 3.51 Driver"
      strAttributes = "SERVER=" & pserver & Chr$(0)
      strAttributes = strAttributes & "DESCRIPTION=Connection to MyODBC / MySQL" & Chr$(0)
      strAttributes = strAttributes & "DSN=TITMySQL" & Chr$(0)
      strAttributes = strAttributes & "DATABASE=" & pdatabase & Chr$(0)
      strAttributes = strAttributes & "USER=" & puser & Chr$(0)
      strAttributes = strAttributes & "PASSWORD=" & ppasswd & Chr$(0)
      strAttributes = strAttributes & "PORT=" & pport & Chr$(0)
      intRet = SQLConfigDataSource(vbAPINull, ODBC_ADD_DSN, strDriver, strAttributes)
      
      If intRet = 0 Then
          TulisPesan "gagal dibuat"
      End If
End Sub

Public Sub deletemyodbc()
      #If Win32 Then
          Dim intRet As Long
      #Else
          Dim intRet As Integer
      #End If
      Dim strDriver As String
      Dim strAttributes As String
      
      strDriver = "MySQL ODBC 5.1 Driver"
      strAttributes = strAttributes & "DSN=TITMySQL" & Chr$(0)
      intRet = SQLConfigDataSource(vbAPINull, ODBC_REMOVE_DSN, strDriver, strAttributes)
      
      If intRet = 0 Then
          TulisPesan "gagal dihapus"
      End If
End Sub
'/*--------------------------------------------------------------------------*/

Public Sub mykoneksi()
    On Error GoTo errkoneksi
    
    If m_con.State = adStateOpen Then
          m_con.Close
    End If
    
    
    Dim app_ini As String
    app_ini = App.Path & "\setting.ini"
    
    Dim dbtype_ As dbType
    dbtype_ = getstringinifiles("dbtype", "type", "", app_ini)
    
    Select Case dbtype_
        Case 0: m_con.ConnectionString = "DSN=TITMySQL"
        Case 1: m_con.ConnectionString = "Provider=OraOLEDB.Oracle.1;User ID=" & DecryptText(getstringinifiles("dbconfig", "user", "", app_ini), "sis2010") & ";" & _
                                         "Data Source=" & getstringinifiles("dbconfig", "dbname", "", app_ini) & ";" & _
                                         "Password=" & DecryptText(getstringinifiles("dbconfig", "pwd", "", app_ini), "sis2010")
                                         
        Case 2: m_con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & getstringinifiles("dbconfig", "dbname", "", app_ini) & ";JET OLEDB:DATABASE PASSWORD=" & Decrypt(getstringinifiles("dbconfig", "pwd", "", app_ini))
        
        Case 3: m_con.ConnectionString = "Driver={PostgreSQL ANSI};Server=" & getstringinifiles("dbconfig", "server", "", app_ini) & ";" & _
                                         "Uid=" & DecryptText(getstringinifiles("dbconfig", "user", "", app_ini), "sis2010") & ";Pwd=" & DecryptText(getstringinifiles("dbconfig", "pwd", "", app_ini), "sis2010") & ";" & _
                                         "Port=" & getstringinifiles("dbconfig", "port", "", app_ini) & ";Database=" & getstringinifiles("dbconfig", "dbname", "", app_ini)
'        Case 3: m_con.ConnectionString = "Provider=PostgreSQL OLE DB Provider;Password=" & DecryptText(getstringinifiles("dbconfig", "pwd", "", app_ini), "sis2010") & ";User ID=" & DecryptText(getstringinifiles("dbconfig", "user", "", app_ini), "sis2010") & ";Data Source=" & getstringinifiles("dbconfig", "server", "", app_ini) & ";Location=" & getstringinifiles("dbconfig", "dbname", "", app_ini) '& ";Port=" & getstringinifiles("dbconfig", "port", "", app_ini)
        Case 4: m_con.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Decrypt(getstringinifiles("dbconfig", "user", "", app_ini)) & ";pwd=" & Decrypt(getstringinifiles("dbconfig", "pwd", "", app_ini)) & ";Initial Catalog=" & getstringinifiles("dbconfig", "dbname", "", app_ini) & ";Data Source=" & getstringinifiles("dbconfig", "server", "", app_ini) & ";" & _
                                         "Port=" & getstringinifiles("dbconfig", "port", "1433", app_ini)
        Case 5: m_con.ConnectionString = "DSN=" & getstringinifiles("dbconfig", "dbname", "", app_ini)
    
    End Select
       
    strconnection = m_con.ConnectionString
    m_con.Open
          
errkoneksi:
    If Err <> 0 Then
      TulisPesan "[" & Err.Number & "] " & Err.Description & " Konfigurasi error!!!"
      
'      setstringinifiles "dbconfig", "server", "", app_ini
'      setstringinifiles "dbconfig", "dbname", "", app_ini
'      setstringinifiles "dbconfig", "user", "", app_ini
'      setstringinifiles "dbconfig", "pwd", "", app_ini
'      setstringinifiles "dbconfig", "port", "", app_ini
      
      m_con.Close
      Set m_con = Nothing
      Err.Clear
      Exit Sub
    End If
End Sub

Public Function teskoneksi() As Boolean
    On Error GoTo errkoneksi
    
    If m_con.State = adStateOpen Then
       m_con.Close
    End If
    
    Dim app_ini As String
    app_ini = App.Path & "\setting.ini"
    
    Dim dbtype_ As dbType
    dbtype_ = getstringinifiles("dbtype", "type", "", app_ini)
    
    Select Case dbtype_
       Case 0: LihatODBC: m_con.ConnectionString = "DSN=TITMySQL"
       Case 1: m_con.ConnectionString = "Provider=OraOLEDB.Oracle.1;User ID=" & DecryptText(getstringinifiles("dbconfig", "user", "", app_ini), "sis2010") & ";" & _
                                        "Data Source=" & getstringinifiles("dbconfig", "dbname", "", app_ini) & ";" & _
                                        "Password=" & DecryptText(getstringinifiles("dbconfig", "pwd", "", app_ini), "sis2010")
       Case 2: m_con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & getstringinifiles("dbconfig", "dbname", "", app_ini) & ";JET OLEDB:DATABASE PASSWORD=" & Decrypt(getstringinifiles("dbconfig", "pwd", "", app_ini))
       Case 3: m_con.ConnectionString = "Driver={PostgreSQL ANSI};Server=" & getstringinifiles("dbconfig", "server", "", app_ini) & ";" & _
                                        "Uid=" & DecryptText(getstringinifiles("dbconfig", "user", "", app_ini), "sis2010") & ";Pwd=" & DecryptText(getstringinifiles("dbconfig", "pwd", "", app_ini), "sis2010") & ";" & _
                                        "Port=" & getstringinifiles("dbconfig", "port", "", app_ini) & ";Database=" & getstringinifiles("dbconfig", "dbname", "", app_ini)
       Case 4: m_con.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Decrypt(getstringinifiles("dbconfig", "user", "", app_ini)) & ";pwd=" & Decrypt(getstringinifiles("dbconfig", "pwd", "", app_ini)) & ";Initial Catalog=" & getstringinifiles("dbconfig", "dbname", "", app_ini) & ";Data Source=" & getstringinifiles("dbconfig", "server", "", app_ini) & ";" & _
                                        "Port=" & getstringinifiles("dbconfig", "port", "", app_ini)
'       Case 4: m_con.ConnectionString = "Driver={SQLServer};Server=" & getstringinifiles("dbconfig", "server", "", app_ini) & ";" & _
                                        "Database=" & getstringinifiles("dbconfig", "dbname", "", app_ini) & ";" & _
                                        "Uid=" & DecryptText(getstringinifiles("dbconfig", "user", "", app_ini), "sis2010") & ";" & _
                                        "Pwd=" & DecryptText(getstringinifiles("dbconfig", "pwd", "", app_ini), "sis2010") & ";"
       Case 5: m_con.ConnectionString = "DSN=" & getstringinifiles("dbconfig", "dbname", "", app_ini)
    End Select
    
    m_con.Open
    teskoneksi = True
    'MsgBox "Koneksi tersambung.", vbInformation, "Konfigurasi berhasil"
         
errkoneksi:
    If Err <> 0 Then
      TulisPesan "Koneksi gagal!" & " Konfigurasi error!!!"
      Err.Clear
      
'      setstringinifiles "dbconfig", "server", "", app_ini
'      setstringinifiles "dbconfig", "dbname", "", app_ini
'      setstringinifiles "dbconfig", "user", "", app_ini
'      setstringinifiles "dbconfig", "pwd", "", app_ini
'      setstringinifiles "dbconfig", "port", "", app_ini
      setstringinifiles "basic", "config", "0", app_ini
      teskoneksi = False
      Exit Function
    End If
End Function

Private Sub LihatODBC()
    On Error Resume Next
    Dim app_ini As String
    app_ini = App.Path & "\setting.ini"
    
    #If Win32 Then
    Dim intRet As Long
    #Else
    Dim intRet As Integer
    #End If
    Dim strDriver As String
    Dim strAttributes As String
    
    strDriver = "MySQL ODBC 3.51 Driver"
    strAttributes = strAttributes & "DSN=TITMySQL" & Chr$(0)
    intRet = SQLConfigDataSource(vbAPINull, ODBC_REMOVE_DSN, strDriver, strAttributes)
    
    createmyodbc getstringinifiles("dbconfig", "server", "", app_ini), getstringinifiles("dbconfig", "dbname", "", app_ini), _
                DecryptText(getstringinifiles("dbconfig", "user", "", app_ini), "sis2010"), DecryptText(getstringinifiles("dbconfig", "pwd", "", app_ini), "sis2010"), _
                getstringinifiles("dbconfig", "port", "", app_ini)

End Sub

Public Function ExecuteQuery(ByVal SQL As String) As Boolean
  On Error GoTo errExec
  
  mykoneksi
  If rst.State = adStateOpen Then
    rst.Close
    Set rst = Nothing
  End If
  
  rst.Open SQL, m_con, adOpenDynamic, adLockOptimistic, adCmdText
  
  If Not rst.EOF Then
      ExecuteQuery = True
  Else
      ExecuteQuery = False
  End If
      
    
errExec:
    If Err <> 0 Then
        TulisPesan "[" & Err.Number & "] " & Err.Description
        Err.Clear
    End If
End Function

Public Function ExecuteQueryrst(ByVal SQL As String) As ADODB.Recordset
  On Error GoTo errExec
  
  mykoneksi
  If rst.State = adStateOpen Then
    rst.Close
    Set rst = Nothing
  End If
    
  rst.CursorLocation = adUseClient
  rst.Open SQL, m_con, adOpenDynamic, adLockOptimistic, adCmdText
 
  Set ExecuteQueryrst = rst

errExec:
    If Err <> 0 Then
        TulisPesan "[" & Err.Number & "] " & Err.Description
        Err.Clear
    End If
End Function

Public Function ExecuteQueryrstAdd(ByVal SQL As String) As ADODB.Recordset
  On Error GoTo errExec
  Dim rstadd As New ADODB.Recordset
  
  If m_con.State = adStateClosed Then
    mykoneksi
  End If
  
  If rstadd.State = adStateOpen Then
    rstadd.Close
    Set rstadd = Nothing
  End If
    
  rstadd.CursorLocation = adUseClient
  rstadd.Open SQL, m_con, adOpenDynamic, adLockOptimistic, adCmdText
 
  Set ExecuteQueryrstAdd = rstadd
errExec:
    If Err <> 0 Then
        TulisPesan "[" & Err.Number & "] " & Err.Description
        Err.Clear
    End If
End Function

Public Sub CloseConnection()
  On Error Resume Next
    m_con.Close
    Set m_con = Nothing
    Dim dbtype_ As dbType
    dbtype_ = getstringinifiles("dbtype", "type", "", app_ini)
    If dbtype_ = MsAccess Then
      CompactDB
    End If
End Sub

Public Sub ExecuteUpdate(ByVal SQL As String)
  On Error GoTo errExec
  
  mykoneksi
  If m_com.State = adStateOpen Then
     Set m_com = Nothing
  End If
  m_com.ActiveConnection = m_con
  m_com.CommandText = SQL
  m_com.CommandType = adCmdText
  m_com.Execute
  
errExec:
    If Err <> 0 Then
        TulisPesan "[" & Err.Number & "] " & Err.Description & " Update error!!!"
        Err.Clear
    End If
End Sub

'---------------------------------------------------------------------------
' PROCEDURE : CompactDB
' DESCRIPTION :
' RETURN VALUE: true or false depending wether there were errors or not
' NOTE :
' Add "Microsoft Jet and Replication Objects 2.x Library" to
' the references of your project
' EXAMPLE :
'---------------------------------------------------------------------------
'Public Function CompactDB(pFileName As String) As Boolean
Public Function CompactDB() As Boolean
On Error GoTo Trace
  Dim CONN As New JRO.JetEngine
  Dim ConnstringSorg As String, ConnstringDest As String
  Dim FileDatabase As String
  FileDatabase = getstringinifiles("dbconfig", "dbname", App.Path & "\Database\Order.Mdb", app_ini)
  
  ' Ensure file is not read only
  SetAttr FileDatabase, vbNormal
  ConnstringSorg = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FileDatabase & ";JET OLEDB:DATABASE PASSWORD=" & Decrypt(getstringinifiles("dbconfig", "pwd", "", app_ini))
  ConnstringDest = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\Temp.mdb" & ";JET OLEDB:DATABASE PASSWORD=" & Decrypt(getstringinifiles("dbconfig", "pwd", "", app_ini))
  
  Screen.MousePointer = vbHourglass
  CONN.CompactDatabase ConnstringSorg, ConnstringDest
  Screen.MousePointer = vbDefault
  
  'Copyfile.
  Kill FileDatabase
  FileCopy App.Path & "\Database\Temp.mdb", FileDatabase
  Kill App.Path & "\Database\Temp.mdb"
  
  Set CONN = Nothing
  CompactDB = True
  
Trace:
  If Err.Number <> 0 Then
    Screen.MousePointer = vbDefault
    Debug.Print Err.Description
    Err.Clear
  End If
End Function

