Attribute VB_Name = "Function"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
         ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
         
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, _
         ByVal lpFileName As String) As Long
         
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
         ByVal lpFileName As String) As Long
         
Dim buffer As String * 255
Dim str As String
Public cOra As New clsAdo

Public Function getstringinifiles(ByVal psection As String, ByVal pkey As String, _
                                  ByVal pdefault As String, ByVal ppath As String) As String
  Dim X As Long
  
  X = GetPrivateProfileString(psection, pkey, pdefault, buffer, 255, ppath)
  str = Left(buffer, X)
  
  If Trim(str) <> "" Then
     getstringinifiles = str
  End If
End Function

Public Sub TulisPesan(ByVal str As String)
  frmService.RichTextBox1.Text = str & " ... " & Format(Now, "dd/MM/yyyy HH:mm:ss") & vbCrLf & frmService.RichTextBox1.Text
End Sub

Public Function getintinifiles(ByVal psection As String, ByVal pkey As String, _
                               ByVal pdefault As String, ByVal ppath As String) As Integer
  Dim xint As Long
  xint = GetPrivateProfileInt(psection, pkey, pdefault, ppath)
  getintinifiles = xint
End Function

Public Sub setstringinifiles(ByVal psection As String, ByVal pkey As String, _
                             ByVal pValue As String, ByVal ppath As String)
  WritePrivateProfileString psection, pkey, pValue, ppath
End Sub

Public Function FixApostropi(ByVal str As String) As String
  FixApostropi = Replace(str, "'", "''")
End Function

Public Function FixKoma(ByVal dbl As Double) As Double
  FixKoma = Replace(dbl, ",", ".")
End Function

Public Function SetDate(ByVal dt As Date) As String
  SetDate = Format(dt, "yyyy-MM-dd")
End Function

Public Function NullToString(X) As String
  NullToString = IIf(IsNull(X), "", X)
End Function

Public Function NullToDouble(ByVal X) As Double
  If (IsNull(X) Or Trim(X) = "") Then NullToDouble = 0 Else NullToDouble = X
End Function

Public Function NullToDate(ByVal X) As Date
  If (IsNull(X) Or Trim(X) = "") Then NullToDate = #1/1/1900# Else NullToDate = CDate(X)
End Function

Public Function NullToBoolean(X) As Integer
  If IsNull(X) Then
    NullToBoolean = 0
  Else
    If X Then
      NullToBoolean = 1
    Else
      NullToBoolean = 0
    End If
  End If
End Function

Public Function NullToLong(X) As Long
  If IsNull(X) Then
    NullToLong = 0
  ElseIf Trim(X) = "" Then
    NullToLong = 0
  Else
    NullToLong = CLng(X)
  End If
End Function

Public Function GetNewKodeByFilter(ByVal nmTable As String, ByVal nmfield As String, _
ByVal intStart As Integer, ByVal intLong As Integer, ByVal Filter As String) As String
On Error GoTo Trace

Dim SQL As String
Dim cOra As New clsAdo
Dim rst As New ADODB.Recordset
  SQL = "SELECT STR(MAX(MID(" & nmTable & "." & nmfield & "," & intStart & "," & intLong & "))) AS NEWKODE FROM " & nmTable & vbCrLf
  SQL = SQL & Filter
  
  Set rst = cOra.ExecuteQueryrst(SQL)
  If rst.EOF Or rst.BOF Then
    GetNewKodeByFilter = 1
  Else
    GetNewKodeByFilter = CLng(NullToLong(rst.Fields("NEWKODE"))) + 1
  End If
Trace:
  If Err.Number = 0 Then
    cOra.CloseConnection
  Else
    TulisPesan Err.Description
    GetNewKodeByFilter = 1
    Err.Clear
  End If
End Function

Public Function Encrypt(ByVal str As String) As String
  Encrypt = cOra.EncryptText(str, "Yanto")
End Function

Public Function Decrypt(ByVal str As String) As String
  Decrypt = cOra.DecryptText(str, "Yanto")
End Function
