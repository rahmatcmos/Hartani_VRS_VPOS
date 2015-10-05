Attribute VB_Name = "mdlIni"
'Attribute VB_Name = "mdlsys"
Option Explicit
'----------------------------------------
'/*Icon  notify
'/*Read and Write *.ini files*/
'/*Author by yanto hariyono*/
'Date = 13 - 05 - 2009 / Surabaya
'----------------------------------------

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias " Shell_NotifyIconA" (ByVal dwMessage As Long _
        , lpData As NOTIFYICONDATA) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
         ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
         
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, _
         ByVal lpFileName As String) As Long
         
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
         ByVal lpFileName As String) As Long

Private buffer As String * 255
Private str As String

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

#Const CASE_SENSITIVE_PASSWORD = False
Private Const NIM_ADD = &H0
Private Const NIM_DELETE = &H2
Private Const NIM_MODIFY = &H1
Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4
Private Const NIF_DOALL = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_RBUTTONDOWN = &H204


Public Function getstringinifiles(ByVal psection As String, ByVal pkey As String, _
                                  ByVal pdefault As String, ByVal ppath As String) As String
  Dim x As Long
  
  x = GetPrivateProfileString(psection, pkey, pdefault, buffer, 255, ppath)
  str = Left(buffer, x)
  
  If Trim(str) <> "" Then
     getstringinifiles = str
  End If
End Function

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

Public Sub createIcon(ByVal iconNotify As PictureBox)
Dim Tic As NOTIFYICONDATA
Dim erg As Long

Tic.cbSize = Len(Tic)
Tic.hwnd = iconNotify.hwnd
Tic.uID = 1&
Tic.uFlags = NIF_DOALL
Tic.uCallbackMessage = WM_MOUSEMOVE
Tic.hIcon = iconNotify.Picture
Tic.szTip = "V Point Developer"
erg = Shell_NotifyIcon(NIM_ADD, Tic)

If erg = 0 Then
    MsgBox "tray gagal dibuat", vbExclamation
End If
End Sub

Public Sub deleteIcon(ByVal iconNotify As PictureBox)
Dim Tic As NOTIFYICONDATA
Dim erg As Long

Tic.cbSize = Len(Tic)
Tic.hwnd = iconNotify.hwnd
Tic.uID = 1&
erg = Shell_NotifyIcon(NIM_DELETE, Tic)

If erg = 0 Then
    MsgBox "tray gagal dihapus", vbExclamation
End If
End Sub



