VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSerial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim clsSembunyikan As New CCrypto
Dim IDHarware As String

Public Enum Hasilnya
  HarusDitutup = 0
  Trial = 1
  Aktif = 2
End Enum
Public HasilX As Hasilnya
Public contactPerson As String

Public Function ActivationKeys(ByVal IDHardware As String, IDSoftware As String) As String
  On Error GoTo Trace
  'Combinasi Kanan Kiri
  Dim tmp As String, i As Integer
  tmp = ""
  For i = 1 To 25
    tmp = Mid(IDHardware, i, 1) & Mid(IDSoftware, 25 - i, 1)
  Next
  ActivationKeys = clsSembunyikan.Encrypt(Mid(tmp, 1, 25), Mid(tmp, 25, 25), True, frezBlockEncryption)
Trace:
  If Err.Number <> 0 Then
    MsgBox "Kesalahan : " & Err.Number & " - " & Err.Description, vbCritical, App.Title
    Err.Clear
  End If
End Function

Public Function CekCPUTrial(ByVal SerialSoftware25Digit As String) As Boolean
  On Error GoTo Trace
  Dim StrHasil As String, i As Integer
  Dim xActivasi As String
  contactPerson = "Email : vpointindonesia@yahoo.co.id, Telp : 031-81111918, Cell : 087862405489"
  If Len(SerialSoftware25Digit) <> 25 Then
    MsgBox "Invalid ID Software." & vbCrLf & "Aplikasi akan ditutup.", vbCritical, App.Title
    HasilX = HarusDitutup
    CekCPUTrial = False
    Exit Function
  End If
  StrHasil = MBSerialNumber & GetCpuID
  If Len(Replace(StrHasil, " ", "")) <> 25 Then
    For i = 1 To 5
      If i = 1 Then
        IDHarware = Left(Replace(StrHasil, " ", ""), 5)
      ElseIf i = 2 Then
        IDHarware = IDHarware & Right(Replace(StrHasil, " ", ""), 5)
      ElseIf i = 3 Then
        IDHarware = IDHarware & Mid(Replace(StrHasil, " ", ""), IIf(Len(Replace(StrHasil, " ", "")) / 2 >= 5, Len(Replace(StrHasil, " ", "")) / 2, Len("2ERT4")), 5)
      ElseIf i = 4 Then
        IDHarware = IDHarware & Mid(Replace(StrHasil, " ", ""), IIf(Len(Replace(StrHasil, " ", "")) / 3 >= 5, Len(Replace(StrHasil, " ", "")) / 3, Len("56YH9")), 5)
      Else
        IDHarware = IDHarware & Mid(Replace(StrHasil, " ", ""), IIf(Len(Replace(StrHasil, " ", "")) / 4 >= 5, Len(Replace(StrHasil, " ", "")) / 4, Len("90UHJ")), 5)
      End If
    Next
  End If
  StrHasil = clsSembunyikan.Encrypt(IDHarware, SerialSoftware25Digit, True, frezBlockEncryption)
  If BacaFile <> StrHasil Then
    If Now > CDate("2015-12-01") Then
      MsgBox "Aplikasi Trial sudah Habis." & vbCrLf & _
      "Aplikasi akan ditutup." & _
      "Hubungi Customer service kami di " & contactPerson & vbCrLf & _
      "Segeralah untuk registrasi.", vbCritical, App.Title
      HasilX = HarusDitutup
      CekCPUTrial = False
      Exit Function
    ElseIf Now = CDate("2015-12-01") Then
      xActivasi = StrHasil
      frmActivation.Tampil IDHarware, SerialSoftware25Digit, xActivasi
      If xActivasi <> StrHasil Then
        MsgBox "Hari Terakhir Aplikasi Trial." & vbCrLf & _
        "Hubungi Customer service kami di " & contactPerson & vbCrLf & _
        "Segeralah untuk registrasi.", vbCritical, App.Title
        HasilX = Trial
        CekCPUTrial = False
        Exit Function
      Else
        TulisSerial xActivasi
        HasilX = Aktif
        CekCPUTrial = True
        Exit Function
      End If
    Else
      xActivasi = StrHasil
      frmActivation.Tampil IDHarware, SerialSoftware25Digit, xActivasi
      If xActivasi <> StrHasil Then
        MsgBox "Aplikasi Trial diaktifkan." & vbCrLf & _
        "Hubungi Customer service kami di " & contactPerson & vbCrLf & _
        "Segeralah untuk registrasi.", vbCritical, App.Title
        HasilX = Trial
        CekCPUTrial = False
        Exit Function
      Else
        TulisSerial xActivasi
        HasilX = Aktif
        CekCPUTrial = True
        Exit Function
      End If
    End If
  Else
    HasilX = Aktif
    CekCPUTrial = True
    Exit Function
  End If
Trace:
If Err.Number <> 0 Then
  MsgBox "Kesalahan : " & Err.Number & " - " & Err.Description, vbCritical, App.Title
  HasilX = Trial
  CekCPUTrial = False
  Err.Clear
End If
End Function
Public Sub TulisSerial(ByVal str As String)
On Error GoTo Trace
'  Open App.path & "\Keys.dll" For Output As #1
'  Print #1, clsSembunyikan.Encrypt(str, "SGI_Framework", False, frezStreamEncryption)
'  Close #1
  setstringinifiles "Application", "Keys", str, App.path & "\Keys.dll"
Trace:
  If Err.Number <> 0 Then
    MsgBox "Kesalahan : " & Err.Number & " - " & Err.Description, vbCritical, App.Title
    Err.Clear
  End If
End Sub
Public Function BacaFile() As String
On Error GoTo Trace
If Dir(App.path & "\Keys.dll") <> "" Then
'  Dim i As Integer
'  Dim kal() As String
'  Open App.path & "\Keys.dll" For Input As #1
'  i = 0
'  While Not EOF(1)
'    ReDim Preserve kal(i + 1)
'    Input #1, kal(i)
'    i = i + 1
'  Wend
'  Close #1
 ' BacaFile = clsSembunyikan.Decrypt(getstringinifiles("Application", "Keys", "0000000000000000000000000", App.path & "\Keys.dll"), "SGI_Framework", False, frezStreamEncryption) 'clsSembunyikan.Decrypt(kal(0), "SGI_Framework", False, frezStreamEncryption)
 BacaFile = getstringinifiles("Application", "Keys", "0000000000000000000000000", App.path & "\Keys.dll")
Else
  BacaFile = "0000000000000000000000000"
End If
Trace:
If Err.Number <> 0 Then
  MsgBox "Kesalahan : " & Err.Number & " - " & Err.Description, vbCritical, App.Title
  Err.Clear
End If
End Function

Public Function MBSerialNumber() As String
 
'RETRIEVES SERIAL NUMBER OF MOTHERBOARD
'IF THERE IS MORE THAN ONE MOTHERBOARD, THE SERIAL
'NUMBERS WILL BE DELIMITED BY COMMAS

'YOU MUST HAVE WMI INSTALLED AND A REFERENCE TO
'Microsoft WMI Scripting Library IS REQUIRED
On Error GoTo Trace
Dim objs As Object
 
Dim obj As Object
Dim wmi As Object
Dim procid As String
Dim strComputer As String
strComputer = "."
Set wmi = GetObject("winmgmts:" & _
            "{impersonationLevel=impersonate}!\\" & _
            strComputer & "\root\cimv2")
'Set objs = wmi.InstancesOf("Win32_BaseBoard")
Set objs = wmi.ExecQuery("select * from Win32_BaseBoard", , 48)
'Set objs = WMI.InstancesOf(WindowState)
procid = ""
For Each obj In objs
  procid = procid & NullToStr(obj.SerialNumber)
'  If procid < objs.Count Then procid = procid & ","
Next
MBSerialNumber = NullToStr(procid)
'procid = LTrim$(procid)
'procid = RTrim$(procid)
'MsgBox "Proc_id :" + procid
Trace:
  If Err.Number <> 0 Then
    MsgBox "Kesalahan : " & Err.Number & " - " & Err.Description, vbCritical, App.Title
    Err.Clear
  End If
End Function

Public Function GetCpuID() As String
On Error GoTo Trace
  Dim wmi, cpu, cpuid, colProcesses
  Dim strComputer As String
  strComputer = "."
  Set wmi = GetObject("winmgmts:" & _
            "{impersonationLevel=impersonate}!\\" & _
            strComputer & "\root\cimv2")
  Set colProcesses = wmi.ExecQuery("select * from Win32_Processor ")
  cpuid = ""
  For Each cpu In colProcesses
    If Replace(UCase(NullToStr(cpu.ProcessorID)), " ", "") = UCase("") Then
      cpuid = cpuid + Mid(Replace(NullToStr(cpu.Name), " ", ""), Len(Replace(NullToStr(cpu.Name), " ", "")) / 2, 5) + NullToStr(cpu.Family) + NullToStr(cpu.Architecture)
    Else
      cpuid = cpuid + NullToStr(cpu.ProcessorID)
    End If
  Next
  GetCpuID = CStr(NullToStr(cpuid))
Trace:
  If Err.Number <> 0 Then
    MsgBox "Kesalahan : " & Err.Number & " - " & Err.Description, vbCritical, App.Title
    Err.Clear
  End If

End Function


