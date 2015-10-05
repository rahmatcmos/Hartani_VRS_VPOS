Attribute VB_Name = "Module6"
Public Function MBSerialNumber() As String
 
'RETRIEVES SERIAL NUMBER OF MOTHERBOARD
'IF THERE IS MORE THAN ONE MOTHERBOARD, THE SERIAL
'NUMBERS WILL BE DELIMITED BY COMMAS

'YOU MUST HAVE WMI INSTALLED AND A REFERENCE TO
'Microsoft WMI Scripting Library IS REQUIRED

Dim objs As Object
 
Dim obj As Object
Dim wmi As Object
 

 
Set wmi = GetObject("WinMgmts:")
Set objs = wmi.InstancesOf("Win32_BaseBoard")
'Set objs = WMI.InstancesOf(WindowState)
For Each obj In objs
procid = procid & obj.SerialNumber
If procid < objs.Count Then procid = procid & ","
Next
MBSerialNumber = procid
'procid = LTrim$(procid)
'procid = RTrim$(procid)
'MsgBox "Proc_id :" + procid
End Function

Function GetCpuID()
  Dim wmi, cpu, cpuid
  Set wmi = GetObject("winmgmts:")
  For Each cpu In wmi.InstancesOf("Win32_Processor")
   cpuid = cpuid + cpu.ProcessorID
  Next
  GetCpuID = cpuid
End Function
