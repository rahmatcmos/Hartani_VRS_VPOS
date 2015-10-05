VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsRegistry
' DateTime  : 01 - Feb - 2004
' Author    : Haniif Badrii
' Purpose   : This module Implements routines for manipulating the registry.
'---------------------------------------------------------------------------------------

Option Explicit

Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal lngHKey As Long) As Long

Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal lngHKey As Long, ByVal lpSubKey As String, ByVal reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal lngHKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long

Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal lngHKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long

Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal lngHKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long

Private Declare Function RegQueryValueExBinary Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal lngHKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
  
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal lngHKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long

Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal lngHKey As Long, ByVal lpValueName As String, ByVal reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long

Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal lngHKey As Long, ByVal lpValueName As String, ByVal reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long

Private Declare Function RegSetValueExBinary Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal lngHKey As Long, ByVal lpValueName As String, ByVal reserved As Long, ByVal dwType As Long, ByVal lpValue As Long, ByVal cbData As Long) As Long
  
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal lngHKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long

Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal lngHKey As Long, ByVal lpClass As String, ByVal lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, ByVal lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, ByVal lpcbMaxValueLen As Long, ByVal lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long

Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal lngHKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, ByVal lpType As Long, ByVal lpData As Byte, ByVal lpcbData As Long) As Long

Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal lngHKey As Long, ByVal lpSubKey As String) As Long

Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal lngHKey As Long, ByVal lpValueName As String) As Long

Public Enum EnumRegistryRootKeys
  rrkHKeyClassesRoot = &H80000000
  rrkHKeyCurrentUser = &H80000001
  rrkHKeyLocalMachine = &H80000002
  rrkHKeyUsers = &H80000003
End Enum

Public Enum EnumRegistryValueType
  rrkRegSZ = 1
  rrkregBinary = 3
  rrkRegDWord = 4
End Enum

Private Const mcregOptionNonVolatile = 0

Private Const mcregErrorNone = 0
Private Const mcregErrorBadDB = 1
Private Const mcregErrorBadKey = 2
Private Const mcregErrorCantOpen = 3
Private Const mcregErrorCantRead = 4
Private Const mcregErrorCantWrite = 5
Private Const mcregErrorOutOfMemory = 6
Private Const mcregErrorInvalidParameter = 7
Private Const mcregErrorAccessDenied = 8
Private Const mcregErrorInvalidParameterS = 87
Private Const mcregErrorNoMoreItems = 259

Private Const mcregSynchronize = &H100000


Private Const mcregKeyQueryValue = &H1
Private Const mcregKeySetValue = &H2
Private Const mcregKeyCreateSubKey = &H4
Private Const mcregKeyEnumerateSubKeys = &H8
Private Const mcregKeyCreateLink = &H20
Private Const mcregKeyNotify = &H10
Private Const mcregReadControl = &H20000
Private Const mcregStandardRightsAll = &H1F0000
Private Const mcregStandardRightsRead = (mcregReadControl)
Private Const mcregStandardRightsWrite = (mcregReadControl)

Private Const mcregKeyAllAccess = ((mcregStandardRightsAll Or mcregKeyQueryValue Or mcregKeySetValue Or mcregKeyCreateSubKey Or mcregKeyEnumerateSubKeys Or mcregKeyNotify Or mcregKeyCreateLink) And (Not mcregSynchronize))
Private Const mcregKeyRead = ((mcregStandardRightsRead Or mcregKeyQueryValue Or mcregKeyEnumerateSubKeys Or mcregKeyNotify) And (Not mcregSynchronize))
Private Const mcregKeyWrite = ((mcregStandardRightsWrite Or mcregKeySetValue Or mcregKeyCreateSubKey) And (Not mcregSynchronize))

Public Sub CreateNewKey(eRootKey As EnumRegistryRootKeys, strKeyName As String)
    ' Comments  : Creates a new key in the system registry
    ' Parameters: eRootKey - The root key
    '             strKeyName - The name of the key to create
    ' Returns   : Nothing
    
    Dim lngRetVal As Long
    Dim lngHKey As Long
    
    On Error GoTo PROC_ERR
    
    ' Create the key
    lngRetVal = RegCreateKeyEx(eRootKey, strKeyName, 0&, vbNullString, mcregOptionNonVolatile, mcregKeyWrite, 0&, lngHKey, 0&)
    
    ' if the key was created, then close it
    If lngRetVal = mcregErrorNone Then
        RegCloseKey (lngHKey)
    End If
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "RegistryCreateNewKey"
    Resume PROC_EXIT
    
End Sub

Public Sub DeleteKey(eRootKey As EnumRegistryRootKeys, strKeyName As String)
    ' Comments  : Deletes a key from the system registry
    ' Parameters: eRootKey - The root key
    '             strKeyName - The name of the key to delete
    ' Returns   : Nothing
    '
    Dim lngRetVal As Long
  
    On Error GoTo PROC_ERR
      
    ' Delete the key
    lngRetVal = RegDeleteKey(eRootKey, strKeyName)
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "RegistryDeleteKey"
    Resume PROC_EXIT
    
End Sub

Public Sub DeleteValue(eRootKey As EnumRegistryRootKeys, strKeyName As String, strValueName As String)
    ' Comments  : Deletes a value from the system registry
    ' Parameters: eRootKey - The root key
    '             strKeyName - The name of the key to delete
    '             strValueName - The name of the value to delete
    ' Returns   : Nothing
    
    Dim lngRetVal As Long
    Dim lngHKey As Long

    On Error GoTo PROC_ERR

    ' Open the key
    lngRetVal = RegOpenKeyEx(eRootKey, strKeyName, 0, mcregKeyWrite, lngHKey)

    ' If the key was opened successfully, then delete it
    If lngRetVal = mcregErrorNone Then
        lngRetVal = RegDeleteValue(lngHKey, strValueName)
    End If

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "RegistryDeleteValue"
    Resume PROC_EXIT

End Sub

Public Sub EnumerateSubKeys(eRootKey As EnumRegistryRootKeys, strKeyName As String, astrKeys() As String, lngKeyCount As Long)
    ' Comments  : Enumerates the sub keys of the specified key
    ' Parameters: eRootKey - The root key
    '             strKeyName - The name of the key to enumerate
    '             astrKeys - An array of strings to fill with sub key names
    '             lngKeyCount - The number of sub keys returned in the parameter
    '             astrKeys
    ' Returns   : Nothing
    
    Dim lngRetVal As Long
    Dim lngHKey As Long
    Dim lngKeyIndex As Long
    Dim strSubKeyName As String
    Dim lngSubkeyCount As Long
    Dim lngMaxKeyLen As Long
    Dim typFT As FILETIME
  
    On Error GoTo PROC_ERR
  
    ' Open the key
    lngRetVal = RegOpenKeyEx(eRootKey, strKeyName, 0, mcregKeyRead, lngHKey)
  
    If lngRetVal = mcregErrorNone Then
        'find the number of subkeys, and redim the return string array
        lngRetVal = RegQueryInfoKey(lngHKey, vbNullString, 0, 0, lngSubkeyCount, lngMaxKeyLen, 0, 0, 0, 0, 0, typFT)
        If mcregErrorNone = lngRetVal Then
            If lngSubkeyCount > 0 Then
                ReDim astrKeys(lngSubkeyCount - 1) As String
        
                'set up the while loop
                lngKeyIndex = 0
                ' Pad the string to the maximum length of a sub key, plus 1 for null
                ' termination
                lngMaxKeyLen = lngMaxKeyLen + 1
                strSubKeyName = Space$(lngMaxKeyLen)
        
                Do While RegEnumKey(lngHKey, lngKeyIndex, strSubKeyName, lngMaxKeyLen + 1) = 0
        
                    ' Set the string array to the key name, removing null termination
                    If InStr(1, strSubKeyName, vbNullChar) > 0 Then
                        astrKeys(lngKeyIndex) = Left$(strSubKeyName, InStr(1, strSubKeyName, vbNullChar) - 1)
                    End If
                    ' Increment the key index for the return string array
                    lngKeyIndex = lngKeyIndex + 1
        
                Loop
            End If
            ' return the new dimension of the return string array
            lngKeyCount = lngSubkeyCount
        End If
    
        ' Close the key
        RegCloseKey (lngHKey)
    End If
  
PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
       "RegistryEnumerateSubKeys"
    Resume PROC_EXIT

End Sub

Public Sub EnumerateValues( _
       eRootKey As EnumRegistryRootKeys, _
       strKeyName As String, _
       astrValues() As String, _
       lngValueCount As Long)
    ' Comments  : Enumerates the values of the specified key
    ' Parameters: eRootKey - The root key
    '             strKeyName - The name of the key to enumerate
    '             astrValues - An array of strings to fill with value names
    '             lngValueCount - The number of values returned in the parameter
    '             astrValues
    ' Returns   : Nothing
    
    Dim lngRetVal As Long
    Dim lngHKey As Long
    Dim lngKeyIndex As Long
    Dim strValueName As String
    Dim lngTempValueCount As Long
    Dim lngMaxValueLen As Long
    Dim typFT As FILETIME
  
    On Error GoTo PROC_ERR
  
    ' Open the key
    lngRetVal = RegOpenKeyEx(eRootKey, strKeyName, 0, mcregKeyRead, _
       lngHKey)
  
    If lngRetVal = mcregErrorNone Then
        'find the number of subkeys, and redim the return string array
        lngRetVal = RegQueryInfoKey(lngHKey, vbNullString, 0, 0, 0, _
           0, 0, lngTempValueCount, lngMaxValueLen, 0, 0, typFT)
        If mcregErrorNone = lngRetVal Then
            If lngTempValueCount > 0 Then
                ReDim astrValues(lngTempValueCount - 1) As String
        
                'set up the while loop
                lngKeyIndex = 0
                ' Pad the string to the maximum length of a sub key, plus 1 for null
                ' termination
                lngMaxValueLen = lngMaxValueLen + 1
                strValueName = Space$(lngMaxValueLen)
        
                Do While RegEnumValue(lngHKey, lngKeyIndex, strValueName, _
                   lngMaxValueLen + 1, 0, 0, 0, 0) = 0
        
                    ' Set the string array to the key name, removing null termination
                    If InStr(1, strValueName, vbNullChar) > 0 Then
                        astrValues(lngKeyIndex) = Left$(strValueName, InStr(1, strValueName, _
                           vbNullChar) - 1)
                    End If
                    ' Increment the key index for the return string array
                    lngKeyIndex = lngKeyIndex + 1
        
                Loop
            End If
            ' return the new dimension of the return string array
            lngValueCount = lngTempValueCount
        End If
    
        ' Close the key
        RegCloseKey (lngHKey)
    End If
  
PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
       "RegistryEnumerateValues"
    Resume PROC_EXIT

End Sub

'FIXIT: Declare 'GetKeyValue' with an early-bound data type                                FixIT90210ae-R1672-R1B8ZE
Public Function GetKeyValue( _
       eRootKey As EnumRegistryRootKeys, _
       strKeyName As String, _
       strValueName As String) _
       As Variant
    ' Comments  : Returns a value from the system registry
    ' Parameters: eRootKey - The root key
    '             strKeyName - The name of the key
    '             strValueName - The name of the value
    ' Returns   : The data in the registry value
    
    Dim lngRetVal As Long
    Dim lngHKey As Long
'FIXIT: Declare 'varValue' with an early-bound data type                                   FixIT90210ae-R1672-R1B8ZE
    Dim varValue As Variant
    Dim strValueData As String
    Dim abytValueData() As Byte
    Dim lngValueData As Long
    Dim lngValueType As Long
    Dim lngDataSize As Long
  
    On Error GoTo PROC_ERR
  
    varValue = Empty
  
    lngRetVal = RegOpenKeyEx(eRootKey, strKeyName, 0&, mcregKeyRead, _
       lngHKey)
  
    If mcregErrorNone = lngRetVal Then
    
        lngRetVal = RegQueryValueExNULL(lngHKey, strValueName, 0&, lngValueType, _
           0&, lngDataSize)
    
        If lngRetVal = mcregErrorNone Then
      
            Select Case lngValueType
      
                    ' String type

                Case rrkRegSZ:
                    If lngDataSize > 0 Then
'FIXIT: Replace 'String' function with 'String$' function                                  FixIT90210ae-R9757-R1B8ZE
                        strValueData = String(lngDataSize, 0)
                        lngRetVal = RegQueryValueExString(lngHKey, strValueName, 0&, _
                           lngValueType, strValueData, lngDataSize)
                        If InStr(strValueData, vbNullChar) > 0 Then
                            strValueData = Mid$(strValueData, 1, InStr(strValueData, _
                               vbNullChar) - 1)
                        End If
                    End If
                    If mcregErrorNone = lngRetVal Then
                        varValue = Left$(strValueData, lngDataSize)
                    Else
                        varValue = Empty
                    End If
        
                    ' Long type
                Case rrkRegDWord:
                    lngRetVal = RegQueryValueExLong(lngHKey, strValueName, 0&, _
                       lngValueType, lngValueData, lngDataSize)
                    If mcregErrorNone = lngRetVal Then
                        varValue = lngValueData
                    End If
                
                    ' Binary type
                Case rrkregBinary
                    If lngDataSize > 0 Then
                        ReDim abytValueData(lngDataSize - 1) As Byte
'FIXIT: Keyword 'VarPtr' not supported in Visual Basic .NET                                FixIT90210ae-R6614-H1984
                        lngRetVal = RegQueryValueExBinary(lngHKey, strValueName, 0&, _
                           lngValueType, VarPtr(abytValueData(0)), lngDataSize)
                    End If
                    If mcregErrorNone = lngRetVal Then
                        varValue = abytValueData
                    Else
                        varValue = Empty
                    End If
                
                Case Else
                    'No other data types supported
                    lngRetVal = -1
        
            End Select
      
        End If
    
        RegCloseKey (lngHKey)
    End If
  
    'Return varValue
    GetKeyValue = varValue
  
PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
       "RegistryGetKeyValue"
    Resume PROC_EXIT
  
End Function

'FIXIT: Declare 'varData' with an early-bound data type                                    FixIT90210ae-R1672-R1B8ZE
Public Sub SetKeyValue( _
       eRootKey As EnumRegistryRootKeys, _
       strKeyName As String, _
       strValueName As String, _
       varData As Variant, _
       eDataType As EnumRegistryValueType)
    ' Comments  : This procedure sets a key value
    ' Parameters: eRootKey - The root key
    '             strKeyName - The name of the key
    '             strValueName - The name of the value
    '             varData - The data to store in the value
    '             eDataType - The type of data to store in the value
    ' Returns   : Nothing
    
    Dim lngRetVal As Long
    Dim lngHKey As Long
    Dim strData As String
    Dim lngData As Long
    Dim abytData() As Byte
    
    On Error GoTo PROC_ERR
  
    ' Open the specified key, if it does not exist then create it
    lngRetVal = RegCreateKeyEx(eRootKey, strKeyName, 0&, vbNullString, _
       mcregOptionNonVolatile, mcregKeyRead Or mcregKeyWrite, 0&, lngHKey, 0&)
  
    ' Determine the data type of the key
    Select Case eDataType
  
        Case rrkRegSZ
            strData = varData & vbNullChar
            lngRetVal = RegSetValueExString(lngHKey, strValueName, 0&, eDataType, _
               strData, Len(strData))
    
        Case rrkRegDWord
            lngData = varData
            lngRetVal = RegSetValueExLong(lngHKey, strValueName, 0&, eDataType, _
               lngData, Len(lngData))
  
            ' Binary type
        Case rrkregBinary
            abytData = varData
'FIXIT: Keyword 'VarPtr' not supported in Visual Basic .NET                                FixIT90210ae-R6614-H1984
            lngRetVal = RegSetValueExBinary(lngHKey, strValueName, 0&, eDataType, _
               VarPtr(abytData(0)), UBound(abytData) + 1)
  
    End Select
  
    RegCloseKey (lngHKey)
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
       "RegistrySetKeyValue"
    Resume PROC_EXIT
    
End Sub




