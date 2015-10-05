Attribute VB_Name = "SetReg"
'---------------------------------------------------------------------------------------
' Module    : mdlRegistry
' DateTime  : Pebruari - 2005
' Author    : Haniif Badrii
' Purpose   : Module to get Registry Setting
'             Registry will be save in \\LOCAL_MACHINE\SOFTWARE\AHS\ERP\
'---------------------------------------------------------------------------------------
' Copy via Amik and diikhlaskan to edit by Syalim sesuai kebutuhan
' Thanks's All and happy coding

Option Explicit
'Private Const REG_SETTING As String = "SOFTWARE\Testing\ProjectManagement\"
Private Const REG_SETTING As String = "SOFTWARE\VPOINT\VPOS\"
Private Const REG_RUN As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run\"
Private oReg As New clsRegistry

Function getRegistry(ByVal strValueName As String, _
                     ByVal strSection As String) As String
    getRegistry = oReg.GetKeyValue(rrkHKeyLocalMachine, _
                                   REG_SETTING & strSection, _
                                   strValueName)
End Function

Sub SetRegistry(ByVal strValueName As String, _
                ByVal strValue As String, _
                Optional ByVal strSection As String)
    oReg.SetKeyValue rrkHKeyLocalMachine, _
                     REG_SETTING & strSection, _
                     strValueName, strValue, rrkRegSZ
End Sub

Sub SetRunApp(ByVal strValueName As String, _
                ByVal strValue As String, _
                Optional ByVal strSection As String)
    oReg.SetKeyValue rrkHKeyLocalMachine, _
                     REG_RUN & strSection, _
                     strValueName, strValue, rrkRegSZ
End Sub

Sub SetCredits()
    SetRegistry "Company", "AHS", "Credits"
    SetRegistry "Address", "Jl.Gebang Lor 69 A Surabaya", "Credits"
    SetRegistry "Project Manager", "Isyalim", "Credits"
End Sub

Function getCredits() As String()
'FIXIT: Non Zero lowerbound arrays are not supported in Visual Basic .NET                  FixIT90210ae-R9815-H1984
    Dim temp(1 To 2) As String
    temp(1) = getRegistry("Project Manager", "Credits")
    temp(2) = getRegistry("Programmer", "Credits")
    getCredits = temp
End Function


Sub SetCProfile(ByVal Perusahaan As String, _
                ByVal Alamat1 As String, _
                ByVal Alamat2 As String, _
                ByVal Kota As String, _
                ByVal Propinsi As String, _
                ByVal Telp As String, _
                ByVal Fax As String, _
                ByVal Contact As String)
    SetRegistry "Perusahaan", Perusahaan, "Profile"
    SetRegistry "Alamat1", Alamat1, "Profile"
    SetRegistry "Alamat2", Alamat2, "Profile"
    SetRegistry "Kota", Kota, "Profile"
    SetRegistry "Propinsi", Propinsi, "Profile"
    SetRegistry "Telp", Telp, "Profile"
    SetRegistry "Fax", Fax, "Profile"
    SetRegistry "Contact", Contact, "Profile"
End Sub

Function getCProfile() As String()
'FIXIT: Non Zero lowerbound arrays are not supported in Visual Basic .NET                  FixIT90210ae-R9815-H1984
    Dim temp(1 To 8) As String
    temp(1) = getRegistry("Perusahaan", "Profile")
    temp(2) = getRegistry("Alamat1", "Profile")
    temp(3) = getRegistry("Alamat2", "Profile")
    temp(4) = getRegistry("Kota", "Profile")
    temp(5) = getRegistry("Propinsi", "Profile")
    temp(6) = getRegistry("Telp", "Profile")
    temp(7) = getRegistry("Fax", "Profile")
    temp(8) = getRegistry("Contact", "Profile")
    getCProfile = temp
End Function
Sub SetSMS(ByVal nPort As String, ByVal nBits As String, ByVal nData As String)
    SetRegistry "nPort", nPort, "SetSMS"
    SetRegistry "nBits", nBits, "SetSMS"
    SetRegistry "nData", nData, "SetSMS"
End Sub
Sub SetSetup(ByVal IsStart As String)
    SetRegistry "IsSetup", "1", "Setup"
    SetRegistry "IsStart", IsStart, "Setup"
    SetRunApp "IsStart", IsStart, "Setup"
End Sub
Function GetSMS() As String()
Dim temp(1 To 3) As String
temp(1) = getRegistry("nPort", "SetSMS")
temp(2) = getRegistry("nBits", "SetSMS")
temp(3) = getRegistry("nData", "SetSMS")
GetSMS = temp
End Function
'Public Sub SetDatabaseToRegsitry()
'    ' Database Type :
'    ' 1 : MsAccess
'    ' 2 : SQL server
'    ' 3 : Oracle
'
'    ' Server
'    ' Access        : file location
'    ' SQL Server    : server name
'    ' Oracle        : TNS NAme
'    SetRegistry "Database Name", setDatabaseName, "Setting Database"
'    SetRegistry "Server Name", setServerName, "Setting Database"
'    SetRegistry "Windows Authentication", IIf(setWindowsAuthentication, "1", "0"), "Setting Database"
'    SetRegistry "Db Pwd", setDatabasePassword, "Setting Database"
'    SetRegistry "UID", setDatabaseUser, "Setting Database"
'    SetRegistry "ODBC Name", setODBCName, "Setting Database"
'
'End Sub

'Public Sub getDatabasefromRegistry()
''FIXIT: Non Zero lowerbound arrays are not supported in Visual Basic .NET                  FixIT90210ae-R9815-H1984
'  Dim dbProvider As String
'  dbProvider = "MSDASQL.1"
'    setDatabaseName = getRegistry("Database Name", "Setting Database")
'    setServerName = getRegistry("Server Name", "Setting Database")
'    setWindowsAuthentication = NulltoDbl(getRegistry("Windows Authentication", "Setting Database"))
'    setDatabasePassword = getRegistry("Db Pwd", "Setting Database")
'    setDatabaseUser = getRegistry("UID", "Setting Database")
'    setODBCName = getRegistry("ODBC Name", "Setting Database")
'    If setWindowsAuthentication Then
'      KoneksiStr = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & setDatabaseName & ";Data Source=" & setServerName
'    Else
'      KoneksiStr = "Provider=SQLOLEDB.1;Password=" & setDatabasePassword & ";Persist Security Info=True;User ID=" & setDatabaseUser & ";Initial Catalog=" & setDatabaseName & ";Data Source=" & setServerName
'    End If
''    getDatabase = temp
'End Sub


