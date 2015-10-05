Attribute VB_Name = "mdlPublic"
Public app_ini As String

Sub Main()
  app_ini = App.Path & "\Setting.ini"
  strLogin = BelumSamaSekali
  NamaPerusahaan = getstringinifiles("APPCONFIG", "NamaPerusahaan", "RETORAN CEPAT SAJI", App.Path & "\Setting.ini")
  AlamatPerusahaan = getstringinifiles("APPCONFIG", "AlamatPerusahaan", "Jl. ABC ", App.Path & "\Setting.ini")
    
  If cOra.teskoneksi Then
    frmService.Show
    frmService.SetFocus
  Else
    If MsgBox("Melalukan verify setting database", vbInformation + vbYesNo) = vbYes Then
      If VerifyDatabase Then
        frmService.Show
        frmService.SetFocus
      Else
        End
      End If
    Else
      End
    End If
  End If
  cOra.CloseConnection
End Sub

Public Function VerifyDatabase() As Boolean
  setstringinifiles "DBCONFIG", "SERVER", "SERVER", app_ini
  setstringinifiles "DBCONFIG", "DBNAME", "DBCITYTOYS", app_ini
  setstringinifiles "DBCONFIG", "USER", Encrypt("sa"), app_ini
  setstringinifiles "DBCONFIG", "PWD", Encrypt("sgi"), app_ini
  setstringinifiles "dbtype", "TYPE", 4, app_ini
  VerifyDatabase = cOra.teskoneksi
End Function
