VERSION 5.00
Begin VB.Form frmPayment2 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   1800
      TabIndex        =   3
      Top             =   2190
      Width           =   525
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   1800
      TabIndex        =   2
      Top             =   1740
      Width           =   525
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   1800
      TabIndex        =   1
      Top             =   1290
      Width           =   525
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   1800
      TabIndex        =   0
      Top             =   840
      Width           =   525
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lembar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2430
      TabIndex        =   13
      Top             =   2295
      Width           =   825
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lembar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2430
      TabIndex        =   12
      Top             =   1845
      Width           =   825
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lembar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2430
      TabIndex        =   11
      Top             =   1350
      Width           =   825
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UKURAN 15"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   450
      TabIndex        =   10
      Top             =   2295
      Width           =   1275
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UKURAN 21"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   450
      TabIndex        =   9
      Top             =   1800
      Width           =   1275
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UKURAN 28"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   450
      TabIndex        =   8
      Top             =   1350
      Width           =   1275
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ISILAH JUMLAH PEMAKAIAN LALU TEKAN ENTER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   30
      TabIndex        =   7
      Top             =   2925
      Width           =   5325
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lembar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2460
      TabIndex        =   6
      Top             =   915
      Width           =   825
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UKURAN 35"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   480
      TabIndex        =   5
      Top             =   870
      Width           =   1275
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   2730
      Left            =   135
      Top             =   90
      Width           =   5190
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   2835
      Left            =   90
      Top             =   60
      Width           =   5295
   End
   Begin VB.Label lbTanya 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PEMAKAIAN KANTONG PLASTIK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   270
      TabIndex        =   4
      Top             =   300
      Width           =   4665
   End
End
Attribute VB_Name = "frmPayment2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim jawab As Boolean
Dim tanya As String
Dim dbs As Database
Dim rs As Recordset
Dim KeyOk As Integer

Dim Ax As Double
Dim Bx As Double
Dim Cx As Double
Dim Dx As Double

Private Sub Form_Load()
isHasilKonversi = False
  jawab = False
  If isOnline = False Then
    Set dbs = OpenDatabase(DirDatabase & "\dbMaster.mdb")
  Else
    Set dbs = OpenDatabase(DirDbServer & "\dbMaster.mdb")
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
dbs.Close
Set dbs = Nothing
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim hasil As String
hasil = Trim(SendByCode(KeyCode))
Select Case hasil
Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
  Text1.Text = Text1.Text & hasil
  Text1.SelStart = Len(Text1.Text)
 Case "SPC"
  Text1.Text = Text1.Text & " "
  Text1.SelStart = Len(Text1.Text)
Case "BKS"
  If Len(Text1.Text) > 0 Then Text1.Text = Left(Text1.Text, Len(Text1.Text) - 1)
   Text1.SelStart = Len(Text1.Text)
Case "CLR"
        Text1.Text = ""
Case "ENT"
'        Set rs = dbs.OpenRecordset("Select Password from Memp Where isPengawas=true")
'        If rs.EOF And rs.BOF Then
'        Else
'          rs.MoveFirst
'          rs.FindFirst "Password='" & Replace(Text1.Text, "'", "''") & "'"
'          If rs.NoMatch Then
'            Text1.Text = ""
'            Exit Sub
'          Else
'            jawab = True
'            Unload Me
'          End If
'        End If
    Text2.SetFocus
Case "ESC"
    jawab = False
    Unload Me
End Select
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Sub Tampil(ByRef jawaban As Boolean, A As Double, B As Double, C As Double, D As Double)
  'KeyOk = Key
  Me.Show 1
  jawaban = jawab
  A = Ax
  B = Bx
  C = Cx
  D = Dx
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim hasil As String
hasil = Trim(SendByCode(KeyCode))
Select Case hasil
Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
  Text2.Text = Text2.Text & hasil
  Text2.SelStart = Len(Text2.Text)
 Case "SPC"
  Text2.Text = Text2.Text & " "
  Text2.SelStart = Len(Text2.Text)
Case "BKS"
  If Len(Text2.Text) > 0 Then Text2.Text = Left(Text2.Text, Len(Text2.Text) - 1)
   Text2.SelStart = Len(Text2.Text)
Case "CLR"
        Text2.Text = ""
Case "UP"
        Text1.SetFocus
Case "ENT"
'        Set rs = dbs.OpenRecordset("Select Password from Memp Where isPengawas=true")
'        If rs.EOF And rs.BOF Then
'        Else
'          rs.MoveFirst
'          rs.FindFirst "Password='" & Replace(Text2.Text, "'", "''") & "'"
'          If rs.NoMatch Then
'            Text2.Text = ""
'            Exit Sub
'          Else
'            jawab = True
'            Unload Me
'          End If
'        End If
    Text4.SetFocus
Case "ESC"
    jawab = False
    Unload Me
End Select
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
Dim hasil As String
hasil = Trim(SendByCode(KeyCode))
Select Case hasil
Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
  Text4.Text = Text4.Text & hasil
  Text4.SelStart = Len(Text4.Text)
 Case "SPC"
  Text4.Text = Text4.Text & " "
  Text4.SelStart = Len(Text4.Text)
Case "BKS"
  If Len(Text4.Text) > 0 Then Text4.Text = Left(Text4.Text, Len(Text4.Text) - 1)
   Text4.SelStart = Len(Text4.Text)
Case "CLR"
        Text4.Text = ""
Case "UP"
        Text2.SetFocus
Case "ENT"
'        Set rs = dbs.OpenRecordset("Select Password from Memp Where isPengawas=true")
'        If rs.EOF And rs.BOF Then
'        Else
'          rs.MoveFirst
'          rs.FindFirst "Password='" & Replace(Text4.Text, "'", "''") & "'"
'          If rs.NoMatch Then
'            Text4.Text = ""
'            Exit Sub
'          Else
'            jawab = True
'            Unload Me
'          End If
'        End If
    Text5.SetFocus
Case "ESC"
    jawab = False
    Unload Me
End Select
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
Dim hasil As String
hasil = Trim(SendByCode(KeyCode))
Select Case hasil
Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
  Text5.Text = Text5.Text & hasil
  Text5.SelStart = Len(Text5.Text)
 Case "SPC"
  Text5.Text = Text5.Text & " "
  Text5.SelStart = Len(Text5.Text)
Case "BKS"
  If Len(Text5.Text) > 0 Then Text5.Text = Left(Text5.Text, Len(Text5.Text) - 1)
   Text5.SelStart = Len(Text5.Text)
Case "CLR"
        Text5.Text = ""
Case "UP"
        Text4.SetFocus
Case "ENT"
  Ax = NullToNol(Text1.Text)
  Bx = NullToNol(Text2.Text)
  Cx = NullToNol(Text4.Text)
  Dx = NullToNol(Text5.Text)
  
        jawab = True
        Unload Me
Case "ESC"
    jawab = False
    Unload Me
End Select
End Sub
