VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   2385
   ClientLeft      =   2790
   ClientTop       =   3150
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1409.136
   ScaleMode       =   0  'User
   ScaleWidth      =   4154.835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2400
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1650
      TabIndex        =   1
      Top             =   840
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   480
      TabIndex        =   4
      Top             =   1725
      Width           =   1740
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2220
      TabIndex        =   5
      Top             =   1725
      Width           =   1740
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1650
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1230
      Width           =   2325
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SIMPLE RESTAURANT PROGRAM LOGIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   0
      Left            =   465
      TabIndex        =   0
      Top             =   855
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   1
      Left            =   465
      TabIndex        =   2
      Top             =   1245
      Width           =   1080
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   1920
      Picture         =   "frmLogin.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function CreateRoundRectRgn _
Lib "gdi32" (ByVal X1 As Long, _
ByVal Y1 As Long, _
ByVal X2 As Long, _
ByVal Y2 As Long, _
ByVal X3 As Long, _
ByVal Y3 As Long) As Long

Private Declare Function SetWindowRgn _
Lib "user32" (ByVal hwnd As Long, _
ByVal hRgn As Long, _
ByVal bRedraw As Boolean) As Long


Public Sub CreateRoundRectFromWindow(ByRef oWindow As Object)

Dim lRight As Long
Dim lBottom As Long
Dim hRgn As Long

With oWindow
lRight = .Width / Screen.TwipsPerPixelX
lBottom = .Height / Screen.TwipsPerPixelY
hRgn = CreateRoundRectRgn(0, 0, lRight, lBottom, 100, 100)
SetWindowRgn .hwnd, hRgn, True
End With

End Sub

Private Sub Command1_Click()
Unload Me
Unload Form4
MDIForm1.Show
End Sub

Private Sub Form_Load()
Call DB_Login
CreateRoundRectFromWindow Me

End Sub


Private Sub Form_Activate()
Data1.Refresh
txtUserName = ""
txtPassword = ""
txtUserName.SetFocus
End Sub


Private Sub cmdCancel_Click()
    LoginSucceeded = False
    x = MsgBox("Apakah anda yakin ingin keluar...???", vbYesNo, "Exit Program")
    If x = vbYes Then
        End
    End If
End Sub

Private Sub cmdOK_Click()
    Data1.RecordSource = "select * from user where user='" & txtUserName & "' and pass='" & txtPassword & "'"
    Data1.Refresh
    With Data1.Recordset
    If Not .BOF Then
        LoginSucceeded = True
        Me.Hide
        Form1.Label1.Caption = "User Name : " & !User
        Form1.Show
        frmSplash.Show
    Else
        MsgBox "Invalid username or Password, please try again!", , "Login"
        txtUserName.SetFocus
        'SendKeys "{Home}+{End}"
    End If
    End With
End Sub


