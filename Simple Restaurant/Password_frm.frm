VERSION 5.00
Begin VB.Form Password_frm 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ganti Password"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   Icon            =   "Password_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      BackColor       =   &H00000000&
      Caption         =   "data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BATAL"
      Height          =   615
      Index           =   1
      Left            =   2760
      TabIndex        =   8
      Top             =   2160
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PROSES"
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   2760
      PasswordChar    =   "*"
      TabIndex        =   6
      Text            =   " "
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2760
      PasswordChar    =   "*"
      TabIndex        =   5
      Text            =   " "
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   2760
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   " "
      Top             =   600
      Width           =   2535
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   5280
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KONFIRMASI PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PASSWORD BARU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5280
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PASSWORD LAMA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "USER NAME : ADMIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   13
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "Password_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Dim cek_input As Boolean
Dim cek_data As Boolean
Select Case Index
Case 0
    If Text1(0) = "" Or Text1(1) = "" Or Text1(2) = "" Then
        MsgBox "Input belum lengkap...", vbInformation, "Validasi Input"
        If Text1(0) = "" Then
            Text1(0).SetFocus
        ElseIf Text1(1) = "" Then
            Text1(1).SetFocus
        Else
            Text1(2).SetFocus
        End If
    Else
        With Data1.Recordset
            If Text1(0) <> !pass Then
                MsgBox "Maaf password lama tidak valid...", vbCritical, "Validasi Password"
                Text1(0).SetFocus
            Else
                If Text1(1) <> Text1(2) Then
                    MsgBox "Maaf password baru tidak valid...", vbCritical, "Validasi Password"
                    Text1(1).SetFocus
                Else
                    .Edit
                    !pass = Text1(1)
                    .Update
                    MsgBox "Password berhasil diubah...", vbInformation, "Validasi Password"
                    Unload Me
                End If
            End If
        End With
    End If
Case 1
    Unload Me
End Select
End Sub

Private Sub Form_Activate()
Data1.DatabaseName = Form1.Data4.DatabaseName
Data1.RecordSource = Form1.Data4.RecordSource
Data1.Refresh
Text1(0).SetFocus
End Sub

Private Sub Form_Load()
kosong
End Sub

Sub kosong()
Text1(0) = ""
Text1(1) = ""
Text1(2) = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Enabled = True
    Form1.Visible = True
    Form1.Show
End Sub
