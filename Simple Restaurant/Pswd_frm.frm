VERSION 5.00
Begin VB.Form Pswd_frm 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Masukan Password"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3360
   ClipControls    =   0   'False
   Icon            =   "Pswd_frm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "BATAL"
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "Pswd_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    If Text1 = "a" Then
        PSWD = True
        Form1.Visible = False
        If prd = True Then
            Produk_frm.Show
        ElseIf lap = True Then
            lap_frm.Show
        ElseIf sett = True Then
            Setting_frm.Show
        ElseIf use = True Then
            User_frm.Show
        ElseIf dat = True Then
            Data_frm.Show
        End If
        Unload Me
    Else
        MsgBox "Maaf password yang anda masukan salah...", vbInformation, "Validasi Password"
        Text1.SetFocus
        PSWD = False
    End If
Case 1
    Unload Me
    Form1.Visible = True
End Select
End Sub

Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Form_Load()
Text1 = ""
Text1.MaxLength = 50
End Sub

Private Sub Form_Unload(Cancel As Integer)
prd = False
lap = False
sett = False
use = False
dat = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1_Click (0)
End If
End Sub
