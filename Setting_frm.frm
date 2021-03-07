VERSION 5.00
Begin VB.Form Setting_frm 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Persentase Tax and Discount"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3375
   ClipControls    =   0   'False
   Icon            =   "Setting_frm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   3375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1560
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Batal"
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   5
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Simpan"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   300
      Index           =   1
      Left            =   2400
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   300
      Index           =   0
      Left            =   2400
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Global Discount (%)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tax and Service (%)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Index           =   2
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Setting_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    If Command1(0).Caption = "Simpan" Then
        simpan
    Else
        buka
        cmd_simpan
        Text1(0).SetFocus
    End If
Case 1
    If Command1(1).Caption = "Exit" Then
        Unload Me
    Else
        tutup
        cmd_awal
    End If
End Select
End Sub

Private Sub Form_Activate()
Data1.Refresh
With Data1.Recordset
If Data1.Recordset.BOF Then
    .AddNew
    !tax = 0
    !disc = 0
    .Update
Else
    isi
End If
End With
End Sub

Private Sub Form_Load()
Call db_set
kosong
tutup
cmd_awal
End Sub

Sub kosong()
Text1(0) = ""
Text1(1) = ""
End Sub

Sub tutup()
Text1(0).Enabled = False
Text1(1).Enabled = False
End Sub

Sub buka()
Text1(0).Enabled = True
Text1(1).Enabled = True
End Sub

Sub isi()
With Data1.Recordset
If Not .BOF And Command1(0).Caption <> "Simpan" Then
    Text1(0) = !tax
    Text1(1) = !disc
End If
End With
End Sub

Sub cmd_awal()
Command1(0).Caption = "Edit"
Command1(1).Caption = "Exit"
End Sub

Sub cmd_simpan()
Command1(0).Caption = "Simpan"
Command1(1).Caption = "Batal"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Visible = True
Form1.Show
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Sub simpan()
If Text1(0) = "" Or Text1(1) = "" Then
    MsgBox "Maaf data tidak boleh kosong...", vbInformation, "Validasi Input"
    If Text1(0) = "" Then
        Text1(0).SetFocus
    Else
        Text1(1).SetFocus
    End If
Else
    With Data1.Recordset
        .Edit
        !tax = Text1(0)
        !disc = Text1(1)
        .Update
    End With
    Data1.Refresh
    tutup
    cmd_awal
End If
End Sub
