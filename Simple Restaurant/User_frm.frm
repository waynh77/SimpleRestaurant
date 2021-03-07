VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form User_frm 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Account"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10935
   ClipControls    =   0   'False
   Icon            =   "User_frm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   10935
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   6000
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   5400
      Top             =   120
   End
   Begin VB.TextBox Text1 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   5880
      PasswordChar    =   "*"
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1800
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   5880
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1320
      Width           =   3975
   End
   Begin VB.Data Data1 
      BackColor       =   &H00000000&
      Caption         =   "DATA PENGGUNA PROGRAM"
      Connect         =   "Access"
      DatabaseName    =   " "
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   5160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   " "
      Top             =   2400
      Width           =   3735
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "User_frm.frx":1982
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "User_frm.frx":3314
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "User_frm.frx":4CA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "User_frm.frx":6638
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "User_frm.frx":7FCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "User_frm.frx":995C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "User_frm.frx":A636
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "User_frm.frx":B310
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   3480
      Left            =   10125
      TabIndex        =   1
      Top             =   0
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   6138
      ButtonWidth     =   1191
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Tambah"
            Object.ToolTipText     =   "Tambah Data Baru"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            Object.ToolTipText     =   "Edit Data"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Hapus"
            Object.ToolTipText     =   "Hapus Data"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Cari"
            Object.ToolTipText     =   "Cari Data"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Preview"
            Object.ToolTipText     =   "Print Preview"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Cetak"
            Object.ToolTipText     =   "Cetak Data Ke Printer"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Keluar"
            Object.ToolTipText     =   "Tutup"
            ImageIndex      =   7
         EndProperty
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "User_frm.frx":CCA2
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBGrid1 
      Bindings        =   "User_frm.frx":CFBC
      Height          =   3135
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3735
      _Version        =   196614
      BevelColorFace  =   192
      AllowUpdate     =   0   'False
      RowHeight       =   423
      Columns(0).Width=   5953
      Columns(0).Caption=   "USER NAME"
      Columns(0).Name =   "JENIS PRODUK"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "user"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      _ExtentX        =   6588
      _ExtentY        =   5530
      _StockProps     =   79
      ForeColor       =   16777215
      BackColor       =   8421504
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   0
      Left            =   4080
      Picture         =   "User_frm.frx":CFD0
      Top             =   480
      Width           =   720
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PASSWORD"
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
      Left            =   4080
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "USER NAME"
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
      Index           =   1
      Left            =   4080
      TabIndex        =   0
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "USER ACCOUNT"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   4080
      TabIndex        =   6
      Top             =   360
      Width           =   5775
   End
End
Attribute VB_Name = "User_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tambah As Boolean
Dim cek As Boolean
Dim cek2 As Boolean

Private Sub Data1_Reposition()
isi
End Sub

Private Sub Form_Activate()
Data1.Refresh
isi
End Sub

Private Sub Form_Load()
Call DB_user
tutup
kosong
Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Visible = True
Form1.Show
End Sub

Private Sub SSDBGrid1_Click()
isi
End Sub

Private Sub Timer1_Timer()
If Image1(0).Left < 9120 Then
    Image1(0).Left = Image1(0).Left + 100
Else
    Timer1.Enabled = False
    Timer2.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
If Image1(0).Left > 4080 Then
    Image1(0).Left = Image1(0).Left - 100
Else
    Timer2.Enabled = False
    Timer1.Enabled = True
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    If Toolbar1.Buttons(1).Caption = "Tambah" Then
        buka
        kosong
        tambah = True
        cmd_simpan
        Text1(0).SetFocus
    Else
        simpan
    End If
Case 2
    If Toolbar1.Buttons(2).Caption = "Edit" Then
        If Text1(0) <> "" Then
            buka
            tambah = False
            cmd_simpan
        Else
            MsgBox "Data Kosong", vbInformation, "Validasi Data"
        End If
    Else
        cmd_awal
        tutup
        isi
    End If
Case 3
    Hapus
Case 4
Case 5
Case 6
Case 7
    Unload Me
End Select
End Sub

Sub kosong()
Text1(0) = ""
Text1(1) = ""
End Sub

Sub isi()
With Data1.Recordset
    If Not .BOF And Data1.Enabled = True Then
        Text1(0) = !User
        Text1(1) = !pass
    End If
End With
End Sub

Sub tutup()
    Text1(0).Enabled = False
    Text1(1).Enabled = False
End Sub

Sub buka()
    Text1(0).Enabled = True
    Text1(1).Enabled = True
End Sub

Sub simpan()
cek_input
If cek2 = False Then
    MsgBox "Input tidak valid, mohon diperiksa kembali", vbInformation, "Validasi Input"
Else
    With Data1.Recordset
        If tambah = True Then
            cek_tambah
            If cek = False Then
            Data1.Refresh
                .AddNew
                !User = Text1(0)
                !pass = Text1(1)
                .Update
                tutup
                cmd_awal
            Else
                MsgBox "Data sudah ada,silahkan isi yang lain...", vbInformation, "Validasi Data"
            End If
        Else
            .Edit
            !User = Text1(0)
            !pass = Text1(1)
            .Update
            tutup
            cmd_awal
        End If
    End With
    Data1.Refresh
End If
End Sub

Sub cek_input()
cek2 = False
If Text1(0) = "" Or Text1(1) = "" Then
    cek2 = False
Else
    cek2 = True
End If
End Sub

Sub cek_tambah()
cek = False
With Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        If Text1(0) = !User Then
            cek = True
            .MoveLast
        End If
        .MoveNext
    Loop
End If
End With
End Sub

Sub Hapus()
With Data1.Recordset
    If Not .BOF Then
        If .RecordCount > 1 Then
            x = MsgBox("Apakah anda yakin menghapus data?", vbYesNo, "Hapus Data")
            If x = vbYes Then
                .Delete
                kosong
                Data1.Refresh
            End If
        Else
            MsgBox "Maaf data user tidak boleh kosong (minimal 1 user)", vbCritical, "Validasi Data"
        End If
    Else
        MsgBox "Data masih kosong/belum dipilih...", vbInformation, "Validasi Data"
    End If
End With
End Sub

Sub cmd_awal()
With Toolbar1
    .Buttons(1).Image = 1
    .Buttons(2).Image = 2
    .Buttons(1).Caption = "Tambah"
    .Buttons(2).Caption = "Edit"
    .Buttons(1).ToolTipText = "Tambah Data"
    .Buttons(2).ToolTipText = "Edit Data"
    .Buttons(3).Visible = True
    .Buttons(4).Visible = False
    .Buttons(5).Visible = False
    .Buttons(6).Visible = False
    .Buttons(7).Visible = True
End With
Data1.Enabled = True
End Sub

Sub cmd_simpan()
With Toolbar1
    .Buttons(1).Image = 8
    .Buttons(2).Image = 3
    .Buttons(1).Caption = "Simpan"
    .Buttons(2).Caption = "Batal"
    .Buttons(1).ToolTipText = "Simpan Data"
    .Buttons(2).ToolTipText = "Batal Data"
    .Buttons(3).Visible = False
    .Buttons(4).Visible = False
    .Buttons(5).Visible = False
    .Buttons(6).Visible = False
    .Buttons(7).Visible = False
End With
Data1.Enabled = False
End Sub
