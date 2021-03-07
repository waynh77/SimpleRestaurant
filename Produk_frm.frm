VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form Produk_frm 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Produk"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8880
   Icon            =   "Produk_frm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   960
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New"
      Height          =   315
      Left            =   7320
      TabIndex        =   8
      Top             =   120
      Width           =   495
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBGrid1 
      Bindings        =   "Produk_frm.frx":1982
      Height          =   5295
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   7575
      _Version        =   196614
      BevelColorFace  =   192
      AllowUpdate     =   0   'False
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   3200
      Columns(0).Caption=   "JENIS PRODUK"
      Columns(0).Name =   "JENIS PRODUK"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Jenis_Produk"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   6244
      Columns(1).Caption=   "NAMA PRODUK"
      Columns(1).Name =   "NAMA PRODUK"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Nama_Produk"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Caption=   "HARGA"
      Columns(2).Name =   "HARGA"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Harga"
      Columns(2).DataType=   5
      Columns(2).NumberFormat=   "###,###.00"
      Columns(2).FieldLen=   256
      _ExtentX        =   13361
      _ExtentY        =   9340
      _StockProps     =   79
      Caption         =   "TABEL PRODUK"
      ForeColor       =   16777215
      BackColor       =   8421504
   End
   Begin VB.Data Data1 
      BackColor       =   &H00000000&
      Caption         =   "DATA PRODUK"
      Connect         =   "Access"
      DatabaseName    =   " "
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      ForeColor       =   &H0000FF00&
      Height          =   345
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   " "
      Top             =   840
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   2640
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   2640
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   480
      Width           =   5175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2640
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   120
      Width           =   4575
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8280
      Top             =   3360
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
            Picture         =   "Produk_frm.frx":1996
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Produk_frm.frx":3328
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Produk_frm.frx":4CBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Produk_frm.frx":664C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Produk_frm.frx":7FDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Produk_frm.frx":9970
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Produk_frm.frx":A64A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Produk_frm.frx":B324
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   6780
      Left            =   8070
      TabIndex        =   0
      Top             =   0
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   11959
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
            Caption         =   "Close"
            Object.ToolTipText     =   "Tutup"
            ImageIndex      =   7
         EndProperty
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "Produk_frm.frx":CCB6
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HARGA"
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
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NAMA PRODUK"
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
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "JENIS PRODUK"
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
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Produk_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tambah As Boolean
Dim cek As Boolean
Dim cek2 As Boolean

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command1_Click()
Me.Visible = False
Jenis_Produk.Show
End Sub

Private Sub Data1_Reposition()
isi
End Sub

Private Sub Form_Activate()
Data1.Refresh
Data2.Refresh
isi
End Sub

Private Sub Form_Load()
Call Db_Produk
tutup
kosong
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Visible = True
End Sub

Private Sub SSDBGrid1_Click()
isi
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
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
        ISI_cmb
        Combo1.SetFocus
    Else
        simpan
    End If
Case 2
    If Toolbar1.Buttons(2).Caption = "Edit" Then
        If Text1 <> "" Then
            buka
            ISI_cmb
            isi
            Text2 = Format(Text2, "###")
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
Text1 = ""
Text2 = 0
End Sub

Sub isi()
If Data1.Enabled = True Then
kosong
With Data1.Recordset
    If Not .BOF Then
        Combo1 = !Jenis_Produk
        Text1 = !nama_produk
        Text2 = Format(!harga, "###,###.00")
    End If
End With
End If
End Sub

Sub ISI_cmb()
Combo1.Clear
With Data2.Recordset
If Not .BOF And Combo1.Enabled = True Then
    .MoveFirst
    Do While Not .EOF
        Combo1.AddItem !Jenis_Produk
        .MoveNext
    Loop
    Combo1.ListIndex = 0
End If
End With
End Sub

Sub tutup()
    Text1.Enabled = False
    Text2.Enabled = False
    Combo1.Enabled = False
End Sub

Sub buka()
    Combo1.Enabled = True
    Text1.Enabled = True
    Text2.Enabled = True
End Sub

Sub simpan()
Cek_Input
If cek2 = False Then
    MsgBox "Input tidak valid, mohon diperiksa kembali", vbInformation, "Validasi Input"
Else
    With Data1.Recordset
        If tambah = True Then
            cek_tambah
            If cek = False Then
                Data1.Refresh
                .AddNew
                !Jenis_Produk = Combo1
                !nama_produk = Text1
                !harga = Text2
                .Update
                tutup
                cmd_awal
            Else
                MsgBox "Data sudah ada,silahkan isi yang lain...", vbInformation, "Validasi Data"
            End If
        Else
            .Edit
            !Jenis_Produk = Combo1
            !nama_produk = Text1
            !harga = Text2
            .Update
            tutup
            cmd_awal
        End If
    End With
    Data1.Refresh
End If
End Sub

Sub Cek_Input()
cek2 = False
If Text1 = "" Or Text2 = "" Or Combo1 = "" Then
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
        If Text1 = !nama_produk And Combo1 = !Jenis_Produk Then
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
        x = MsgBox("Apakah anda yakin menghapus data?", vbYesNo, "Hapus Data")
        If x = vbYes Then
            .Delete
            kosong
            Data1.Refresh
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

