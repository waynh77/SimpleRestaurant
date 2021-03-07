VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form Trans_frm 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaksi Penjualan"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14955
   Icon            =   "Trans_frm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   14955
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   2640
      TabIndex        =   25
      Text            =   "Combo3"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   9600
      TabIndex        =   24
      Text            =   " "
      Top             =   2160
      Width           =   4215
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   10560
      TabIndex        =   23
      Text            =   " "
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   10560
      TabIndex        =   22
      Text            =   " "
      Top             =   1560
      Width           =   3255
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   9600
      TabIndex        =   21
      Text            =   " "
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   9600
      TabIndex        =   20
      Text            =   " "
      Top             =   1560
      Width           =   855
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3840
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data Data4 
      Caption         =   "JENIS PRD"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7920
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CLOSING"
      Height          =   375
      Index           =   1
      Left            =   11160
      TabIndex        =   16
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Data Data3 
      Caption         =   "PENJUALAN"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8640
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data Data2 
      Caption         =   "PRODUK"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8280
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PRINT"
      Height          =   375
      Index           =   0
      Left            =   10200
      TabIndex        =   15
      Top             =   3480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   9600
      TabIndex        =   14
      Top             =   600
      Width           =   4215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   2640
      TabIndex        =   13
      Top             =   2640
      Width           =   4215
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   2640
      TabIndex        =   12
      Top             =   1920
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   2640
      TabIndex        =   11
      Top             =   1560
      Width           =   4215
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2640
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   960
      Width           =   4215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2640
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   600
      Width           =   4215
   End
   Begin VB.Data Data1 
      BackColor       =   &H00000000&
      Caption         =   "DATA TRANSAKSI PENJUALAN"
      Connect         =   "Access"
      DatabaseName    =   " "
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   7200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   " "
      Top             =   2640
      Width           =   3855
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
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
            Picture         =   "Trans_frm.frx":1982
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trans_frm.frx":3314
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trans_frm.frx":4CA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trans_frm.frx":6638
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trans_frm.frx":7FCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trans_frm.frx":995C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trans_frm.frx":A636
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trans_frm.frx":B310
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   7290
      Left            =   14145
      TabIndex        =   1
      Top             =   0
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   12859
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
      MouseIcon       =   "Trans_frm.frx":CCA2
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBGrid1 
      Bindings        =   "Trans_frm.frx":CFBC
      Height          =   3975
      Left            =   240
      TabIndex        =   0
      Top             =   3120
      Width           =   13695
      _Version        =   196614
      BevelColorFace  =   192
      AllowUpdate     =   0   'False
      RowHeight       =   423
      Columns.Count   =   5
      Columns(0).Width=   5186
      Columns(0).Caption=   "JENIS PRODUK"
      Columns(0).Name =   "JENIS PRODUK"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Jenis_Produk"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   7938
      Columns(1).Caption=   "NAMA PRODUK"
      Columns(1).Name =   "NAMA PRODUK"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Nama_Produk"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3836
      Columns(2).Caption=   "HARGA"
      Columns(2).Name =   "HARGA"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Harga"
      Columns(2).DataType=   5
      Columns(2).NumberFormat=   "###,###.00"
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Caption=   "QTY"
      Columns(3).Name =   "QTY"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "qty"
      Columns(3).DataType=   4
      Columns(3).NumberFormat=   "###"
      Columns(3).FieldLen=   256
      Columns(4).Width=   3200
      Columns(4).Caption=   "JUMLAH"
      Columns(4).Name =   "JUMLAH"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "jumlah"
      Columns(4).DataType=   5
      Columns(4).NumberFormat=   "###,###.00"
      Columns(4).FieldLen=   256
      _ExtentX        =   24156
      _ExtentY        =   7011
      _StockProps     =   79
      Caption         =   "TABEL TRANSAKSI PENJUALAN"
      ForeColor       =   16777215
      BackColor       =   8421504
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   7200
      X2              =   13800
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   7200
      X2              =   13800
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TOTAL"
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
      Index           =   9
      Left            =   7200
      TabIndex        =   19
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DISCOUNT"
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
      Index           =   8
      Left            =   7200
      TabIndex        =   18
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TAX AND SERVICE"
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
      Index           =   7
      Left            =   7200
      TabIndex        =   17
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   240
      X2              =   6840
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   6840
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SUB TOTAL"
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
      Index           =   6
      Left            =   7200
      TabIndex        =   9
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HARGA * QTY"
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
      Index           =   5
      Left            =   240
      TabIndex        =   8
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "JUMLAH / QTY"
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
      Index           =   4
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HARGA SATUAN"
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
      Index           =   3
      Left            =   240
      TabIndex        =   6
      Top             =   1560
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
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   960
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
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TRANSAKSI MEJA/BAR"
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
      TabIndex        =   2
      Top             =   120
      Width           =   13575
   End
End
Attribute VB_Name = "Trans_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tambah As Boolean
Dim cek As Boolean
Dim cek2 As Boolean
Dim nomor As String

Private Sub Combo1_Click()
ISI_cmb2
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo2_Click()
isi_hrg
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Sub isi_cmb3()
Dim x As Byte
Combo3.Clear
x = 0
Do Until x = 100
    x = x + 1
    Combo3.AddItem x
Loop
Combo3.ListIndex = 0
End Sub

Private Sub Combo3_Click()
Text3 = Format(Val(Format(Combo3, "###")) * Val(Format(Text1, "###.00")), "###,###.00")
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    If Data1.Recordset.BOF Then
        MsgBox "Maaf belum ada transaksi...", vbCritical, "Validasi Data"
    Else
        Me.Enabled = False
        bayar_frm.Text1 = Text4
        bayar_frm.Show
    End If
Case 1
    If Data1.Recordset.BOF Then
        MsgBox "Maaf belum ada transaksi...", vbCritical, "Validasi Data"
    Else
        Me.Enabled = False
        bayar_frm.Text1 = Text5(4)
        bayar_frm.Show
'        simpan_Trans
    End If
End Select
End Sub

Private Sub Data1_Reposition()
isi
End Sub

Private Sub Form_Activate()
kosong
Data1.RecordSource = "select * from temp_penjualan where nama_table='" & Me.Caption & "' ORDER BY jenis_produk"
HIT_ttl
Data1.Refresh
isi
isi_cmb3
End Sub

Private Sub Form_Load()
Call db_Trans
tutup
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Visible = True
End Sub

Private Sub SSDBGrid1_Click()
isi
End Sub

Private Sub combo3_KeyPress(KeyAscii As Integer)
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
            Combo3 = Format(Combo3, "###")
            tambah = False
            cmd_simpan
        Else
            MsgBox "Data Kosong", vbInformation, "Validasi Data"
        End If
    Else
        kosong
        cmd_awal
        tutup
        isi
    End If
Case 3
    hapus
Case 4
Case 5
Case 6
Case 7
    Unload Me
End Select
End Sub

Sub kosong()
Text1 = ""
Combo3 = 1
Text3 = 0
Text4 = 0
Text5(2) = 0
Text5(3) = 0
Text5(4) = 0
Combo1 = ""
Combo2 = ""
End Sub

Sub isi()
If Data1.Enabled = True Then
With Data1.Recordset
    If Not .BOF Then
        Combo1 = !Jenis_Produk
        Combo2 = !nama_produk
        Text1 = Format(!harga, "###,###.00")
        Combo3 = Format(!qty, "###")
        Text3 = Format(!harga * !qty, "###,###.00")
    End If
End With
End If
End Sub

Sub ISI_cmb()
Combo1.Clear
With Data4.Recordset
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

Sub ISI_cmb2()
Combo2.Clear
Data2.RecordSource = "select * from produk where jenis_produk = '" & Combo1 & "'"
Data2.Refresh
With Data2.Recordset
If Not .BOF And Combo1.Enabled = True Then
    .MoveFirst
    Do While Not .EOF
        Combo2.AddItem !nama_produk
        .MoveNext
    Loop
    Combo2.ListIndex = 0
End If
End With
End Sub

Sub isi_hrg()
Data2.RecordSource = "select * from produk where jenis_produk='" & Combo1 & "' and nama_produk='" & Combo2 & "'"
Data2.Refresh
Text1 = Format(Data2.Recordset!harga, "###,###.00")
Text3 = Format(Val(Format(Combo3, "###")) * Val(Format(Text1, "###.00")), "###,###.00")
End Sub

Sub tutup()
    Text1.Enabled = False
    Combo3.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Combo1.Enabled = False
    Combo2.Enabled = False
    Text5(0).Enabled = False
    Text5(1).Enabled = False
    Text5(2).Enabled = False
    Text5(3).Enabled = False
    Text5(4).Enabled = False
End Sub

Sub buka()
    Combo1.Enabled = True
    Combo2.Enabled = True
    Combo3.Enabled = True
End Sub

Sub simpan()
cek_input
If cek2 = False Then
    MsgBox "Input tidak valid, mohon diperiksa kembali", vbInformation, "Validasi Input"
Else
    With Data1.Recordset
        If tambah = True Then
            .AddNew
            !Jenis_Produk = Combo1
            !nama_produk = Combo2
            !harga = Text1
            !qty = Combo3
            !nama_table = Me.Caption
            !jumlah = !qty * !harga
            .Update
            tutup
            cmd_awal
        Else
            .Edit
            !Jenis_Produk = Combo1
            !nama_produk = Combo2
            !harga = Text1
            !qty = Combo3
            !nama_table = Me.Caption
            !jumlah = !qty * !harga
            .Update
            tutup
            cmd_awal
        End If
    End With
    HIT_ttl
    Data1.Refresh
End If
End Sub

Sub simpan_Trans()
Data1.Enabled = False
Inv_auto
With Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        With Data3.Recordset
            .AddNew
            !tgl = Date
            !jam = Time
            !Jenis_Produk = Data1.Recordset!Jenis_Produk
            !nama_produk = Data1.Recordset!nama_produk
            !harga = Data1.Recordset!harga
            !qty = Data1.Recordset!qty
            !nama_table = Data1.Recordset!nama_table
            !jumlah = Data1.Recordset!jumlah
            !no_trans = nomor
            !tax = Text5(0)
            !disc = Text5(1)
            !jml_tax = Val(Format(Text5(2), "###.##"))
            !jml_disc = Val(Format(Text5(3), "###.##"))
            !total = Val(Format(Text5(4), "###.##"))
            !User = Mid(Form1.Label1.Caption, 13)
            .Update
        End With
        .MoveNext
    Loop
    MsgBox "Transaksi telah disimpan dengan nomor " & nomor, vbInformation, "Closing"
    'hapus_temp
    Cetak_Transaksi.Label18.Caption = "No." & nomor
    kosong
Else
    MsgBox "Maaf data masih kosong...", vbInformation, "Validasi Data"
End If
End With
Data1.Enabled = True
End Sub

Sub cek_input()
cek2 = False
If Combo3 = "" Or Combo1 = "" Or Combo2 = "" Then
    cek2 = False
Else
    cek2 = True
End If
End Sub

Sub hapus()
With Data1.Recordset
    If Not .BOF Then
        x = MsgBox("Apakah anda yakin menghapus data?", vbYesNo, "Hapus Data")
        If x = vbYes Then
            .Delete
            kosong
            HIT_ttl
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

Sub HIT_ttl()
Dim tot As Double
Dim disc As Double
Dim tax As Double
Data1.Enabled = False
Data1.Refresh
tot = 0
With Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        tot = tot + !jumlah
        .MoveNext
    Loop
End If
End With
Text4 = Format(tot, "###,###.00")
Data1.Enabled = True
With Data5.Recordset
If Not .BOF Then
    Text5(0) = !tax
    Text5(1) = !disc
Else
    .AddNew
    !tax = 0
    !disc = 0
    .Update
    Text5(0) = 0
    Text5(1) = 0
End If
tax = tot * !tax / 100
disc = tot * !disc / 100
End With
Text5(2) = Format(tax, "###,###.00")
Text5(3) = Format(disc, "###,###.00")
Text5(4) = Format(tot + tax - disc, "###,###.00")
End Sub

Sub Inv_auto()
Dim urutan As String * 10
Dim hitung As Single
Data3.RecordSource = "select * from penjualan order by no_trans asc"
Data3.Refresh
With Data3.Recordset
    If .RecordCount = 0 Then
        urutan = "INV" & "0000001"
    Else
        .MoveLast
'        If Val(Left(.Fields("No_trans"), 7)) <> "0000000" Then
            hitung = Val(Right(.Fields("no_trans"), 7)) + 1
            urutan = "INV" & Right("0000000" & hitung, 7)
'        End If
    End If
    nomor = urutan
End With
End Sub

Sub hapus_temp()
Dim C As Single
Data1.Enabled = False
C = Data1.Recordset.RecordCount
Data1.Refresh
With Data1.Recordset
    If Not .BOF Then
        .MoveFirst
        Do Until C = 0
            .Delete
            C = C - 1
            .MoveNext
        Loop
        kosong
    End If
End With
Data1.Enabled = True
Data1.Refresh
End Sub

