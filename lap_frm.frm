VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form lap_frm 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan Penjualan"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8145
   ClipControls    =   0   'False
   Icon            =   "lap_frm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   8145
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2520
      TabIndex        =   9
      Text            =   "Combo2"
      Top             =   120
      Width           =   3015
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   3720
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBGrid1 
      Bindings        =   "lap_frm.frx":1982
      Height          =   4575
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   7695
      _Version        =   196614
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   13573
      _ExtentY        =   8070
      _StockProps     =   79
      Caption         =   "LAPORAN PENJUALAN"
   End
   Begin VB.Data Data1 
      Caption         =   "DATA LAPORAN"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1560
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "BATAL"
      Height          =   1335
      Index           =   1
      Left            =   6840
      Picture         =   "lap_frm.frx":1996
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "PREVIEW"
      Height          =   1335
      Index           =   0
      Left            =   5640
      Picture         =   "lap_frm.frx":2660
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   300
      Index           =   0
      Left            =   2520
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      Format          =   16515073
      CurrentDate     =   39914
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2520
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   600
      Width           =   3015
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   300
      Index           =   1
      Left            =   4080
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      Format          =   16515073
      CurrentDate     =   39914
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "STATUS LAPORAN"
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
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TANGGAL"
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
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "JENIS LAPORAN"
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
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
End
Attribute VB_Name = "lap_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tgl_awal As String
Dim tgl_akhir As String

Private Sub Combo1_Click()
If Combo1.ListIndex = 0 Then
    DTPicker1(1).Visible = False
Else
    DTPicker1(1).Visible = True
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command1_Click(Index As Integer)
Dim tgl_awal As String
Dim tgl_akhir As String
tgl_awal = Format(DTPicker1(0), "m/dd/yyyy")
tgl_akhir = Format(DTPicker1(1), "m/dd/yyyy")
Select Case Index
Case 0
    If Combo2.ListIndex = 0 Then
        cetak2
    Else
        CETAK1
    End If
Case 1
    Unload Me
End Select
End Sub

Sub CETAK1()
    If Combo1.ListIndex = 0 Then
        Data1.RecordSource = "select * from penjualan where cdate(tgl)='" & DTPicker1(0) & "' order by tgl asc,jam asc"
'        Data1.RecordSource = "select tgl,penjualan.NAMA_produk,sum(qty)as qty1,sum(jumlah)as jml,jml/qty1 as hrg from penjualan where cdate(tgl)='" & DTPicker1(0) & "' group by tgl,nama_produk"
        Data1.Refresh
        With Laporan_Penjualan
'        With Lap_Rekap
            .Label2.Caption = "Tanggal : " & Format(DTPicker1(0), "d mmmm yyyy")
            .DAODataControl1.DatabaseName = Data1.DatabaseName
            .DAODataControl1.RecordSource = Data1.RecordSource
            .Show
            .WindowState = 2
        End With
    Else
        Data1.RecordSource = "select * from penjualan where cdate(tgl)>='" & DTPicker1(0) & "' and cdate(tgl)<='" & DTPicker1(1) & "' order by tgl asc, jam asc"
'        Data1.RecordSource = "select tgl,penjualan.NAMA_produk,sum(qty)as qty1,sum(jumlah)as jml,jml/qty1 as hrg from penjualan where cdate(tgl)>='" & tgl_awal & "' and cdate(tgl)<='" & tgl_akhir & "' group by tgl,nama_produk"
        Data1.Refresh
        With Laporan_Penjualan
'        With Lap_Rekap
            .Label2.Caption = "Tanggal : " & Format(DTPicker1(0), "d mmmm yyyy") & " s/d " & Format(DTPicker1(1), "d mmmm yyyy")
            .DAODataControl1.DatabaseName = Data1.DatabaseName
            .DAODataControl1.RecordSource = Data1.RecordSource
            .Show
            .WindowState = 2
        End With
    End If
    Data1.Refresh
End Sub

Sub cetak2()
If Combo1.ListIndex = 0 Then
        CrystalReport1.ReportFileName = App.Path & "\LAPORAN REKAP PENJUALAN-harian.rpt"
        CrystalReport1.WindowTitle = "LAPORAN REKAP PENJUALAN HARIAN"
        CrystalReport1.SelectionFormula = "{penjualan.tgl}= date(" & Format(DTPicker1(0), "yyyy,mm,dd") & ")"
Else
        CrystalReport1.ReportFileName = App.Path & "\LAPORAN REKAP PENJUALAN.rpt"
        CrystalReport1.WindowTitle = "LAPORAN REKAP PENJUALAN"
        CrystalReport1.SelectionFormula = "{penjualan.tgl}>= date(" & Format(DTPicker1(0), "yyyy,mm,dd") & ") and {penjualan.tgl}<= date(" & Format(DTPicker1(1), "yyyy,mm,dd") & ")"
End If
CrystalReport1.RetrieveDataFiles
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Action = 1
End Sub

Private Sub Form_Load()
Call db_LAp
Combo1.Clear
Combo1.AddItem "Harian"
Combo1.AddItem "Periodik"
Combo1.ListIndex = 0
DTPicker1(0) = Date
DTPicker1(1) = Date
isi_cmb2
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Visible = True
End Sub

Sub isi_cmb2()
Combo2.Clear
Combo2.AddItem "REKAP"
Combo2.AddItem "DETIL"
Combo2.ListIndex = 0
End Sub
