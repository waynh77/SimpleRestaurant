VERSION 5.00
Begin VB.Form bayar_frm 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pembayaran"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5610
   ClipControls    =   0   'False
   Icon            =   "bayar_frm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5610
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "BATALKAN"
      Height          =   495
      Index           =   1
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2640
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "LANJUTKAN"
      Height          =   495
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   2160
      TabIndex        =   5
      Text            =   " "
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   2160
      TabIndex        =   4
      Text            =   " "
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   2160
      TabIndex        =   3
      Text            =   " "
      Top             =   120
      Width           =   3375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   5520
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KEMBALIAN"
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
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BAYAR"
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
      Top             =   840
      Width           =   1935
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
      Index           =   2
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "bayar_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    If Val(Text2) <> 0 And Val(Text3) >= 0 Then
        Trans_frm.simpan_Trans
        cetak3
        Unload Me
        Trans_frm.hapus_temp
    Else
        MsgBox "Maaf jumlah pembayaran tidak valid, mohon diperiksa kembali...", vbInformation, "Validasi Pembayaran"
        Text2.SetFocus
    End If
Case 1
    Unload Me
End Select
End Sub

Sub cetak3()
With Cetak_Transaksi
    Load Cetak_Transaksi
'    .Label1.Caption = "Transaksi " & Trans_frm.Caption
    .DAODataControl1.DatabaseName = Trans_frm.Data1.DatabaseName
    .DAODataControl1.RecordSource = Trans_frm.Data1.RecordSource
    .Field10.Text = Format(Val(Text2), "###,###.00")
    .Field11.Text = Format(Text3, "###,###.00")
    .Field7 = Format(Val(Trans_frm.Text5(3)), "###,###.00")
    .Field8 = Format(Val(Trans_frm.Text5(2)), "###,###.00")
    Cetak_Transaksi.PrintReport False
    Unload Cetak_Transaksi
    '.Show
End With
End Sub

Sub Cetak()
Dim tot As Double
Dim x, y As Single
Printer.PaperSize = 1
x = 200
y = 200
Printer.Height = x
Printer.Width = y
Printer.Font = "MS sans serif"
'Printer.Cls
'Printer.Caption = "TRANSAKSI " & Trans_frm.Caption
With Trans_frm.Data1.Recordset
.MoveFirst
Printer.Print
Printer.Print
Printer.FontBold = True
Printer.FontSize = 10
Printer.FontUnderline = True
Printer.Print ; "Transaksi "; Trans_frm.Caption
Printer.FontBold = False
Printer.FontSize = 8
Printer.FontUnderline = False
Printer.Print
Do While Not .EOF
    Printer.Print ; !Jenis_Produk & !nama_produk
    Printer.Print Tab(3); "   @Rp "; rkanan(!harga, "###,###,###");
    Printer.Print Tab(23); "  Qty : "; !qty;
    Printer.Print Tab(33); "   Rp "; rkanan(!jumlah, "###,###,###")
    tot = tot + !jumlah
    Printer.Height = x + 100
    Printer.Width = y + 100
    .MoveNext
Loop
Printer.Print
Printer.FontBold = True
Printer.Print ; "Sub Total ";
Printer.FontUnderline = True
Printer.Print Tab(33); "   Rp "; rkanan(tot, "###,###,###")
Printer.FontUnderline = False
Printer.KillDoc
End With
End Sub

Sub cetak2()
Dim tot As Double
tampil.Show
tampil.Font = "MS sans serif"
tampil.Cls
tampil.Caption = "TRANSAKSI " & Trans_frm.Caption
With Trans_frm.Data1.Recordset
.MoveFirst
tampil.Print
tampil.Print
tampil.FontBold = True
tampil.FontSize = 10
tampil.FontUnderline = True
tampil.Print ; "Transaksi "; Trans_frm.Caption
tampil.FontBold = False
tampil.FontSize = 8
tampil.FontUnderline = False
tampil.Print
Do While Not .EOF
    tampil.Print ; !Jenis_Produk & !nama_produk
    tampil.Print Tab(3); "   @Rp "; rkanan(!harga, "###,###,###");
    tampil.Print Tab(23); "  Qty : "; !qty;
    tampil.Print Tab(33); "   Rp "; rkanan(!jumlah, "###,###,###")
    tot = tot + !jumlah
    .MoveNext
Loop
tampil.Print
tampil.FontBold = True
tampil.Print ; "Sub Total ";
tampil.FontUnderline = True
tampil.Print Tab(33); "   Rp "; rkanan(tot, "###,###,###")
tampil.FontUnderline = False
'tampil.Print Tab(5); "Tax & Service ("; frmseting.Data1.Recordset!tax; "%)";
'tampil.Print Tab(33); "   Rp "; rkanan(frmseting.Data1.Recordset!tax * Label17 / 100, "###,###,###")
'tampil.Print Tab(5); "Discount ( "; Text1; "%)";
'tampil.FontUnderline = True
'tampil.Print Tab(33); "   Rp "; rkanan(Text1 * Label17 / 100, "###,###,###")
'tampil.FontUnderline = False
'tampil.Print Tab(15); "Grand Total ";
'tampil.FontUnderline = True
'tampil.Print Tab(33); "   Rp "; rkanan(Label17 + frmseting.Data1.Recordset!tax * Label17 / 100 - Text1 * Label17 / 100, "###,###,###")
'tampil.FontBold = False
'tampil.FontUnderline = False
'tampil.Print ; "Kasir : "; User
End With
End Sub

Private Sub Form_Load()
Text1.Enabled = False
Text3.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Trans_frm.Enabled = True
'Trans_frm.Show
End Sub

Private Sub Text2_Change()
Text3 = Format(Val(Text2) - Val(Format(Text1, "###.00")), "###,###.00")
End Sub

Private Function rkanan(ndata, cformat) As String
    rkanan = Format(ndata, cformat)
    rkanan = Space(Len(cformat) - Len(rkanan)) + rkanan
End Function

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        Command1_Click (0)
    End If
End Sub
