VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} Laporan_Penjualan 
   Caption         =   "Laporan Detil Penjualan"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10275
   Icon            =   "Laporan_Penjualan.dsx":0000
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   18124
   _ExtentY        =   15399
   SectionData     =   "Laporan_Penjualan.dsx":1982
End
Attribute VB_Name = "Laporan_Penjualan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tot As Double

Private Sub ActiveReport_ReportStart()
Field1.DataField = "tgl"
Field2.DataField = "jam"
Field3.DataField = "no_trans"
Field4.DataField = "jenis_produk"
Field5.DataField = "nama_produk"
Field6.DataField = "harga"
Field7.DataField = "qty"
Field8.DataField = "jumlah"
Field10.DataField = "nama_table"
tot = 0
End Sub

Private Sub Detail_Format()
tot = tot + Val(Format(Field8, "###.00"))
End Sub

Private Sub GroupFooter1_Format()
Field9 = Format(tot, "###,###.00")
End Sub
