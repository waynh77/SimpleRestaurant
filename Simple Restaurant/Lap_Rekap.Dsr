VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} Lap_Rekap 
   Caption         =   "Laporan Rekap Penjualan"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9615
   Icon            =   "Lap_Rekap.dsx":0000
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   16960
   _ExtentY        =   12938
   SectionData     =   "Lap_Rekap.dsx":3482
End
Attribute VB_Name = "Lap_Rekap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tot, tot1 As Double
Dim z As Single
Private Sub ActiveReport_ReportStart()
Field2.DataField = "nama_produk"
Field3.DataField = "hrg"
Field4.DataField = "qty1"
Field5.DataField = "jml"
Field8.DataField = "tgl"
tot = 0
z = 1
End Sub

Private Sub Detail_Format()
Field1.Text = z & "."
z = z + 1
tot = tot + Val(Format(Field4, "###"))
tot1 = tot1 + Val(Format(Field5, "###.00"))
End Sub

Private Sub GroupFooter1_Format()
Field6.Text = Format(tot, "###")
Field7.Text = Format(tot1, "###,###.00")
End Sub

