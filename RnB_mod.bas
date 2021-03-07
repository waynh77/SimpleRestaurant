Attribute VB_Name = "RnB_mod"
Dim db As String
Public PSWD As Boolean
Public prd As Boolean
Public lap As Boolean
Public sett As Boolean
Public use As Boolean
Public dat As Boolean

Sub buka_Db()
db = App.Path + "\RnB.mdb"
End Sub

Sub db_kontrol()
buka_Db
With Form1
    .Data1.DatabaseName = db
    .Data1.RecordSource = "select nama_table from temp_penjualan group by nama_table"
End With
End Sub

Sub Db_Produk()
buka_Db
With Produk_frm
.Data1.DatabaseName = db
.Data2.DatabaseName = db
.Data1.RecordSource = "select * from produk order by jenis_produk asc,nama_produk asc"
.Data2.RecordSource = "tbl_jenisproduk"
End With
End Sub

Sub db_jenisproduk()
buka_Db
With Jenis_Produk
.Data1.DatabaseName = db
.Data1.RecordSource = "tbl_jenisproduk"
End With
End Sub

Sub db_Trans()
buka_Db
With Trans_frm
.Data1.DatabaseName = db
.Data2.DatabaseName = db
.Data3.DatabaseName = db
.Data4.DatabaseName = db
.Data5.DatabaseName = db
.Data1.RecordSource = "temp_penjualan"
.Data2.RecordSource = "produk"
.Data3.RecordSource = "penjualan"
.Data4.RecordSource = "tbl_JENISproduk"
.Data5.RecordSource = "SETTING"
End With
End Sub

Sub db_LAp()
buka_Db
With lap_frm
.Data1.DatabaseName = db
End With
End Sub

Sub db_set()
buka_Db
With Setting_frm
.Data1.DatabaseName = db
.Data1.RecordSource = "setting"
End With
End Sub

Sub DB_user()
buka_Db
With User_frm
.Data1.DatabaseName = db
.Data1.RecordSource = "user"
End With
End Sub

Public Sub DB_Login()
buka_Db
frmLogin.Data1.DatabaseName = db
frmLogin.Data1.RecordSource = "user"
End Sub

Sub passwd()
'x = InputBox("Silahkan masukan password...", "Administrator")
'If x = "nauzan" Then
'    PSWD = True
'Else
'    PSWD = False
'End If
Form1.Visible = False
Pswd_frm.Show
End Sub

Public Sub DB_dATA()
buka_Db
Data_frm.Data1.DatabaseName = db
Data_frm.Data1.RecordSource = "select * from PENJUALAN order by no_trans desc, tgl desc, jam desc"
End Sub

