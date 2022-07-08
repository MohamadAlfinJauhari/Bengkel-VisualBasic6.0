VERSION 5.00
Begin VB.Form FormMenuUtama 
   BackColor       =   &H00808000&
   Caption         =   "SISTEM INFORMASI BENGKEL KOMPAK"
   ClientHeight    =   10635
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   15615
   BeginProperty Font 
      Name            =   "Arial Rounded MT Bold"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   10635
   ScaleWidth      =   15615
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   11535
      Left            =   0
      Picture         =   "FormMenuUtama.frx":0000
      Top             =   -240
      Width           =   20490
   End
   Begin VB.Menu data 
      Caption         =   "DATA"
      Begin VB.Menu datakaryawan 
         Caption         =   "DATA KARYAWAN"
      End
      Begin VB.Menu datasupplier 
         Caption         =   "DATA SUPPLIER"
      End
      Begin VB.Menu datapelanggan 
         Caption         =   "DATA PELANGGAN"
      End
      Begin VB.Menu persediaan 
         Caption         =   "DATA PERSEDIAN"
      End
   End
   Begin VB.Menu transaksi 
      Caption         =   "TRANSAKSI"
      Begin VB.Menu transaksipenjualan 
         Caption         =   "TRANSAKSI PENJUALAN"
      End
      Begin VB.Menu transaksipembelian 
         Caption         =   "TRANSAKSI PEMBELIAN"
      End
      Begin VB.Menu servicekendaraan 
         Caption         =   "TRANSAKSI SERVIS KENDARAAN"
      End
   End
   Begin VB.Menu laporan 
      Caption         =   "LAPORAN"
      Begin VB.Menu laporanpersediaan 
         Caption         =   "LAPORAN DATA PERSEDIAAN"
         Index           =   1
      End
      Begin VB.Menu laporankaryawan 
         Caption         =   "LAPORAN DATA KARYAWAN"
         Index           =   2
      End
      Begin VB.Menu laporansupplier 
         Caption         =   "LAPORAN DATA SUPPLIER"
      End
      Begin VB.Menu laporanpelanggan 
         Caption         =   "LAPORAN DATA PELANGGAN"
      End
      Begin VB.Menu laporanjual 
         Caption         =   "LAPORAN PENJUALAN SPAREPART"
      End
      Begin VB.Menu laporanbeli 
         Caption         =   "LAPORAN PEMBELIAN SPAREPART"
      End
      Begin VB.Menu laporanservice 
         Caption         =   "LAPORAN SERVIS KENDARAAN"
      End
   End
   Begin VB.Menu keluar 
      Caption         =   "KELUAR"
   End
End
Attribute VB_Name = "FormMenuUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub datakaryawan_Click()
Load FormKaryawan
FormKaryawan.Show
End Sub

Private Sub datapelanggan_Click()
Load FormInputPelanggan
FormInputPelanggan.Show
End Sub

Private Sub datasupplier_Click()
Load FormSupplier
FormSupplier.Show
End Sub

Private Sub Form_Load()
WindowState = 2
End Sub

Private Sub keluar_Click()
End
End Sub

Private Sub laporanbeli_Click()
Load FormLaporanTransaksiPembelian
FormLaporanTransaksiPembelian.Show
End Sub

Private Sub laporanjual_Click()
Load FormLaporanTransaksiPenjualan
FormLaporanTransaksiPenjualan.Show
End Sub

Private Sub laporankaryawan_Click(Index As Integer)
Load FormLaporanDataKaryawan
FormLaporanDataKaryawan.Show
End Sub

Private Sub laporanpelanggan_Click()
Load FormLaporanDataCustomer
FormLaporanDataCustomer.Show
End Sub

Private Sub laporanpersediaan_Click(Index As Integer)
Load FormLaporanDataPersediaan
FormLaporanDataPersediaan.Show
End Sub

Private Sub laporanservice_Click()
Load FormLaporanTransaksiService
FormLaporanTransaksiService.Show
End Sub

Private Sub laporansupplier_Click()
Load FormLaporanDataSupplier
FormLaporanDataSupplier.Show
End Sub
Private Sub persediaan_Click()
Load FormPersediaan
FormPersediaan.Show
End Sub

Private Sub servicekendaraan_Click()
Load FormService
FormService.Show
End Sub

Private Sub transaksipembelian_Click()
Load FormPembelianSparepart
FormPembelianSparepart.Show
End Sub

Private Sub transaksipenjualan_Click()
Load FormPenjualanSparepart
FormPenjualanSparepart.Show
End Sub
