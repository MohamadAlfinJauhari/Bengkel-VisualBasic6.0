VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormPersediaan 
   BackColor       =   &H00808000&
   Caption         =   "INPUT PERSEDIAAN"
   ClientHeight    =   9195
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15510
   LinkTopic       =   "Form1"
   ScaleHeight     =   9195
   ScaleWidth      =   15510
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H00808000&
      Height          =   1815
      Left            =   360
      TabIndex        =   20
      Top             =   360
      Width           =   3495
      Begin VB.Label Label11 
         BackColor       =   &H00808000&
         Caption         =   "Bengkel Kompak"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   1215
         Left            =   1560
         TabIndex        =   22
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label16 
         BackColor       =   &H8000000E&
         Caption         =   "Jln. Merdeka No.77A Semarang"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1560
         Width           =   3135
      End
      Begin VB.Image Image1 
         Height          =   1305
         Left            =   120
         Picture         =   "FormPersediaan.frx":0000
         Top             =   240
         Width           =   1305
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FormPersediaan.frx":15C7
      Height          =   2415
      Left            =   3240
      TabIndex        =   18
      Top             =   5400
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   4260
      _Version        =   393216
      BackColor       =   65535
      HeadLines       =   1
      RowHeight       =   23
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "kode_barang"
         Caption         =   "kode_barang"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3072
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "nama_barang"
         Caption         =   "nama_barang"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3072
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "satuan"
         Caption         =   "satuan"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3072
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "harga_beli"
         Caption         =   "harga_beli"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3072
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "harga_jual"
         Caption         =   "harga_jual"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3072
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "jumlah"
         Caption         =   "jumlah"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3072
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   959,811
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739,906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   14400
      Top             =   7200
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=bengkel"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "bengkel"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tb_persediaan"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox thargajual 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   17
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00808000&
      Height          =   855
      Left            =   5160
      TabIndex        =   5
      Top             =   8040
      Width           =   6015
      Begin VB.CommandButton tbkeluar 
         Caption         =   "KELUAR"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4680
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton tbhapus 
         Caption         =   "HAPUS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton tbtambah 
         Caption         =   "TAMBAH"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton tbsimpan 
         Caption         =   "SIMPAN"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000E&
      Caption         =   "Data Barang"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   4080
      TabIndex        =   0
      Top             =   1320
      Width           =   7575
      Begin VB.TextBox tsatuan 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3360
         TabIndex        =   14
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox tkodebarang 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3360
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox thargabeli 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         TabIndex        =   12
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox tjumlah 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3360
         TabIndex        =   11
         Top             =   3240
         Width           =   1695
      End
      Begin VB.TextBox tnamabarang 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3360
         TabIndex        =   10
         Top             =   840
         Width           =   3855
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000E&
         Caption         =   "HARGA JUAL :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "SATUAN :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000E&
         Caption         =   "JUMLAH :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   4
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         Caption         =   "HARGA BELI :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "NAMA BARANG :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "KODE BARANG :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "DATA PERSEDIAAN SPAREPART"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   19
      Top             =   480
      Width           =   6615
   End
End
Attribute VB_Name = "FormPersediaan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
WindowState = 2
End Sub

Private Sub tbhapus_Click()
Cleartextbox FormPersediaan
End Sub

Private Sub tbkeluar_Click()
Unload Me
End Sub

Private Sub tbsimpan_Click()
Dim remove As String
remove = Replace(Replace(tkodebarang.Text, "(", "", 1, -1, vbTextCompare), ")-", "", 1, -1, vbTextCompare)
tkodebarang.Text = remove
Call BukaDatabase
rs1.Open "select * from tb_persediaan where kode_barang is null", dbkoneksi, adOpenDynamic, adLockOptimistic
rs1.AddNew
rs1!kode_barang = Trim(tkodebarang.Text)
rs1!nama_barang = Trim(tnamabarang.Text)
rs1!satuan = Trim(tsatuan.Text)
rs1!harga_beli = Val(thargabeli.Text)
rs1!harga_jual = Val(thargajual.Text)
rs1!jumlah = Trim(tjumlah.Text)
rs1.Update
rs1.Close
Set rs1 = Nothing
Adodc1.Refresh
End Sub

Private Sub tbtambah_Click()
Cleartextbox FormPersediaan
tkodebarang.SetFocus
End Sub

Private Sub tkodebarang_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
tnamabarang.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub tnamabarang_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
tsatuan.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub tsatuan_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
thargabeli.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub thargajual_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
tjumlah.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub thargabeli_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
thargajual.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub thargajual_Click()
thargajual.Text = (thargabeli * 110) / 100
End Sub

