VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form6 
   BackColor       =   &H00808000&
   Caption         =   "INPUT PEMBELIAN SPAREPART"
   ClientHeight    =   9015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15150
   LinkTopic       =   "Form1"
   ScaleHeight     =   9015
   ScaleWidth      =   15150
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FormPembelianSparepart.frx":0000
      Height          =   2175
      Left            =   360
      TabIndex        =   37
      Top             =   6600
      Width           =   19935
      _ExtentX        =   35163
      _ExtentY        =   3836
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "no_faktur"
         Caption         =   "no_faktur"
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
         DataField       =   "tanggal_transaksi"
         Caption         =   "tanggal_transaksi"
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
         DataField       =   "kode_supplier"
         Caption         =   "kode_supplier"
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
         DataField       =   "nama_supplier"
         Caption         =   "nama_supplier"
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
         DataField       =   "alamat_supplier"
         Caption         =   "alamat_supplier"
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
         DataField       =   "telepon"
         Caption         =   "telepon"
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
      BeginProperty Column06 
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
      BeginProperty Column07 
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
      BeginProperty Column08 
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
      BeginProperty Column09 
         DataField       =   "kuantitas"
         Caption         =   "kuantitas"
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
      BeginProperty Column10 
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
      BeginProperty Column11 
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
      BeginProperty Column12 
         DataField       =   "total"
         Caption         =   "total"
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
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1289,764
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1289,764
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1739,906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   12120
      Top             =   5520
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
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
      Connect         =   "DSN=myodbcbengkel"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "myodbcbengkel"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tb_beli"
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
   Begin VB.CommandButton cmdtambah 
      Caption         =   "Tambah"
      Height          =   495
      Left            =   1920
      TabIndex        =   36
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00808000&
      Height          =   1935
      Left            =   12120
      TabIndex        =   29
      Top             =   720
      Width           =   2895
      Begin VB.Label Label16 
         BackColor       =   &H8000000E&
         Caption         =   "Nama Jalan"
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
         TabIndex        =   31
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label Label11 
         BackColor       =   &H00808000&
         Caption         =   "Nama Bengkel"
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
         Left            =   600
         TabIndex        =   30
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdhapus 
      BackColor       =   &H000000C0&
      Caption         =   "Clear"
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
      Left            =   3360
      MaskColor       =   &H000000C0&
      TabIndex        =   28
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdkeluar 
      BackColor       =   &H8000000E&
      Caption         =   "Close"
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
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   26
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdsimpan 
      Caption         =   "Simpan"
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
      Left            =   480
      TabIndex        =   25
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox ttotal 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   9000
      TabIndex        =   23
      Top             =   5760
      Width           =   2895
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000E&
      Caption         =   "Data Sparepart"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   6240
      TabIndex        =   5
      Top             =   1560
      Width           =   5655
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
         Height          =   375
         Left            =   2040
         TabIndex        =   35
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox tjual 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2040
         TabIndex        =   34
         Top             =   2880
         Width           =   1695
      End
      Begin VB.ComboBox cbobarang 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2040
         TabIndex        =   22
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox tbeli 
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
         Left            =   2040
         TabIndex        =   20
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox tjml 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2040
         TabIndex        =   19
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox tnamabrg 
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
         Left            =   2040
         TabIndex        =   18
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000E&
         Caption         =   "Harga Jual"
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
         Left            =   480
         TabIndex        =   33
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label15 
         BackColor       =   &H8000000E&
         Caption         =   "Harga Beli :"
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
         Left            =   480
         TabIndex        =   27
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000E&
         Caption         =   "Satuan :"
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
         Left            =   480
         TabIndex        =   21
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000E&
         Caption         =   "Jumlah Beli :"
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
         Left            =   480
         TabIndex        =   17
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000E&
         Caption         =   "Nama Barang :"
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
         Left            =   480
         TabIndex        =   16
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000E&
         Caption         =   "Kode Barang :"
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
         Left            =   480
         TabIndex        =   15
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000E&
      Caption         =   "Data Supplier"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   5775
      Begin VB.TextBox tnamasup 
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
         Left            =   2160
         TabIndex        =   13
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox ttelp 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2160
         TabIndex        =   12
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox talamat 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2160
         TabIndex        =   11
         Top             =   1440
         Width           =   2535
      End
      Begin VB.ComboBox cbosupplier 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2160
         TabIndex        =   10
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000E&
         Caption         =   "Telepon :"
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
         Left            =   480
         TabIndex        =   9
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         Caption         =   "Nama Supplier :"
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
         Left            =   480
         TabIndex        =   8
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "Alamat :"
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
         Left            =   480
         TabIndex        =   7
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "Kode Supplier :"
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
         Left            =   480
         TabIndex        =   6
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   11535
      Begin VB.TextBox ttgl 
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
         Left            =   8760
         TabIndex        =   14
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox tfaktur 
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
         Left            =   1920
         TabIndex        =   2
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808000&
         Caption         =   "Tanggal Transaksi :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   6840
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808000&
         Caption         =   "No.Faktur :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Label Label17 
      BackColor       =   &H8000000E&
      Caption         =   "INPUT PEMBELIAN SPAREPART"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   32
      Top             =   240
      Width           =   5055
   End
   Begin VB.Label Label12 
      BackColor       =   &H8000000E&
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7560
      TabIndex        =   24
      Top             =   5760
      Width           =   1095
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub displaydata(prs As ADODB.Recordset)
tfaktur.Text = prs!no_faktur
ttgl.Text = Format(Str(prs!tanggal_transaksi), "dd-mm-yyyy")
cbosupplier.Text = prs!kode_supplier
tnamasup.Text = prs!nama_supplier
talamat.Text = prs!alamat_supplier
ttelp.Text = prs!telepon
cbobarang.Text = prs!kode_barang
tnamabrg.Text = prs!nama_barang
tsatuan.Text = prs!satuan
tjml.Text = prs!kuantitas
tbeli.Text = prs!harga_beli
tjual.Text = prs!harga_jual
ttotal.Text = prs!total
End Sub

Private Sub cmdsimpan_Click()
Adodc1.Recordset.AddNew
Adodc1.Recordset!no_faktur = tfaktur.Text
Adodc1.Recordset!tanggal_transaksi = ttgl.Text
Adodc1.Recordset!kode_supplier = cbosupplier.Text
Adodc1.Recordset!nama_supplier = tnamasup.Text
Adodc1.Recordset!alamat_supplier = talamat.Text
Adodc1.Recordset!telepon = ttelp.Text
Adodc1.Recordset!kode_barang = cbobarang.Text
Adodc1.Recordset!nama_barang = tnamabrg.Text
Adodc1.Recordset!satuan = tsatuan.Text
Adodc1.Recordset!kuantitas = tjml.Text
Adodc1.Recordset!harga_beli = tbeli.Text
Adodc1.Recordset!harga_jual = tjual.Text
Adodc1.Recordset!total = ttotal.Text

'Set rs5 = New ADODB.Recordset
'rs5.CursorLocation = adUseServer
'rs5.Open "update tb_persediaan,tb_beli set tb_persediaan.jumlah = tb_persediaan.jumlah + tb_beli.kuantitas where tb_persediaan.kode_barang=tb_beli.kode_barang", dbkoneksi, adOpenDynamic, adLockOptimistic
End Sub

Private Sub cmdtambah_Click()
tfaktur.Text = ""
ttgl.Text = ""
cbosupplier.Text = ""
tnamasup.Text = ""
talamat.Text = ""
ttelp.Text = ""
cbobarang.Text = ""
tnamabrg.Text = ""
tsatuan.Text = ""
tjml.Text = ""
tbeli.Text = ""
tjual.Text = ""
ttotal.Text = ""
End Sub

Private Sub cmdhapus_Click()
Adodc1.Recordset.Delete
Adodc1.Refresh
End Sub

Private Sub cmdkeluar_Click()
Load Form9
Form9.Show
End Sub

Private Sub tfaktur_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
ttgl.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub ttgl_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
cbosupplier.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub cbosupplier_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
tnamasup.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub tnamasup_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
talamat.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub talamat_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
ttelp.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub ttelp_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
cbobarang.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub cbobarang_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
tnamabrg.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub tnamabrg_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
tsatuan.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub tsatuan_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
tjml.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub tbeli_LostFocus()
tbeli.Text = Format(tbeli.Text, "#########")
End Sub
Private Sub ttotal_LostFocus()
ttotal.Text = Format(Str(Val(tbeli.Text) * Val(normalize(tjml.Text))), "############")
End Sub
Private Sub tjual_LostFocus()
tjual.Text = Format(tjual.Text, "#########")
End Sub
Private Sub tjml_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
tbeli.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub tbeli_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
tjual.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub tjual_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
ttotal.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub cbosupplier_Click()
If cbosupplier.Text = "S01" Then
tnamasup.Text = "JAYA MAKMUR"
talamat.Text = "SEMARANG"
ttelp.Text = "089527627527"
ElseIf cbosupplier.Text = "S02" Then
tnamasup.Text = "JAYA ABADI"
talamat.Text = "SEMARANG"
ttelp.Text = "089543234567"
ElseIf cbosupplier.Text = "S03" Then
tnamasup.Text = "ABADI MOTOR"
talamat.Text = "SEMARANG"
ttelp.Text = "087654321456"
End If
End Sub
Private Sub cbobarang_Validate(Cancel As Boolean)
Call BukaDatabase
Dim rs5 As New ADODB.Recordset
rs5.CursorLocation = adUseClient
rs5.Open "select nama_barang,harga_beli from tb_persediaan where kode_barang = '" + cbobarang.Text + "'", dbkoneksi, adOpenDynamic, adLockOptimistic
If Not rs5.EOF Then
tnamabrg.Text = rs5!nama_barang
tbeli.Text = rs5!harga_beli
End If
Set rs5 = Nothing
Set dbkoneksi = Nothing
End Sub
Private Sub Form_Load()
Call BukaDatabase
Dim rs5 As New ADODB.Recordset
rs5.CursorLocation = adUseClient
rs5.Open "select kode_barang from tb_persediaan order by kode_barang lock in share mode", dbkoneksi, adOpenDynamic, adLockOptimistic
cbobarang.Clear
If Not rs5.EOF Then
rs5.MoveFirst
Do While Not rs5.EOF
cbobarang.AddItem rs5!kode_barang
rs5.MoveNext
Loop
End If
Set rs5 = Nothing
Set dbkoneksi = Nothing

cbosupplier.AddItem "S01"
cbosupplier.AddItem "S02"
cbosupplier.AddItem "S03"
End Sub
Private Sub tjual_Click()
tjual.Text = (tbeli * 110) / 100
End Sub

Private Sub ttotal_Click()
ttotal.Text = Val(tbeli.Text) * Val(tjml.Text)
End Sub


