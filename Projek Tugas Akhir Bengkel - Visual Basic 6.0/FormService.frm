VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormService 
   BackColor       =   &H00808000&
   Caption         =   "INPUT TRANSAKSI SERVICE"
   ClientHeight    =   10485
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15525
   LinkTopic       =   "Form1"
   ScaleHeight     =   10485
   ScaleWidth      =   15525
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H00808000&
      Height          =   1815
      Left            =   13920
      TabIndex        =   42
      Top             =   240
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
         Index           =   1
         Left            =   1440
         TabIndex        =   44
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
         Index           =   1
         Left            =   120
         TabIndex        =   43
         Top             =   1560
         Width           =   3135
      End
      Begin VB.Image Image1 
         Height          =   1305
         Left            =   120
         Picture         =   "FormService.frx":0000
         Top             =   240
         Width           =   1305
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FormService.frx":15C7
      Height          =   2895
      Left            =   240
      TabIndex        =   41
      Top             =   7560
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   5106
      _Version        =   393216
      BackColor       =   65535
      HeadLines       =   1
      RowHeight       =   23
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14345
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14345
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   855
      Left            =   14040
      Top             =   8880
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      Connect         =   "DSN=bengkel"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "bengkel"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tb_persediaan"
      Caption         =   "Adodc2"
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "FormService.frx":15DC
      Height          =   2895
      Left            =   7200
      TabIndex        =   38
      Top             =   7560
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5106
      _Version        =   393216
      BackColor       =   65535
      HeadLines       =   1
      RowHeight       =   23
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14345
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14345
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000E&
      Caption         =   "Detail Service"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   240
      TabIndex        =   29
      Top             =   3720
      Width           =   6135
      Begin VB.TextBox tlayanan 
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
         Left            =   2160
         TabIndex        =   36
         Top             =   1560
         Width           =   3735
      End
      Begin VB.TextBox tbiayaservis 
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
         Left            =   2160
         TabIndex        =   32
         Top             =   2160
         Width           =   3735
      End
      Begin VB.TextBox tnamakaryawan 
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
         Left            =   2160
         TabIndex        =   31
         Top             =   960
         Width           =   3735
      End
      Begin VB.ComboBox cbokodekaryawan 
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
         TabIndex        =   30
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000E&
         Caption         =   "Jenis Layanan :"
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
         TabIndex        =   37
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000E&
         Caption         =   "Kode Karyawan :"
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
         Left            =   240
         TabIndex        =   35
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000E&
         Caption         =   "Nama Karyawan :"
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
         Left            =   240
         TabIndex        =   34
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000E&
         Caption         =   "Biaya Servis :"
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
         Index           =   0
         Left            =   240
         TabIndex        =   33
         Top             =   2280
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000E&
      Caption         =   "Data Penggantian Sparepart"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   6600
      TabIndex        =   16
      Top             =   1560
      Width           =   7095
      Begin VB.TextBox ttotalharga 
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
         Left            =   4560
         TabIndex        =   27
         Top             =   3000
         Width           =   2295
      End
      Begin VB.TextBox tstock 
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
         Left            =   4560
         TabIndex        =   26
         Top             =   2280
         Width           =   1095
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
         Height          =   495
         Left            =   1920
         TabIndex        =   20
         Top             =   1080
         Width           =   3735
      End
      Begin VB.TextBox thargabarang 
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
         Left            =   1920
         TabIndex        =   19
         Top             =   1680
         Width           =   3735
      End
      Begin VB.ComboBox cbokodebarang 
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
         Left            =   1920
         TabIndex        =   18
         Top             =   480
         Width           =   3735
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
         Height          =   495
         Left            =   1920
         TabIndex        =   17
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label16 
         BackColor       =   &H8000000E&
         Caption         =   "Total :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   3600
         TabIndex        =   28
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label15 
         BackColor       =   &H8000000E&
         Caption         =   "Stock :"
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
         Left            =   3720
         TabIndex        =   25
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label14 
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
         Left            =   240
         TabIndex        =   24
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label7 
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
         Height          =   495
         Left            =   240
         TabIndex        =   23
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000E&
         Caption         =   "Harga Barang :"
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
         Left            =   240
         TabIndex        =   22
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000E&
         Caption         =   "Jumlah Barang :"
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
         Left            =   240
         TabIndex        =   21
         Top             =   2400
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "Data Pelanggan"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      TabIndex        =   11
      Top             =   1560
      Width           =   6135
      Begin VB.TextBox tnamapelanggan 
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
         Left            =   2160
         TabIndex        =   13
         Top             =   1200
         Width           =   3495
      End
      Begin VB.ComboBox cbokodepelanggan 
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
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "Kode Pelanggan :"
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
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         Caption         =   "Nama Pelanggan :"
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
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   1695
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   13920
      Top             =   7800
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      Connect         =   "DSN=bengkel"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "bengkel"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tb_service"
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
   Begin VB.CommandButton Command2 
      Caption         =   "Tambah"
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
      Left            =   10080
      TabIndex        =   10
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
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
      Left            =   8880
      TabIndex        =   9
      Top             =   6120
      Width           =   1215
   End
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
      Height          =   495
      Left            =   8760
      TabIndex        =   8
      Top             =   840
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Keluar"
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
      Left            =   12480
      TabIndex        =   7
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Hapus"
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
      Left            =   11280
      TabIndex        =   6
      Top             =   6120
      Width           =   1215
   End
   Begin VB.TextBox ttotalbayar 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   4
      Top             =   5400
      Width           =   3015
   End
   Begin VB.TextBox tnofaktur 
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
      Left            =   1920
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Data Persediaan Sparepart :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   40
      Top             =   6960
      Width           =   2895
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Data Service: "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   39
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Label Label12 
      BackColor       =   &H8000000E&
      Caption         =   "Total Bayar :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   5
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      Caption         =   "Tanggal :"
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
      Height          =   495
      Left            =   7560
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808000&
      Caption         =   "Nomor Faktur :"
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
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "SERVICE KENDARAAN BENGKEL"
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
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "FormService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbokodekaryawan_Validate(Cancel As Boolean)
Call BukaDatabase
Dim rsx As New ADODB.Recordset
Set rsx = New ADODB.Recordset
rsx.CursorLocation = adUseClient
rsx.Open "select * from tb_karyawan where kode_karyawan='" + Trim(cbokodekaryawan.Text) + "'", dbkoneksi, adOpenDynamic, adLockOptimistic
If Not rsx.EOF Then
tnamakaryawan.Text = rsx!nama_karyawan
End If
Set rsx = Nothing
dbkoneksi.Close
Set dbkoneksi = Nothing
End Sub

Private Sub cbokodepelanggan_Validate(Cancel As Boolean)
Call BukaDatabase
Dim rsx As New ADODB.Recordset
Set rsx = New ADODB.Recordset
rsx.CursorLocation = adUseClient
rsx.Open "select * from tb_pelanggan where kode_pelanggan='" + Trim(cbokodepelanggan.Text) + "'", dbkoneksi, adOpenDynamic, adLockOptimistic
If Not rsx.EOF Then
tnamapelanggan.Text = rsx!nama_pelanggan
End If
Set rsx = Nothing
dbkoneksi.Close
Set dbkoneksi = Nothing
End Sub

Private Sub Command1_Click()
tellme = MsgBox("Do you wish to save now?", vbYesNo + vbInformation, "Confirmation")
If tellme = vbNo Then
Exit Sub
End If
Call BukaDatabase
rs1.Open "select * from tb_service where kode_pelanggan is null", dbkoneksi, adOpenDynamic, adLockOptimistic
rs1.AddNew
rs1!nomor_faktur = Trim(tnofaktur.Text)
rs1!tanggal = Format(Trim(ttgl.Text), "dd-mm-yyyy")
rs1!kode_pelanggan = Trim(cbokodepelanggan.Text)
rs1!nama_pelanggan = Trim(tnamapelanggan.Text)
rs1!kode_karyawan = Trim(cbokodekaryawan.Text)
rs1!nama_karyawan = Trim(tnamakaryawan.Text)
rs1!jenis_layanan = Trim(tlayanan.Text)
rs1!biaya_servis = Val(tbiayaservis.Text)
rs1!kode_barang = Trim(cbokodebarang.Text)
rs1!nama_barang = Trim(tnamabarang.Text)
rs1!harga = Val(ttotalharga.Text)
rs1!total_bayar = Val(ttotalharga.Text) + Val(tbiayaservis.Text)
rs1.Update
rs1.Close
Set rs1 = Nothing
rs2.CursorLocation = adUseClient
rs3.Open "select * from tb_jual where kode_customer is null", dbkoneksi, adOpenDynamic, adLockOptimistic
rs3.AddNew
rs3!no_faktur = Trim(tnofaktur.Text)
rs3!tanggal = Format(Trim(ttgl.Text), "dd-mm-yyyy")
rs3!kode_customer = Trim(cbokodepelanggan.Text)
rs3!nama_customer = Trim(tnamapelanggan.Text)
rs3!kode_barang = Trim(cbokodebarang.Text)
rs3!nama_barang = Trim(tnamabarang.Text)
rs3!harga_jual = Val(thargabarang.Text)
rs3!jumlah = Val(tjumlah.Text)
rs3!total_bayar = Val(ttotalharga.Text) * Val(tjumlah.Text)
rs3.Update
rs3.Close
Set rs3 = Nothing
rs2.CursorLocation = adUseClient
rs2.Open "select * from tb_persediaan where kode_barang='" + Trim(cbokodebarang.Text) + "'", dbkoneksi, adOpenDynamic, adLockOptimistic
If rs2.EOF Then
rs2.AddNew
End If
rs2!kode_barang = Trim(cbokodebarang.Text)
rs2!nama_barang = Trim(tnamabarang.Text)
rs2!harga_jual = Val(thargabarang.Text)
rs2!jumlah = rs2!jumlah - Val(tjumlah.Text)
rs2.Update
rs2.Close
Set rs2 = Nothing
dbkoneksi.Close
Set dbkoneksi = Nothing
Adodc1.Refresh
Adodc2.Refresh

End Sub

Private Sub Command2_Click()
Cleartextbox FormService
tnofaktur.SetFocus
End Sub

Private Sub Command3_Click()
Cleartextbox FormService
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub tnofaktur_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
ttgl.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub ttgl_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
cbokodepelanggan.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub cbokodepelanggan_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
tnamapelanggan.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub tnamapelanggan_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
cbokodebarang.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub cbokodebarang_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
tnamabarang.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub tnamabarang_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
tjumlah.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub tjumlah_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
cbokodekaryawan.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub cbokodekaryawan_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
tnamakaryawan.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub tnamakaryawan_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
thargabarang.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub thargabarang_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
tongkoskaryawan.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub thargabarang_LostFocus()
thargabarang.Text = Format(thargabarang.Text, "#########")
End Sub
Private Sub tongkoskaryawan_LostFocus()
ttotal.Text = Format(Str(Val(tongkoskaryawan.Text) + Val(normalize(thargabarang.Text))), "############")
End Sub
Private Sub ttotal_LostFocus()
ttotal.Text = Format(ttotal.Text, "#########")
End Sub
Private Sub tongkoskaryawan_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
ttotal.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cbokodebarang_Validate(Cancel As Boolean)
Call BukaDatabase
Dim rs5 As New ADODB.Recordset
rs5.CursorLocation = adUseClient
rs5.Open "select * from tb_persediaan where kode_barang='" + Trim(cbokodebarang.Text) + "'", dbkoneksi, adOpenDynamic
If Not rs5.EOF Then
tnamabarang.Text = rs5!nama_barang
thargabarang = rs5!harga_jual
tstock.Text = rs5!jumlah
End If
Set rs5 = Nothing
Set dbkoneksi = Nothing

End Sub

Private Sub Form_Load()
WindowState = 2
ttgl.Text = Date$

cbokodepelanggan.AddItem "C001"
cbokodepelanggan.AddItem "C002"
cbokodepelanggan.AddItem "C003"
cbokodepelanggan.AddItem "C004"
cbokodepelanggan.AddItem "C005"
cbokodepelanggan.AddItem "C006"
cbokodepelanggan.AddItem "C007"
cbokodepelanggan.AddItem "C008"
cbokodepelanggan.AddItem "C009"
cbokodepelanggan.AddItem "C010"
cbokodekaryawan.AddItem "K01"
cbokodekaryawan.AddItem "K02"
cbokodekaryawan.AddItem "K03"
cbokodekaryawan.AddItem "K04"
cbokodekaryawan.AddItem "K05"
cbokodekaryawan.AddItem "K06"
cbokodekaryawan.AddItem "K07"
cbokodekaryawan.AddItem "K08"
cbokodekaryawan.AddItem "K09"
cbokodekaryawan.AddItem "K10"
cbokodebarang.AddItem "B01"
cbokodebarang.AddItem "B02"
cbokodebarang.AddItem "B03"
cbokodebarang.AddItem "B04"
cbokodebarang.AddItem "B05"
cbokodebarang.AddItem "B06"
cbokodebarang.AddItem "B07"
cbokodebarang.AddItem "B08"
cbokodebarang.AddItem "B09"
cbokodebarang.AddItem "B10"

ttotalharga.Text = "0"
End Sub


Private Sub ttotalbayar_Click()
ttotalbayar.Text = Val(ttotalharga.Text) + Val(tbiayaservis.Text)
End Sub

Private Sub ttotalharga_Click()
ttotalharga.Text = Val(tjumlah.Text) * Val(thargabarang.Text)
End Sub
