VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormPenjualanSparepart 
   BackColor       =   &H00808000&
   Caption         =   "INPUT PENJUALAN SPAREPART"
   ClientHeight    =   8910
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19065
   LinkTopic       =   "Form2"
   ScaleHeight     =   8910
   ScaleWidth      =   19065
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H00808000&
      Height          =   1815
      Left            =   10800
      TabIndex        =   40
      Top             =   1080
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
         Left            =   1560
         TabIndex        =   42
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
         TabIndex        =   41
         Top             =   1560
         Width           =   3135
      End
      Begin VB.Image Image1 
         Height          =   1305
         Left            =   120
         Picture         =   "FormPenjualanSparePart.frx":0000
         Top             =   240
         Width           =   1305
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   735
      Left            =   9960
      Top             =   3600
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Bindings        =   "FormPenjualanSparePart.frx":15C7
      Height          =   2055
      Left            =   7920
      TabIndex        =   37
      Top             =   7680
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   3625
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FormPenjualanSparePart.frx":15DC
      Height          =   2055
      Left            =   360
      TabIndex        =   36
      Top             =   7680
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   3625
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   8040
      Top             =   3600
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
      RecordSource    =   "tb_jual"
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
      Left            =   1680
      TabIndex        =   35
      Top             =   6240
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   360
      TabIndex        =   34
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox tstock 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   405
      Left            =   6000
      TabIndex        =   33
      Top             =   5760
      Width           =   1095
   End
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
      Left            =   10440
      TabIndex        =   31
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
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
      Left            =   4080
      TabIndex        =   26
      Top             =   6240
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
      Left            =   2880
      TabIndex        =   25
      Top             =   6240
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
      Left            =   240
      TabIndex        =   14
      Top             =   6240
      Width           =   1455
   End
   Begin VB.TextBox tkembali 
      BackColor       =   &H00C0C000&
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
      Left            =   11760
      TabIndex        =   13
      Top             =   6360
      Width           =   2175
   End
   Begin VB.TextBox tbayar 
      BackColor       =   &H0080C0FF&
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
      Left            =   11760
      TabIndex        =   12
      Top             =   5760
      Width           =   2175
   End
   Begin VB.TextBox ttotalbayar 
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
      Left            =   11760
      TabIndex        =   11
      Top             =   5160
      Width           =   2175
   End
   Begin VB.TextBox tunitjual 
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
      Left            =   8640
      TabIndex        =   10
      Top             =   5160
      Width           =   1575
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
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   5160
      Width           =   2415
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
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   5160
      Width           =   3855
   End
   Begin VB.TextBox ttanggal 
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
      Left            =   5760
      TabIndex        =   7
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "Identitas Customer"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   7575
      Begin VB.TextBox tteleponpelanggan 
         BackColor       =   &H8000000E&
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
         Left            =   2400
         TabIndex        =   29
         Top             =   1920
         Width           =   3255
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H8000000E&
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
         Left            =   2400
         TabIndex        =   27
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox talamatpelanggan 
         BackColor       =   &H8000000E&
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
         Left            =   2400
         TabIndex        =   3
         Top             =   1440
         Width           =   4455
      End
      Begin VB.TextBox tnamapelanggan 
         BackColor       =   &H8000000E&
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
         Left            =   2400
         TabIndex        =   2
         Top             =   960
         Width           =   4455
      End
      Begin VB.Label Label16 
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
         Index           =   0
         Left            =   360
         TabIndex        =   28
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000E&
         Caption         =   "Alamat Customer :"
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
         Left            =   360
         TabIndex        =   6
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         Caption         =   "Nama Customer :"
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
         Left            =   360
         TabIndex        =   5
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "Kode Customer :"
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
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   1695
      End
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
      Height          =   405
      Left            =   2040
      TabIndex        =   0
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label15 
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
      Left            =   7920
      TabIndex        =   39
      Top             =   7200
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Data Penjualan Sparepart :"
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
      Left            =   360
      TabIndex        =   38
      Top             =   7200
      Width           =   2895
   End
   Begin VB.Label Label19 
      BackColor       =   &H00808000&
      Caption         =   "Stock Barang"
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
      Left            =   4560
      TabIndex        =   32
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Label Label18 
      BackColor       =   &H00808000&
      Caption         =   "Satuan"
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
      Left            =   10680
      TabIndex        =   30
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label14 
      BackColor       =   &H00808000&
      Caption         =   "Kembali"
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
      Left            =   10800
      TabIndex        =   24
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label Label13 
      BackColor       =   &H00808000&
      Caption         =   "Total Bayar"
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
      Left            =   12120
      TabIndex        =   23
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackColor       =   &H00808000&
      Caption         =   "Bayar"
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
      Left            =   10800
      TabIndex        =   22
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label Label11 
      BackColor       =   &H00808000&
      Caption         =   "Unit Jual"
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
      Index           =   0
      Left            =   9000
      TabIndex        =   21
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label10 
      BackColor       =   &H00808000&
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
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   6600
      TabIndex        =   20
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackColor       =   &H00808000&
      Caption         =   "Nama Barang"
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
      Left            =   3240
      TabIndex        =   19
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackColor       =   &H00808000&
      Caption         =   "Kode Barang"
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
      Left            =   360
      TabIndex        =   18
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      X1              =   120
      X2              =   14280
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label Label7 
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
      Height          =   375
      Left            =   4800
      TabIndex        =   17
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label3 
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
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "INPUT PENJUALAN SPAREPART"
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
      TabIndex        =   15
      Top             =   240
      Width           =   6495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000E&
      X1              =   120
      X2              =   14280
      Y1              =   960
      Y2              =   960
   End
End
Attribute VB_Name = "FormPenjualanSparepart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Validate(Cancel As Boolean)
Call BukaDatabase
Dim rsx As New ADODB.Recordset
Set rsx = New ADODB.Recordset
rsx.CursorLocation = adUseClient
rsx.Open "select * from tb_pelanggan where kode_pelanggan='" + Trim(Combo1.Text) + "'", dbkoneksi, adOpenDynamic, adLockOptimistic
If Not rsx.EOF Then
tnamapelanggan.Text = rsx!nama_pelanggan
talamatpelanggan.Text = rsx!alamat_pelanggan
tteleponpelanggan.Text = rsx!telepon
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
rs1.Open "select * from tb_jual where kode_customer is null", dbkoneksi, adOpenDynamic, adLockOptimistic
rs1.AddNew
rs1!no_faktur = Trim(tnofaktur.Text)
rs1!tanggal = Format(Trim(ttanggal.Text), "dd-mm-yyyy")
rs1!kode_customer = Trim(Combo1.Text)
rs1!nama_customer = Trim(tnamapelanggan.Text)
rs1!alamat_customer = Trim(talamatpelanggan.Text)
rs1!kode_barang = Trim(Combo2.Text)
rs1!nama_barang = Trim(tnamabarang.Text)
rs1!satuan = Trim(tsatuan.Text)
rs1!harga_jual = Val(thargajual.Text)
rs1!jumlah = Val(tunitjual.Text)
rs1!total_bayar = Val(thargajual.Text) * Val(tunitjual.Text)
rs1!bayar = Val(tbayar.Text)
rs1!kembali = Val(tbayar.Text) - Val(ttotalbayar.Text)
rs1.Update
rs1.Close
Set rs1 = Nothing
rs2.CursorLocation = adUseClient
rs2.Open "select * from tb_persediaan where kode_barang='" + Trim(Combo2.Text) + "'", dbkoneksi, adOpenDynamic, adLockOptimistic
If rs2.EOF Then
rs2.AddNew
End If
rs2!kode_barang = Trim(Combo2.Text)
rs2!nama_barang = Trim(tnamabarang.Text)
rs2!satuan = Trim(tsatuan.Text)
rs2!harga_jual = Val(thargajual.Text)
rs2!jumlah = rs2!jumlah - Val(tunitjual.Text)
rs2.Update
rs2.Close
Set rs2 = Nothing
dbkoneksi.Close
Set dbkoneksi = Nothing
Adodc1.Refresh
Adodc2.Refresh

End Sub

Private Sub Command2_Click()
Cleartextbox FormPenjualanSparepart
tnofaktur.SetFocus
End Sub

Private Sub Command3_Click()
Cleartextbox FormPenjualanSparepart
End Sub
Private Sub Command5_Click()
Unload Me
End Sub

Private Sub tnofaktur_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
ttanggal.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub ttanggal_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
Combo1.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
tnamapelanggan.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub tnamapelanggan_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
talamatpelanggan.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub talamatpelanggan_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
tteleponpelanggan.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub tteleponpelanggan_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
Combo2.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
tnamabarang.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub tnamabarang_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
thargajual.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub thargajual_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
tunitjual.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub tunitjual_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
tsatuan.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub tsatuan_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
ttotalbayar.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub ttotalbayar_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
tbayar.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub tbayar_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
tkembali.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub thargajual_LostFocus()
thargajual.Text = Format(thargajual.Text, "#########")
End Sub
Private Sub ttotalbayar_LostFocus()
ttotalbayar.Text = Format(Str(Val(thargajual.Text) * Val(normalize(tunitjual.Text))), "############")
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
Call BukaDatabase
Dim rs5 As New ADODB.Recordset
rs5.CursorLocation = adUseClient
rs5.Open "select nama_barang,harga_jual,jumlah,satuan from tb_persediaan where kode_barang = '" + Combo2.Text + "'", dbkoneksi, adOpenDynamic, adLockOptimistic
If Not rs5.EOF Then
tnamabarang.Text = rs5!nama_barang
thargajual.Text = rs5!harga_jual
tstock.Text = rs5!jumlah
tsatuan.Text = rs5!satuan
End If
Set rs5 = Nothing
Set dbkoneksi = Nothing
End Sub
Private Sub Form_Load()
WindowState = 2
Call BukaDatabase
Dim rs5 As New ADODB.Recordset
rs5.CursorLocation = adUseClient
rs5.Open "select kode_barang from tb_persediaan order by kode_barang lock in share mode", dbkoneksi, adOpenDynamic, adLockOptimistic
Combo2.Clear
If Not rs5.EOF Then
rs5.MoveFirst
Do While Not rs5.EOF
Combo2.AddItem rs5!kode_barang
rs5.MoveNext
Loop
End If
Set rs5 = Nothing
Set dbkoneksi = Nothing

Combo1.AddItem "C001"
Combo1.AddItem "C002"
Combo1.AddItem "C003"
Combo1.AddItem "C004"
Combo1.AddItem "C005"
Combo1.AddItem "C006"
Combo1.AddItem "C007"
Combo1.AddItem "C008"
Combo1.AddItem "C009"
Combo1.AddItem "C010"
Combo2.AddItem "B01"
Combo2.AddItem "B02"
Combo2.AddItem "B03"
Combo2.AddItem "B04"
Combo2.AddItem "B05"
Combo2.AddItem "B06"
Combo2.AddItem "B07"
Combo2.AddItem "B08"
Combo2.AddItem "B09"
Combo2.AddItem "B10"
ttanggal.Text = Date$
End Sub
Private Sub tbayar_Lostfocus()
tkembali.Text = tbayar.Text - ttotalbayar.Text
End Sub
Private Sub tunitjual_Lostfocus()
Call BukaDatabase
If dbkoneksi.State < 1 Then
MsgBox "Can't get connected to the server!"
Exit Sub
End If
Dim rs3 As ADODB.Recordset
    If Val(tunitjual.Text) > Val(tstock.Text) Then
        MsgBox "Maaf, Jumlah Stok Tidak Terpenuhi:edit", vbOKOnly
        tunitjual.SetFocus
    Else
        FormPenjualanSparepart.Combo1.SetFocus
    End If
    ttotalbayar.Text = thargajual.Text * tunitjual.Text
End Sub


