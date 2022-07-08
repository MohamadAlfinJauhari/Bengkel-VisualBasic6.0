VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormLaporanTransaksiPenjualan 
   BackColor       =   &H00808000&
   Caption         =   "LAPORAN TRANSAKSI PENJUALAN"
   ClientHeight    =   5070
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Height          =   3855
      Left            =   840
      TabIndex        =   0
      Top             =   960
      Width           =   7335
      Begin VB.CommandButton Command3 
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
         Left            =   3840
         TabIndex        =   10
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000D&
         Caption         =   "CETAK"
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
         Left            =   5280
         MaskColor       =   &H00FFC0C0&
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "CETAK"
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
         Left            =   5280
         TabIndex        =   1
         Top             =   3120
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   4440
         TabIndex        =   2
         Top             =   2400
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   134283265
         CurrentDate     =   44745
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   4440
         TabIndex        =   3
         Top             =   1680
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   134283265
         CurrentDate     =   44745
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808000&
         Caption         =   "SEMUA TRANSAKSI PENJUALAN :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808000&
         Caption         =   "CETAK PER TANGGAL :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   360
         TabIndex        =   7
         Top             =   1080
         Width           =   3615
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "DARI TANGGAL :"
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
         Left            =   360
         TabIndex        =   6
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         Caption         =   "SAMPAI TANGGAL :"
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
         Left            =   360
         TabIndex        =   5
         Top             =   2400
         Width           =   2415
      End
   End
   Begin Crystal.CrystalReport CrystalReport2 
      Left            =   8280
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "C:\Users\ACER AspireOne Z1402\Documents\PEMROGRAMAN JARINGAN\LAPORAN BENGKEL\LAPORAN TRANSAKSI PENJUALAN.rpt"
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   8280
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "C:\Users\ACER AspireOne Z1402\Documents\PEMROGRAMAN JARINGAN\LAPORAN BENGKEL\LAPORAN TRANSAKSI PENJUALAN.rpt"
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "LAPORAN TRANSAKSI PENJUALAN"
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
      Left            =   840
      TabIndex        =   9
      Top             =   360
      Width           =   7335
   End
End
Attribute VB_Name = "FormLaporanTransaksiPenjualan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.CrystalReport1.WindowState = crptMaximized
Me.CrystalReport1.RetrieveDataFiles
Me.CrystalReport1.Action = 1
End Sub

Private Sub Command2_Click()
Dim tanggal As String
tanggal = "{tb_jual.tanggal}in date" + "('" & Format(DTPicker1, "mm-dd-yyyy") & "')to date" + "('" & Format(DTPicker2, "mm-dd-yyyy") & "')"
Me.CrystalReport2.SelectionFormula = tanggal
Me.CrystalReport2.WindowState = crptMaximized
Me.CrystalReport2.RetrieveDataFiles
Me.CrystalReport2.Action = 1
End Sub

Private Sub Command3_Click()
Unload Me
End Sub
