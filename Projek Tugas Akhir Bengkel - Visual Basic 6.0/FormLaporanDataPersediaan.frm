VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FormLaporanDataPersediaan 
   BackColor       =   &H00808000&
   Caption         =   "LAPORAN DATA PERSEDIAAN"
   ClientHeight    =   4065
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   10350
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Height          =   1335
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   6015
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
         Left            =   1200
         MaskColor       =   &H00FFC0C0&
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "EXIT"
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
         Left            =   3600
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   5640
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "C:\Users\ACER AspireOne Z1402\Documents\PEMROGRAMAN JARINGAN\LAPORAN BENGKEL\LAPORAN DATA PERSEDIAAN.rpt"
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "LAPORAN DATA PERSEDIAAN"
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
      Left            =   600
      TabIndex        =   3
      Top             =   360
      Width           =   6015
   End
End
Attribute VB_Name = "FormLaporanDataPersediaan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.CrystalReport1.WindowState = crptMaximized
Me.CrystalReport1.RetrieveDataFiles
Me.CrystalReport1.Action = 1
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

