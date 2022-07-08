VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FormLaporanDataKaryawan 
   BackColor       =   &H00808000&
   Caption         =   "LAPORAN DATA KARYAWAN"
   ClientHeight    =   3945
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8640
   FillColor       =   &H00808000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   8640
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   5040
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "C:\Users\ACER AspireOne Z1402\Documents\PEMROGRAMAN JARINGAN\LAPORAN BENGKEL\LAPORAN DATA KARYAWAN.rpt"
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Height          =   1215
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   5775
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
         Left            =   3120
         TabIndex        =   3
         Top             =   360
         Width           =   1095
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
         Left            =   960
         MaskColor       =   &H00FFC0C0&
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "LAPORAN DATA KARYAWAN"
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
      TabIndex        =   1
      Top             =   480
      Width           =   5775
   End
End
Attribute VB_Name = "FormLaporanDataKaryawan"
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
