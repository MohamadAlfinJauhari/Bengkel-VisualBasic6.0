VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormKaryawan 
   BackColor       =   &H00808000&
   Caption         =   " INPUT DATA KARYAWAN"
   ClientHeight    =   9390
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15015
   LinkTopic       =   "Form1"
   ScaleHeight     =   9390
   ScaleWidth      =   15015
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H00808000&
      Height          =   1815
      Left            =   240
      TabIndex        =   15
      Top             =   600
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
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   1560
         Width           =   3135
      End
      Begin VB.Image Image1 
         Height          =   1305
         Left            =   120
         Picture         =   "Formkaryawan.frx":0000
         Top             =   240
         Width           =   1305
      End
   End
   Begin VB.CommandButton tbtambah 
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
      Height          =   375
      Left            =   5640
      TabIndex        =   14
      Top             =   8160
      Width           =   1575
   End
   Begin VB.CommandButton tbsimpan 
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
      Height          =   375
      Left            =   4080
      TabIndex        =   13
      Top             =   8160
      Width           =   1575
   End
   Begin VB.CommandButton tbhapus 
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
      Height          =   375
      Left            =   7200
      TabIndex        =   12
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton tbkeluar 
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
      Height          =   375
      Left            =   8640
      TabIndex        =   11
      Top             =   8160
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Formkaryawan.frx":15C7
      Height          =   2655
      Left            =   3960
      TabIndex        =   9
      Top             =   5280
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   4683
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
      Left            =   12000
      Top             =   8640
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
      RecordSource    =   "tb_karyawan"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000E&
      Caption         =   "Data Karyawan"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   3960
      TabIndex        =   0
      Top             =   1680
      Width           =   7575
      Begin VB.TextBox ttelpkaryawan 
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
         Left            =   3000
         TabIndex        =   8
         Top             =   2640
         Width           =   2655
      End
      Begin VB.TextBox talamatkaryawan 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3000
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   1440
         Width           =   4455
      End
      Begin VB.TextBox tnmkaryawan 
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
         Left            =   3000
         TabIndex        =   6
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox tkdkaryawan 
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
         Left            =   3000
         TabIndex        =   5
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         Caption         =   "ALAMAT :"
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
         TabIndex        =   4
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "TELEPON :"
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
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "NAMA KARYAWAN :"
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
         TabIndex        =   2
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "KODE KARYAWAN :"
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
         TabIndex        =   1
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "INPUT DATA KARYAWAN"
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
      Left            =   5280
      TabIndex        =   10
      Top             =   600
      Width           =   5055
   End
End
Attribute VB_Name = "FormKaryawan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
WindowState = 2
End Sub

Private Sub tbsimpan_Click()
Dim remove As String
remove = Replace(Replace(ttelpkaryawan.Text, "(", "", 1, -1, vbTextCompare), ")-", "", 1, -1, vbTextCompare)
ttelpkaryawan.Text = remove
Call BukaDatabase
rs1.Open "select * from tb_karyawan where kode_karyawan is null", dbkoneksi, adOpenDynamic, adLockOptimistic
rs1.AddNew
rs1!kode_karyawan = Trim(tkdkaryawan.Text)
rs1!nama_karyawan = Trim(tnmkaryawan.Text)
rs1!alamat_karyawan = Trim(talamatkaryawan.Text)
rs1!telepon = Trim(ttelpkaryawan.Text)
rs1.Update
rs1.Close
Set rs1 = Nothing
Adodc1.Refresh
End Sub
Private Sub tbtambah_Click()
Cleartextbox FormKaryawan
tkdkaryawan.SetFocus
End Sub
Private Sub tbhapus_Click()
Cleartextbox FormKaryawan
End Sub
Private Sub tbkeluar_Click()
Unload Me
End Sub
Private Sub talamatkaryawan_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
talamatkaryawan.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub ttelpkaryawan_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
tbsimpan.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub ttelpkaryawan_LostFocus()
ttelpkaryawan.Text = "(" + Left(Trim(ttelpkaryawan.Text), 4) + ")-" + Mid(Trim(ttelpkaryawan.Text), 5, 18)
tbsimpan.SetFocus
End Sub
Private Sub tnmkaryawan_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
tbsimpan.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

