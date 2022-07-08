VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormInputPelanggan 
   BackColor       =   &H00808000&
   Caption         =   "INPUT DATA PELANGGAN"
   ClientHeight    =   8790
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15525
   LinkTopic       =   "Form2"
   ScaleHeight     =   8790
   ScaleWidth      =   15525
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H00808000&
      Height          =   1815
      Left            =   480
      TabIndex        =   15
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
         Picture         =   "FormInputPelanggan.frx":0000
         Top             =   240
         Width           =   1305
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
      Height          =   3375
      Left            =   4560
      TabIndex        =   6
      Top             =   1200
      Width           =   6375
      Begin VB.TextBox tkode 
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
         Left            =   3120
         TabIndex        =   10
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox tnama 
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
         Left            =   3120
         TabIndex        =   9
         Top             =   1200
         Width           =   2295
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
         Height          =   495
         Left            =   3120
         TabIndex        =   8
         Top             =   1920
         Width           =   2295
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
         Height          =   495
         Left            =   3120
         TabIndex        =   7
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label Label2 
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
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label3 
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
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "Alamat Pelanggan :"
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
         TabIndex        =   12
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label5 
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
         Left            =   360
         TabIndex        =   11
         Top             =   2640
         Width           =   2175
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FormInputPelanggan.frx":15C7
      Height          =   2535
      Left            =   4560
      TabIndex        =   5
      Top             =   4800
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4471
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
      Top             =   8040
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      RecordSource    =   "tb_pelanggan"
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
   Begin VB.CommandButton Command6 
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
      Left            =   6120
      TabIndex        =   4
      Top             =   7560
      Width           =   1575
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
      Left            =   9240
      TabIndex        =   3
      Top             =   7560
      Width           =   1455
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
      Left            =   7680
      TabIndex        =   2
      Top             =   7560
      Width           =   1575
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
      Left            =   4680
      TabIndex        =   1
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "INPUT DATA PELANGGAN"
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
      Left            =   5160
      TabIndex        =   0
      Top             =   480
      Width           =   5295
   End
End
Attribute VB_Name = "FormInputPelanggan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
Dim remove As String
remove = Replace(Replace(ttelp.Text, "(", "", 1, -1, vbTextCompare), ")-", "", 1, -1, vbTextCompare)
ttelp.Text = remove
Call BukaDatabase
rs1.Open "select * from tb_pelanggan where kode_pelanggan is null", dbkoneksi, adOpenDynamic, adLockOptimistic
rs1.AddNew
rs1!kode_pelanggan = Trim(tkode.Text)
rs1!nama_pelanggan = Trim(tnama.Text)
rs1!alamat_pelanggan = Trim(talamat.Text)
rs1!telepon = Trim(ttelp.Text)
rs1.Update
rs1.Close
Set rs1 = Nothing
Adodc1.Refresh
End Sub
Private Sub Command6_Click()
Cleartextbox FormInputPelanggan
tkode.SetFocus
End Sub
Private Sub Command3_Click()
Cleartextbox FormInputPelanggan
End Sub
Private Sub Command4_Click()
FormInputPelanggan.Show
FormInputPelanggan.Caption = "Find Customer name"
End Sub
Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Form_Load()
WindowState = 2
End Sub

Private Sub ttelp_LostFocus()
ttelp.Text = "(" + Left(Trim(ttelp.Text), 4) + ")-" + Mid(Trim(ttelp.Text), 5, 18)
Command1.SetFocus
End Sub
Private Sub tkode_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
tnama.SetFocus
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub tnama_KeyPress(KeyAscii As Integer)
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

