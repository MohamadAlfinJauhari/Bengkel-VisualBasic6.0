VERSION 5.00
Begin VB.Form FormLaporanDataService 
   Caption         =   "Form1"
   ClientHeight    =   4110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   8745
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox CrystalReport1 
      Height          =   480
      Left            =   5760
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   7
      Top             =   2280
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   1560
      TabIndex        =   1
      Top             =   960
      Width           =   4695
      Begin VB.CommandButton Command3 
         Caption         =   "EXIT"
         Height          =   255
         Left            =   3240
         TabIndex        =   6
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "PREVIEW"
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "CETAK LAPORAN"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   960
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "DATE"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Label Label1 
      Caption         =   "LAPORAN DATA SERVICE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   4695
   End
End
Attribute VB_Name = "FormLaporanDataService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
