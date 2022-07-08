VERSION 5.00
Begin VB.Form FormLoginAdmin 
   BackColor       =   &H00808000&
   Caption         =   "LOGIN"
   ClientHeight    =   6645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9930
   LinkTopic       =   "Form2"
   ScaleHeight     =   6645
   ScaleWidth      =   9930
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox tpassword 
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
      Left            =   3480
      TabIndex        =   7
      Top             =   3240
      Width           =   4335
   End
   Begin VB.CommandButton Command2 
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
      Left            =   3480
      TabIndex        =   5
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Masuk"
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
      Left            =   2040
      TabIndex        =   4
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox tuser 
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
      Left            =   3480
      TabIndex        =   3
      Top             =   2520
      Width           =   4335
   End
   Begin VB.TextBox tuserid 
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
      Left            =   3480
      TabIndex        =   2
      Top             =   1800
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000E&
      Height          =   3735
      Left            =   840
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   8175
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "SISTEM INFORMASI BENGKEL KOMPAK"
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
      Left            =   960
      TabIndex        =   8
      Top             =   480
      Width           =   8055
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808000&
      Caption         =   "Nama Admin :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      Caption         =   "Kode Admin :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   1800
      Width           =   1455
   End
End
Attribute VB_Name = "FormLoginAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim rs As New ADODB.Recordset

Set rs = JalankanSQL("select * from tb_admin where kode_admin = '" & tuserid & "'")

If rs.RecordCount = 0 Then
    MsgBox "Kode Admin tidak ditemukan!", vbCritical + vbOKOnly, "Perhatian"
    tuserid.SetFocus
    Exit Sub
End If

Set rs = JalankanSQL("select * from tb_admin where kode_admin = '" & tuserid & "' and password_admin = '" & tpassword.Text & "'")

If rs.RecordCount = 0 Then
    MsgBox "Password Salah tidak ditemukan!", vbCritical + vbOKOnly, "Perhatian"
    tpassword.SetFocus
    Exit Sub
Else
    FormMenuUtama.Show
    Me.Hide
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub


Private Sub tpassword_Change()
tpassword.PasswordChar = "*"
End Sub
