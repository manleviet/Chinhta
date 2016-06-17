VERSION 5.00
Object = "{2B7EDE20-5160-11D1-943D-444553540000}#1.0#0"; "CTBUTTON.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Chuong Trinh Kiem Tra Chinh Ta"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   3855
      Left            =   0
      Picture         =   "frmMain.frx":0CCA
      ScaleHeight     =   3795
      ScaleWidth      =   2235
      TabIndex        =   4
      Top             =   120
      Width           =   2295
   End
   Begin CTBUTTONLibCtl.ctButton cmdThoat 
      Height          =   975
      Left            =   2400
      TabIndex        =   3
      Top             =   3000
      Width           =   2295
      _Version        =   65536
      _ExtentX        =   4048
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   "Thoaït"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VNtimes new roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Thoaït"
      Picture         =   "frmMain.frx":24ACD
      PictureDisabled =   "frmMain.frx":257A7
      PictureDown     =   "frmMain.frx":257C3
      PictureOver     =   "frmMain.frx":257DF
      BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VNtimes new roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderType      =   1
      PicPosition     =   1
   End
   Begin CTBUTTONLibCtl.ctButton cmdTest 
      Height          =   975
      Left            =   2400
      TabIndex        =   2
      Top             =   2040
      Width           =   2295
      _Version        =   65536
      _ExtentX        =   4048
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   "Kiãøm tra chênh taí"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VNtimes new roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Kiãøm tra chênh taí"
      Picture         =   "frmMain.frx":257FB
      PictureDisabled =   "frmMain.frx":264D5
      PictureDown     =   "frmMain.frx":264F1
      PictureOver     =   "frmMain.frx":2650D
      BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VNtimes new roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderType      =   1
      PicPosition     =   1
   End
   Begin CTBUTTONLibCtl.ctButton cmdHoc 
      Height          =   975
      Left            =   2400
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
      _Version        =   65536
      _ExtentX        =   4048
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   "Hoüc chênh taí"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VNtimes new roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Hoüc chênh taí"
      Picture         =   "frmMain.frx":26529
      PictureDisabled =   "frmMain.frx":27203
      PictureDown     =   "frmMain.frx":2721F
      PictureOver     =   "frmMain.frx":2723B
      BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VNtimes new roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderType      =   1
      PicPosition     =   1
   End
   Begin CTBUTTONLibCtl.ctButton cmdChinhTa 
      Height          =   975
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      _Version        =   65536
      _ExtentX        =   4048
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   "Taûo tæì âiãøn chênh taí  "
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VNtimes new roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Taûo tæì âiãøn chênh taí  "
      Picture         =   "frmMain.frx":27257
      PictureDisabled =   "frmMain.frx":27F31
      PictureDown     =   "frmMain.frx":27F4D
      PictureOver     =   "frmMain.frx":27F69
      BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VNtimes new roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderType      =   1
      PicPosition     =   1
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Form nay chi lam moi nhiem vu ket noi toi 3 form con lai
' trong chuong trinh

Private Sub cmdChinhTa_Click()
    frmChinhTa.Show
    Me.Hide
End Sub

Private Sub cmdHoc_Click()
    frmHoc.Show
    Me.Hide
End Sub

Private Sub cmdTest_Click()
    frmTest.Show
    Me.Hide
End Sub

Private Sub cmdThoat_Click()
    End
End Sub
