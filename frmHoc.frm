VERSION 5.00
Object = "{2B7EDE20-5160-11D1-943D-444553540000}#1.0#0"; "CTBUTTON.OCX"
Begin VB.Form frmHoc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hoc Chinh Ta"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   Icon            =   "frmHoc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHocChinhTa 
      BeginProperty Font 
         Name            =   "VNtimes new roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1815
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   360
      Width           =   6135
   End
   Begin CTBUTTONLibCtl.ctButton cmdBack 
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   2280
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Vãö Main"
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
      Caption         =   "Vãö Main"
      Picture         =   "frmHoc.frx":0CCA
      PictureDisabled =   "frmHoc.frx":19A4
      PictureDown     =   "frmHoc.frx":19C0
      PictureOver     =   "frmHoc.frx":19DC
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
      PicPosition     =   3
      TextAlign       =   1
   End
   Begin CTBUTTONLibCtl.ctButton cmdHoc 
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   2280
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Hoüc   "
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
      Caption         =   "Hoüc   "
      Picture         =   "frmHoc.frx":19F8
      PictureDisabled =   "frmHoc.frx":26D2
      PictureDown     =   "frmHoc.frx":26EE
      PictureOver     =   "frmHoc.frx":270A
      BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VNtimes new roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColorOver   =   16711680
      BorderType      =   1
      PicPosition     =   3
      TextAlign       =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Nháûp âoaûn vàn âuïng chênh taí:"
      BeginProperty Font 
         Name            =   "VNtimes new roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "frmHoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TestString As clsTest

Private Sub cmdBack_Click()
    frmMain.Show
    Me.Hide
    SaveArray
End Sub

Private Sub cmdHoc_Click()
    Dim i As Long
    Dim Pos As Long
    If txtHocChinhTa.Text <> "" Then
        Set TestString = New clsTest
        TestString.Text = txtHocChinhTa.Text
        For i = 1 To TestString.TokenCount
            PhanTich LCase(TestString.TokenAt(i))
            SaveWord
        Next i
    Else
        MsgBox "Ban chua nhap gi vao TextBox ca!", vbExclamation, "Error"
        txtHocChinhTa.SetFocus
    End If
End Sub

Private Sub Form_Load()
    LoadArray
    NewWords
End Sub
