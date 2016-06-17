VERSION 5.00
Object = "{2B7EDE20-5160-11D1-943D-444553540000}#1.0#0"; "CTBUTTON.OCX"
Begin VB.Form frmChinhTa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tao tu dien chinh ta"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
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
   Icon            =   "frmChinhTa.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtKiemTra 
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   48
      Top             =   5400
      Width           =   4095
   End
   Begin VB.TextBox txtTu 
      ForeColor       =   &H00FF0000&
      Height          =   3615
      Left            =   2160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   40
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Frame fraPhuAmCuoi 
      Caption         =   "Phuû ám cuäúi:"
      ForeColor       =   &H00FF0000&
      Height          =   2055
      Left            =   4200
      TabIndex        =   30
      Top             =   0
      Width           =   1815
      Begin VB.CheckBox chkPhuAmCuoi 
         Caption         =   "None"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   8
         Left            =   600
         TabIndex        =   39
         Top             =   1680
         Width           =   855
      End
      Begin VB.CheckBox chkPhuAmCuoi 
         Caption         =   "t"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   7
         Left            =   1080
         TabIndex        =   38
         Top             =   1350
         Width           =   495
      End
      Begin VB.CheckBox chkPhuAmCuoi 
         Caption         =   "p"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   6
         Left            =   1080
         TabIndex        =   37
         Top             =   1020
         Width           =   495
      End
      Begin VB.CheckBox chkPhuAmCuoi 
         Caption         =   "nh"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   5
         Left            =   1080
         TabIndex        =   36
         Top             =   690
         Width           =   615
      End
      Begin VB.CheckBox chkPhuAmCuoi 
         Caption         =   "ng"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   4
         Left            =   1080
         TabIndex        =   35
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox chkPhuAmCuoi 
         Caption         =   "n"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   3
         Left            =   120
         TabIndex        =   34
         Top             =   1350
         Width           =   735
      End
      Begin VB.CheckBox chkPhuAmCuoi 
         Caption         =   "m"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   2
         Left            =   120
         TabIndex        =   33
         Top             =   1020
         Width           =   615
      End
      Begin VB.CheckBox chkPhuAmCuoi 
         Caption         =   "ch"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   1
         Left            =   120
         TabIndex        =   32
         Top             =   690
         Width           =   615
      End
      Begin VB.CheckBox chkPhuAmCuoi 
         Caption         =   "c"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.TextBox txtNguyenAm 
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   2160
      MaxLength       =   5
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.Frame fraPhuAmDau 
      Caption         =   "Phuû ám âáöu:"
      ForeColor       =   &H00FF0000&
      Height          =   5055
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1935
      Begin VB.CheckBox chkPhuAmDau 
         Caption         =   "None"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   26
         Left            =   360
         TabIndex        =   28
         Top             =   4680
         Width           =   855
      End
      Begin VB.CheckBox chkPhuAmDau 
         Caption         =   "x"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   25
         Left            =   1080
         TabIndex        =   27
         Top             =   4320
         Width           =   495
      End
      Begin VB.CheckBox chkPhuAmDau 
         Caption         =   "v"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   24
         Left            =   1080
         TabIndex        =   26
         Top             =   3990
         Width           =   495
      End
      Begin VB.CheckBox chkPhuAmDau 
         Caption         =   "tr"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   23
         Left            =   1080
         TabIndex        =   25
         Top             =   3660
         Width           =   495
      End
      Begin VB.CheckBox chkPhuAmDau 
         Caption         =   "th"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   22
         Left            =   1080
         TabIndex        =   24
         Top             =   3330
         Width           =   495
      End
      Begin VB.CheckBox chkPhuAmDau 
         Caption         =   "t"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   21
         Left            =   1080
         TabIndex        =   23
         Top             =   3000
         Width           =   495
      End
      Begin VB.CheckBox chkPhuAmDau 
         Caption         =   "s"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   20
         Left            =   1080
         TabIndex        =   22
         Top             =   2670
         Width           =   495
      End
      Begin VB.CheckBox chkPhuAmDau 
         Caption         =   "r"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   19
         Left            =   1080
         TabIndex        =   21
         Top             =   2340
         Width           =   495
      End
      Begin VB.CheckBox chkPhuAmDau 
         Caption         =   "l"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   10
         Left            =   120
         TabIndex        =   20
         Top             =   3660
         Width           =   495
      End
      Begin VB.CheckBox chkPhuAmDau 
         Caption         =   "n"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   12
         Left            =   120
         TabIndex        =   19
         Top             =   4320
         Width           =   495
      End
      Begin VB.CheckBox chkPhuAmDau 
         Caption         =   "q"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   18
         Left            =   1080
         TabIndex        =   18
         Top             =   2010
         Width           =   495
      End
      Begin VB.CheckBox chkPhuAmDau 
         Caption         =   "ph"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   17
         Left            =   1080
         TabIndex        =   17
         Top             =   1680
         Width           =   615
      End
      Begin VB.CheckBox chkPhuAmDau 
         Caption         =   "p"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   16
         Left            =   1080
         TabIndex        =   16
         Top             =   1350
         Width           =   495
      End
      Begin VB.CheckBox chkPhuAmDau 
         Caption         =   "nh"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   15
         Left            =   1080
         TabIndex        =   15
         Top             =   1020
         Width           =   615
      End
      Begin VB.CheckBox chkPhuAmDau 
         Caption         =   "ngh"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   14
         Left            =   1080
         TabIndex        =   14
         Top             =   690
         Width           =   735
      End
      Begin VB.CheckBox chkPhuAmDau 
         Caption         =   "ng"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   13
         Left            =   1080
         TabIndex        =   13
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox chkPhuAmDau 
         Caption         =   "m"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   11
         Left            =   120
         TabIndex        =   12
         Top             =   3990
         Width           =   495
      End
      Begin VB.CheckBox chkPhuAmDau 
         Caption         =   "kh"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   9
         Left            =   120
         TabIndex        =   11
         Top             =   3330
         Width           =   615
      End
      Begin VB.CheckBox chkPhuAmDau 
         Caption         =   "k"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   8
         Left            =   120
         TabIndex        =   10
         Top             =   3000
         Width           =   495
      End
      Begin VB.CheckBox chkPhuAmDau 
         Caption         =   "h"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   7
         Left            =   120
         TabIndex        =   9
         Top             =   2670
         Width           =   495
      End
      Begin VB.CheckBox chkPhuAmDau 
         Caption         =   "gh"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   6
         Left            =   120
         TabIndex        =   8
         Top             =   2340
         Width           =   615
      End
      Begin VB.CheckBox chkPhuAmDau 
         Caption         =   "g"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   5
         Left            =   120
         TabIndex        =   7
         Top             =   2010
         Width           =   495
      End
      Begin VB.CheckBox chkPhuAmDau 
         Caption         =   "â"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   495
      End
      Begin VB.CheckBox chkPhuAmDau 
         Caption         =   "d"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   1350
         Width           =   495
      End
      Begin VB.CheckBox chkPhuAmDau 
         Caption         =   "ch"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   1020
         Width           =   615
      End
      Begin VB.CheckBox chkPhuAmDau 
         Caption         =   "c"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   690
         Width           =   495
      End
      Begin VB.CheckBox chkPhuAmDau 
         Caption         =   "b"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   495
      End
   End
   Begin CTBUTTONLibCtl.ctButton ctButton1 
      Height          =   495
      Left            =   4320
      TabIndex        =   50
      Top             =   5640
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Kiãøm Tra"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VNtimes new roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Kiãøm Tra"
      Picture         =   "frmChinhTa.frx":0CCA
      PictureDisabled =   "frmChinhTa.frx":0CE6
      PictureDown     =   "frmChinhTa.frx":0D02
      PictureOver     =   "frmChinhTa.frx":0D1E
      BorderType      =   1
   End
   Begin VB.Label Label3 
      Caption         =   "Haîy nháún vaìo nuït bãn âãø xem mçnh âaî taûo nhæîng tæì naìo:"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   49
      Top             =   5040
      Width           =   5655
   End
   Begin CTBUTTONLibCtl.ctButton cmdTaoLai 
      Height          =   495
      Left            =   4320
      TabIndex        =   47
      Top             =   3120
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Taûo Laûi  "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VNtimes new roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Taûo Laûi  "
      Picture         =   "frmChinhTa.frx":0D3A
      PictureDisabled =   "frmChinhTa.frx":1A14
      PictureDown     =   "frmChinhTa.frx":1A30
      PictureOver     =   "frmChinhTa.frx":1A4C
      BorderType      =   1
      PicPosition     =   3
      TextAlign       =   1
   End
   Begin CTBUTTONLibCtl.ctButton cmdNhapMoi 
      Height          =   495
      Left            =   4320
      TabIndex        =   46
      Top             =   3120
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Nháûp Måïi"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VNtimes new roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Nháûp Måïi"
      Picture         =   "frmChinhTa.frx":1A68
      PictureDisabled =   "frmChinhTa.frx":2742
      PictureDown     =   "frmChinhTa.frx":275E
      PictureOver     =   "frmChinhTa.frx":277A
      BorderType      =   1
      PicPosition     =   3
      TextAlign       =   1
   End
   Begin CTBUTTONLibCtl.ctButton cmdGhiNhan 
      Height          =   495
      Left            =   4320
      TabIndex        =   45
      Top             =   2400
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Ghi Nháûn"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VNtimes new roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ghi Nháûn"
      Picture         =   "frmChinhTa.frx":2796
      PictureDisabled =   "frmChinhTa.frx":3470
      PictureDown     =   "frmChinhTa.frx":348C
      PictureOver     =   "frmChinhTa.frx":34A8
      BorderType      =   1
      PicPosition     =   3
      TextAlign       =   1
   End
   Begin CTBUTTONLibCtl.ctButton cmdHienThi 
      Height          =   495
      Left            =   4320
      TabIndex        =   44
      Top             =   2400
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Hiãøn thë "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VNtimes new roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Hiãøn thë "
      Picture         =   "frmChinhTa.frx":34C4
      PictureDisabled =   "frmChinhTa.frx":419E
      PictureDown     =   "frmChinhTa.frx":41BA
      PictureOver     =   "frmChinhTa.frx":41D6
      BorderType      =   1
      PicPosition     =   3
      TextAlign       =   1
   End
   Begin CTBUTTONLibCtl.ctButton cmdBack 
      Height          =   495
      Left            =   4320
      TabIndex        =   43
      Top             =   4560
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Vãö Main"
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
      Picture         =   "frmChinhTa.frx":41F2
      PictureDisabled =   "frmChinhTa.frx":4ECC
      PictureDown     =   "frmChinhTa.frx":4EE8
      PictureOver     =   "frmChinhTa.frx":4F04
      BorderType      =   1
      PicPosition     =   3
      TextAlign       =   1
   End
   Begin CTBUTTONLibCtl.ctButton cmdSave 
      Height          =   495
      Left            =   4320
      TabIndex        =   42
      Top             =   3840
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Læu File "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VNtimes new roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Læu File "
      Picture         =   "frmChinhTa.frx":4F20
      PictureDisabled =   "frmChinhTa.frx":5BFA
      PictureDown     =   "frmChinhTa.frx":5C16
      PictureOver     =   "frmChinhTa.frx":5C32
      BorderType      =   1
      PicPosition     =   3
      TextAlign       =   1
   End
   Begin VB.Label Label2 
      Caption         =   "Nhæîng tæì âaî taûo:"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2160
      TabIndex        =   41
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Täø håüp nguyãn ám:"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2160
      TabIndex        =   29
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "frmChinhTa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GhiNhanOK As Boolean

Private Sub cmdBack_Click()
    frmMain.Show
    Me.Hide
End Sub

Private Sub cmdThoat_Click()
    End
End Sub

Private Sub cmdGhiNhan_Click()
    SaveWord
    NewWords
    GhiNhanOK = True
    ShowFormChinhTa HienThi
End Sub

Private Sub cmdHienThi_Click()
    If txtNguyenAm <> "" Then
        Words.PhuAmDau = DocPhuAmDau
        Words.PhuAmCuoi = DocPhuAmCuoi
        Words.ToHopNguyenAm = txtNguyenAm.Text
        HienThiTu
    Else
        MsgBox "Mot tu khong the khong co nguyen am!!!", vbExclamation, "Error"
        Exit Sub
    End If
    ShowFormChinhTa GhiNhan
End Sub

Private Sub cmdNhapMoi_Click()
    If Not GhiNhanOK And txtNguyenAm <> "" Then
        Dim i As Integer
        i = MsgBox("Ban co ghi nhan am tiet nay hay khong?", vbYesNoCancel, "Information")
        Select Case i
            Case vbYes: cmdGhiNhan_Click
            Case vbCancel: Exit Sub
        End Select
    End If
    ClearForm cmdNhapMoi
    GhiNhanOK = False
End Sub

Private Sub cmdSave_Click()
    SaveArray
End Sub

Private Sub cmdTaoLai_Click()
    ClearForm cmdTaoLai
    ShowFormChinhTa HienThi
End Sub

Private Sub ctButton1_Click()
    Dim i As Integer
    txtKiemTra.Text = ""
    For i = 0 To iArrayWord - 1
        txtKiemTra.Text = txtKiemTra.Text & ArrayWord(i).PhuAmDau & "|" & ArrayWord(i).ToHopNguyenAm & "|" & ArrayWord(i).PhuAmCuoi & vbCrLf
    Next i
End Sub

Private Sub Form_Load()
    LoadArray
    NewWords
    ShowFormChinhTa HienThi
    GhiNhanOK = False
End Sub

