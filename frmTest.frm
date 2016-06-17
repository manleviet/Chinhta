VERSION 5.00
Object = "{2B7EDE20-5160-11D1-943D-444553540000}#1.0#0"; "CTBUTTON.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test Chinh Ta"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox txtKiemTraChinhTa 
      Height          =   1815
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3201
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmTest.frx":0CCA
   End
   Begin CTBUTTONLibCtl.ctButton cmdBack 
      Height          =   615
      Left            =   3600
      TabIndex        =   2
      Top             =   2280
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   1085
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
      Picture         =   "frmTest.frx":0D9F
      PictureDisabled =   "frmTest.frx":1A79
      PictureDown     =   "frmTest.frx":1A95
      PictureOver     =   "frmTest.frx":1AB1
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
   Begin CTBUTTONLibCtl.ctButton cmdTest 
      Height          =   615
      Left            =   960
      TabIndex        =   1
      Top             =   2280
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Kiãøm tra"
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
      SoundClick      =   "beep"
      Caption         =   "Kiãøm tra"
      Picture         =   "frmTest.frx":1ACD
      PictureDisabled =   "frmTest.frx":27A7
      PictureDown     =   "frmTest.frx":27C3
      PictureOver     =   "frmTest.frx":27DF
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
   Begin VB.Label Label1 
      Caption         =   "Âoaûn vàn muäún kiãøm tra chênh taí:"
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
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TestString As clsTest

Private Sub cmdBack_Click()
    frmMain.Show
    Me.Hide
End Sub

Private Sub cmdTest_Click()
    Dim i As Long
    Dim Pos As Integer
    Dim lPos As Long
    If txtKiemTraChinhTa.Text <> "" Then
        Set TestString = New clsTest
        TestString.Text = txtKiemTraChinhTa.Text
        lPos = 1
        For i = 1 To TestString.TokenCount
            PhanTich LCase(TestString.TokenAt(i))
            Pos = FindBinary(Words.ToHopNguyenAm)
            lPos = InStr(lPos, txtKiemTraChinhTa.Text, TestString.TokenAt(i))
            If (Pos <> -1) Then
                If ((ArrayWord(Pos).PhuAmDau And Words.PhuAmDau) <> 0) And ((ArrayWord(Pos).PhuAmCuoi And Words.PhuAmCuoi) <> 0) Then
                    'txtDung.Text = txtDung.Text & TestString.TokenAt(i) & vbCrLf
                Else
                    txtKiemTraChinhTa.SelStart = lPos - 1
                    txtKiemTraChinhTa.SelLength = Len(TestString.TokenAt(i))
                    txtKiemTraChinhTa.SelUnderline = True
                    'txtKiemTraChinhTa.SelColor = vbRed
                End If
            Else
                txtKiemTraChinhTa.SelStart = lPos - 1
                txtKiemTraChinhTa.SelLength = Len(TestString.TokenAt(i))
                txtKiemTraChinhTa.SelUnderline = True
                'txtKiemTraChinhTa.SelColor = vbRed
                'txtSai.Text = txtSai.Text & TestString.TokenAt(i) & vbCrLf
            End If
            lPos = lPos + Len(TestString.TokenAt(i)) - 1
        Next i
    Else
        MsgBox "Ban chua nhap gi vao TextBox ca!", vbExclamation, "Error"
        txtKiemTraChinhTa.SetFocus
    End If
End Sub

Private Sub Form_Load()
'    txtDung.Text = ""
'    txtSai.Text = ""
    txtKiemTraChinhTa.Text = ""
    txtKiemTraChinhTa.SelFontName = "VNtimes new roman"
    txtKiemTraChinhTa.SelFontSize = 12
    txtKiemTraChinhTa.SelColor = &HFF0000
    LoadArray
    NewWords
End Sub
