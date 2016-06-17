Attribute VB_Name = "mdlChinhTa"
'Day la module gom cac ham xu ly chung cho form frmChinhTa

Public Sub ShowFormChinhTa(md As Mode)
    ' Xu ly hien thi cho 4 nut cmdGhiNhan, cmdHienThi
    ' cmdNhapMoi, cmdTaoLai
    ' voi Mode la mot enum duoc khai bao trong mdlMain
    With frmChinhTa
        .cmdGhiNhan.Visible = (GhiNhan = md)
        .cmdHienThi.Visible = (HienThi = md)
        .cmdNhapMoi.Visible = (HienThi = md)
        .cmdTaoLai.Visible = (GhiNhan = md)
    End With
End Sub

Public Sub ClearForm(obj As Object)
    ' Ham xu ly xoa chon trong CheckBox va TextBox
    Dim chki As Integer
    With frmChinhTa
        For chki = 0 To 26
            .chkPhuAmDau(chki).Value = False
        Next
        For chki = 0 To 8
            .chkPhuAmCuoi(chki).Value = False
        Next
        ' Tuy vao doi tuong nut bam ma co xoa trong TextBox hay khong
        If obj.Name = "cmdNhapMoi" Then
            .txtNguyenAm.Text = ""
        End If
        .txtTu = ""
    End With
    NewWords
End Sub

Public Function DocPhuAmDau() As Long
    ' Ham tinh gia tri cho truong Words.PhuAmDau
    Dim i As Long
    Dim chki As Integer
    With frmChinhTa
        For chki = 0 To 26
            If .chkPhuAmDau(chki).Value = 1 Then
                i = i + 2 ^ chki
            End If
        Next
    End With
    DocPhuAmDau = i
End Function

Public Function DocPhuAmCuoi() As Integer
    ' Ham tinh gia tri cho truong Words.PhuAmCuoi
    Dim i As Integer
    Dim chki As Integer
    With frmChinhTa
        For chki = 0 To 8
            If .chkPhuAmCuoi(chki).Value = 1 Then
                i = i + 2 ^ chki
            End If
        Next
    End With
    DocPhuAmCuoi = i
End Function

Public Sub HienThiTu()
    Dim str As String
    Dim chuoi As String
    Dim chki As Integer
    Dim chkj As Integer
    With frmChinhTa
        For chki = 0 To 26
            If .chkPhuAmDau(chki).Value = 1 Then
                If chki = 26 Then
                    chuoi = .txtNguyenAm.Text
                Else
                    chuoi = .chkPhuAmDau(chki).Caption & .txtNguyenAm.Text
                End If
                For chkj = 0 To 8
                    If .chkPhuAmCuoi(chkj).Value = 1 Then
                        If chkj = 8 Then
                            str = str & chuoi & vbCrLf
                        Else
                            str = str & chuoi & .chkPhuAmCuoi(chkj).Caption & vbCrLf
                        End If
                    End If
                Next chkj
            End If
            .txtTu.Text = str & vbCrLf
        Next chki
    End With
End Sub

'Public Function HienThiPhuAmDau(i As Integer) As String
'    Select Case i
'        Case 0: HienThiPhuAmDau = "b"
'        Case 1: HienThiPhuAmDau = "c"
'        Case 2: HienThiPhuAmDau = "ch"
'        Case 3: HienThiPhuAmDau = "d"
'        Case 4: HienThiPhuAmDau = "â"
'        Case 5: HienThiPhuAmDau = "g"
'        Case 6: HienThiPhuAmDau = "gh"
'        Case 7: HienThiPhuAmDau = "h"
'        Case 8: HienThiPhuAmDau = "k"
'        Case 9: HienThiPhuAmDau = "kh"
'        Case 10: HienThiPhuAmDau = "l"
'        Case 11: HienThiPhuAmDau = "m"
'        Case 12: HienThiPhuAmDau = "n"
'        Case 13: HienThiPhuAmDau = "ng"
'        Case 14: HienThiPhuAmDau = "ngh"
'        Case 15: HienThiPhuAmDau = "nh"
'        Case 16: HienThiPhuAmDau = "p"
'        Case 17: HienThiPhuAmDau = "ph"
'        Case 18: HienThiPhuAmDau = "q"
'        Case 19: HienThiPhuAmDau = "r"
'        Case 20: HienThiPhuAmDau = "s"
'        Case 21: HienThiPhuAmDau = "t"
'        Case 22: HienThiPhuAmDau = "th"
'        Case 23: HienThiPhuAmDau = "tr"
'        Case 24: HienThiPhuAmDau = "v"
'        Case 25: HienThiPhuAmDau = "x"
'        Case 26: HienThiPhuAmDau = ""
'    End Select
'End Function

'Public Function HienThiPhuAmCuoi(i As Integer) As String
'    Select Case i
'        Case 0: HienThiPhuAmCuoi = "c"
'        Case 1: HienThiPhuAmCuoi = "ch"
'        Case 2: HienThiPhuAmCuoi = "m"
'        Case 3: HienThiPhuAmCuoi = "n"
'        Case 4: HienThiPhuAmCuoi = "ng"
'        Case 5: HienThiPhuAmCuoi = "nh"
'        Case 6: HienThiPhuAmCuoi = "p"
'        Case 7: HienThiPhuAmCuoi = "t"
'        Case 8: HienThiPhuAmCuoi = ""
'    End Select
'End Function
