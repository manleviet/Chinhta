Attribute VB_Name = "mdlMain"
' Day la module chinh
Option Explicit
Public Enum Mode
    HienThi = 1
    GhiNhan = 2
End Enum

Public Type TypeWord
    PhuAmDau As Long
    ToHopNguyenAm As String * 5
    PhuAmCuoi As Integer
End Type

Public Words As TypeWord
Public ArrayWord() As TypeWord
Public iArrayWord As Integer

Public Sub Main()
    frmMain.Show
End Sub

Public Sub NewWords()
    ' Ham xu ly tao trong cho tu moi
    With Words
        .PhuAmDau = 0
        .ToHopNguyenAm = ""
        .PhuAmCuoi = 0
    End With
End Sub

Private Function GetFileName() As String
    Dim FileName As String
    FileName = App.Path
    If Right(FileName, 1) <> "\" Then
        FileName = FileName & "\"
    End If
    FileName = FileName & "DL.dat"
    GetFileName = FileName
End Function

Public Sub SaveArray()
    Dim FileNum As Integer
    FileNum = FreeFile
    Open GetFileName For Random As #FileNum Len = Len(Words)
    Dim i As Integer
    For i = 1 To iArrayWord
        Put #FileNum, i, ArrayWord(i - 1)
    Next i
    Close #FileNum
End Sub

Public Sub LoadArray()
    Dim FileNum As Integer
    FileNum = FreeFile
    Open GetFileName For Random As #FileNum Len = Len(Words)
    If FileLen(GetFileName) <> 0 Then
        iArrayWord = FileLen(GetFileName) / Len(Words)
        ReDim ArrayWord(iArrayWord)
        Dim i As Integer
        For i = 1 To iArrayWord
            Get #FileNum, i, ArrayWord(i - 1)
        Next i
    Else
        iArrayWord = 0
    End If
    Close #FileNum
'    Kill GetFileName
End Sub

Public Sub Sort()
    Dim i As Integer
    Dim j As Integer
    For i = 0 To iArrayWord - 2
        For j = i + 1 To iArrayWord - 1
            If SoSanh(ArrayWord(i).ToHopNguyenAm, ArrayWord(j).ToHopNguyenAm) = 1 Then
                Swap i, j
            End If
        Next j
    Next i
End Sub

Private Sub Swap(i As Integer, j As Integer)
    Dim temp As TypeWord
    temp = ArrayWord(i)
    ArrayWord(i) = ArrayWord(j)
    ArrayWord(j) = temp
End Sub

Public Function FindBinary(str As String) As Integer
    Dim Low As Integer
    Dim High As Integer
    Dim Mid As Integer
    Low = 0
    High = iArrayWord - 1
    Do While High >= Low
        Mid = (High + Low) \ 2
        Select Case SoSanh(ArrayWord(Mid).ToHopNguyenAm, str)
            Case 1: High = Mid - 1
            Case -1: Low = Mid + 1
            Case 0: Exit Do
        End Select
    Loop
    If High >= Low Then
        FindBinary = Mid
    Else
        FindBinary = -1
    End If
End Function

Private Function SoSanh(ByVal st1 As String, ByVal st2 As String) As Integer
    Dim Nho As Integer
    Dim i As Integer
    i = 1
    Nho = 0
    Do While i <= 3
        If Mid(st1, i, 1) > Mid(st2, i, 1) Then
            Nho = 1
            Exit Do
        ElseIf Mid(st1, i, 1) < Mid(st2, i, 1) Then
            Nho = -1
            Exit Do
        Else
            i = i + 1
        End If
    Loop
    If i > 3 Then
        SoSanh = 0
    Else
        SoSanh = Nho
    End If
End Function

Public Sub SaveWord()
    Dim Pos As Integer
    Pos = FindBinary(Words.ToHopNguyenAm)
    If iArrayWord <> 0 And Pos <> -1 Then
        Dim i As Integer
        For i = 0 To 26
            If ((Words.PhuAmDau And 2 ^ i) <> 0) And ((ArrayWord(Pos).PhuAmDau And 2 ^ i) = 0) Then
                ArrayWord(Pos).PhuAmDau = ArrayWord(Pos).PhuAmDau + 2 ^ i
            End If
        Next i
        For i = 0 To 8
            If (Words.PhuAmCuoi And 2 ^ i <> 0) And (ArrayWord(Pos).PhuAmCuoi And 2 ^ i = 0) Then
                ArrayWord(Pos).PhuAmCuoi = ArrayWord(Pos).PhuAmCuoi + 2 ^ i
            End If
        Next i
    Else
        iArrayWord = iArrayWord + 1
        ReDim Preserve ArrayWord(iArrayWord)
        ArrayWord(iArrayWord - 1) = Words
        Sort
    End If
End Sub

Public Sub PhanTich(ByVal st As String)
    Dim LenDau As Integer
    Dim LenCuoi As Integer
    Dim Kt As String
    NewWords
    Kt = Right(st, 1)
    LenCuoi = 1
    Select Case Kt
        Case "c": Words.PhuAmCuoi = Words.PhuAmCuoi + 2 ^ 0
        Case "m": Words.PhuAmCuoi = Words.PhuAmCuoi + 2 ^ 2
        Case "n": Words.PhuAmCuoi = Words.PhuAmCuoi + 2 ^ 3
        Case "t": Words.PhuAmCuoi = Words.PhuAmCuoi + 2 ^ 7
        Case "p": Words.PhuAmCuoi = Words.PhuAmCuoi + 2 ^ 6
        Case "h": Kt = Right(st, 2)
                  LenCuoi = 2
                  Select Case Kt
                    Case "ch": Words.PhuAmCuoi = Words.PhuAmCuoi + 2 ^ 1
                    Case "nh": Words.PhuAmCuoi = Words.PhuAmCuoi + 2 ^ 5
                  End Select
        Case "g": Words.PhuAmCuoi = Words.PhuAmCuoi + 2 ^ 4
                  LenCuoi = 2
        Case Else:  LenCuoi = 0
                    Words.PhuAmCuoi = Words.PhuAmCuoi + 2 ^ 8
    End Select
    Kt = Left(st, 1)
    LenDau = 1
    Select Case Kt
        Case "b": Words.PhuAmDau = Words.PhuAmDau + 2 ^ 0
        Case "d": Words.PhuAmDau = Words.PhuAmDau + 2 ^ 3
        Case "â": Words.PhuAmDau = Words.PhuAmDau + 2 ^ 4
        Case "h": Words.PhuAmDau = Words.PhuAmDau + 2 ^ 7
        Case "l": Words.PhuAmDau = Words.PhuAmDau + 2 ^ 10
        Case "m": Words.PhuAmDau = Words.PhuAmDau + 2 ^ 11
        Case "q": Words.PhuAmDau = Words.PhuAmDau + 2 ^ 18
        Case "r": Words.PhuAmDau = Words.PhuAmDau + 2 ^ 19
        Case "s": Words.PhuAmDau = Words.PhuAmDau + 2 ^ 20
        Case "v": Words.PhuAmDau = Words.PhuAmDau + 2 ^ 24
        Case "x": Words.PhuAmDau = Words.PhuAmDau + 2 ^ 25
        Case "c": Kt = Left(st, 2)
                  If Kt = "ch" Then
                    LenDau = 2
                    Words.PhuAmDau = Words.PhuAmDau + 2 ^ 2
                  Else
                    Words.PhuAmDau = Words.PhuAmDau + 2 ^ 1
                  End If
        Case "g": Kt = Left(st, 2)
                  If Kt = "gh" Then
                    LenDau = 2
                    Words.PhuAmDau = Words.PhuAmDau + 2 ^ 6
                  Else
                    Words.PhuAmDau = Words.PhuAmDau + 2 ^ 5
                  End If
        Case "k": Kt = Left(st, 2)
                  If Kt = "kh" Then
                    LenDau = 2
                    Words.PhuAmDau = Words.PhuAmDau + 2 ^ 9
                  Else
                    Words.PhuAmDau = Words.PhuAmDau + 2 ^ 8
                  End If
        Case "p": Kt = Left(st, 2)
                  If Kt = "ph" Then
                    LenDau = 2
                    Words.PhuAmDau = Words.PhuAmDau + 2 ^ 17
                  Else
                    Words.PhuAmDau = Words.PhuAmDau + 2 ^ 16
                  End If
        Case "t": Kt = Left(st, 2)
                  If Kt = "th" Then
                    LenDau = 2
                    Words.PhuAmDau = Words.PhuAmDau + 2 ^ 22
                  ElseIf Kt = "tr" Then
                    LenDau = 2
                    Words.PhuAmDau = Words.PhuAmDau + 2 ^ 23
                  Else
                    Words.PhuAmDau = Words.PhuAmDau + 2 ^ 21
                  End If
        Case "n": Kt = Left(st, 2)
                  If Kt = "nh" Then
                    LenDau = 2
                    Words.PhuAmDau = Words.PhuAmDau + 2 ^ 15
                  ElseIf Kt = "ng" Then
                    LenDau = 2
                    Kt = Left(st, 3)
                    If Kt = "ngh" Then
                        Words.PhuAmDau = Words.PhuAmDau + 2 ^ 14
                        LenDau = 3
                    Else
                        Words.PhuAmDau = Words.PhuAmDau + 2 ^ 13
                    End If
                  Else
                    Words.PhuAmDau = Words.PhuAmDau + 2 ^ 12
                  End If
        Case Else:  LenDau = 0
                    Words.PhuAmDau = Words.PhuAmDau + 2 ^ 26
    End Select
    Words.ToHopNguyenAm = Mid(st, LenDau + 1, Len(st) - LenDau - LenCuoi)
End Sub
