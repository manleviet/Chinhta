VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private strText As String      ' the whole string
Private strDelimiter() As String ' the delimiter
Private IMaxToken As Long  ' number of tokens
' Array of tokens after breaking up the string
Private Tokens() As String
'  Current character cursor position in the string, used when you move sequentially from one token to the next
Private IPos As Long
' Current length of the string
Private ILen As Long
Public Property Get Text() As Variant
' Return the whole string
   Text = strText
End Property
Public Property Let Text(ByVal vNewValue As Variant)
' Assign the Text string
    strText = Trim(vNewValue)
    ILen = Len(strText)  ' Update length of string
    IPos = 1
    Tokenise
End Property
Public Function TokenAt(TNum) As String
' Return the Token whose arrayIndex is TNum.
' Tokens in the array are indexed in the range 1..IMaxToken
   If (TNum > 0) And (TNum <= IMaxToken) Then
      TokenAt = Tokens(TNum)
   Else
      TokenAt = ""
   End If
End Function
Private Sub Tokenise()
' Break up the string at the delimiters into the tokens.
' Save the tokens in the array indexing in the range 1..IMaxToken
   Dim i
   i = 0: IPos = 1
   Do Until IPos > ILen
      i = i + 1
      ' Make room for one more token. Preserve the previous tokens
      ReDim Preserve Tokens(i)
      ' Obtain the next token from the string
      Tokens(i) = GetToken
   Loop
   ' Update total number of Tokens
   IMaxToken = i
   ' Reset character Cursor position
   IPos = 1
End Sub
Public Function GetToken() As String
' Return the next token
' Use this function when you move sequentially from one token to the next
    Dim Pos
    Pos = IPos
    Do While IsDelimiter(Mid(strText, Pos, 1))
        Pos = Pos + 1
    Loop
    IPos = Pos
    Do While Not IsDelimiter(Mid(strText, Pos, 1)) And Pos <= ILen
        Pos = Pos + 1
    Loop
    GetToken = Mid(strText, IPos, Pos - IPos)
    IPos = Pos + 1
End Function
Public Property Get TokenCount() As Variant
' Return the number of tokens
   TokenCount = IMaxToken
End Property
Private Function IsDelimiter(st As String) As Boolean
    If st <> "" Then
        Select Case Asc(st)
            Case 0 To 64: GoTo a
            Case 91 To 96: GoTo a
            Case 123 To 126: GoTo a
        End Select
        IsDelimiter = False
        Exit Function
a: IsDelimiter = True
    End If
End Function
