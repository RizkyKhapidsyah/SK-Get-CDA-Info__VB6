Attribute VB_Name = "Module1"
'
' This program retrieves info about music CD's, by reading
' the .CDA files i the root of CD. It get's track number,
' serial-number, CDA version, beginning of track in Red-Book
' format, Length of the track in Red-Book format, Total CD length,
' and Total numbers of Tracks!
' This without using a single API call to Winmm.dll !!!

' OK, enough chitchat! ;)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Declare variables:

Public Version As String
Public Serial_Number As String
Public Track_Number As Integer
Public Track_Length As String
Public TrackStart As String
Public CDA_Version As Long

Dim FileNum As Long ' A free file-number
Dim CDA As New clsCDA

Sub main()


Form1.Show


End Sub



Public Function HexToDecimal(ByVal pValue As String) As Long
  Dim C         As Integer
  Dim Value     As Long
  Dim PosValue  As Long
  Dim CharVal   As Byte
  
  On Error GoTo ErrorHandler
  
  For C = Len(pValue) To 1 Step -1
    Select Case Mid(pValue, C, 1)
      Case "0" To "9": CharVal = CInt(Mid(pValue, C, 1))
      Case Else: CharVal = Asc(Mid(pValue, C, 1)) - 55
    End Select
    Value = Value + ((16 ^ PosValue) * CharVal)
    PosValue = PosValue + 1
  Next C
  
  HexToDecimal = Value
  Exit Function
  
ErrorHandler:
  HexToDecimal = 0

End Function


'----------------------------------------------------
'This function converts 4*1 byte array to 4*8 bits
Private Function ByteToBit(ByteArray) As String
Dim z As Integer, i As Integer

  ByteToBit = ""
  For z = 1 To 4
    For i = 7 To 0 Step -1
      If Int(ByteArray(z) / (2 ^ i)) = 1 Then
        ByteToBit = ByteToBit & "1"
        ByteArray(z) = ByteArray(z) - (2 ^ i)
      Else
        If ByteToBit <> "" Then
          ByteToBit = ByteToBit & "0"
        End If
      End If
    Next i
  Next z
  
End Function
'----------------------------------------------------

'----------------------------------------------------
'This function converts Binary string to decimal integer
Private Function BinToDec(BinValue As String) As Long
Dim i As Integer

  BinToDec = 0
  For i = 1 To Len(BinValue)
    If Mid(BinValue, i, 1) = 1 Then
      BinToDec = BinToDec + 2 ^ (Len(BinValue) - i)
    End If
  Next i
  
End Function
'----------------------------------------------------


