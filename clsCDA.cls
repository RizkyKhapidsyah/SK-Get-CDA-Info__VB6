VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCDA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private CDA_File As String ' used internal in functions

Private CDDrive As String ' i.e. E:

Dim ByteArray() As Byte, x As Byte
Dim z As Integer ' counter

Const SN_Offset As Long = 24
Const SN_Len As Long = 4
Const TrackLen_Offset As Long = 40 ' R-B
Const TrackLen_Len As Long = 4
Const TrackNum_Offset As Long = 23
Const TrackNum_Len As Long = 1 ' (*2) ' there are VERY few CD's with over 100 tracks...
Const TrackStart_Offset = 36
Const TrackStart_Len = 4
Const CDAVersion_Offset = 21
Const CDAVersion_Len = 1
Dim FileNum As Long ' A free file-number



Public Function NumberOfTracks() As Long
On Error Resume Next ' just in case... :)

Dim x As Long

NumberOfTracks = 0

For x = 1 To 99
    'check if the next track file exists!
    If Exists(GetFileName(x)) Then
    ' yes! it exists :)
    NumberOfTracks = NumberOfTracks + 1
Else
    ' no more tracks.... ;)
    Exit For
End If

Next

End Function

Public Function Serial_Number() As String
On Error Resume Next ' just in case... :)

Dim Chicken As String ' "temp" variable

' "Clear" the "temp" string... :)
Chicken = ""

' make 1 to e:\track01.cda, 2 to e:\track02.cda etc...
CDA_File = GetFileName(1)

If Not Exists(CDA_File) Then Exit Function
FileNum = FreeFile ' get a free file number

'open the file
Open CDA_File For Binary Access Read As FileNum

' redim and clear the array that holds the info we gets
ReDim ByteArray(SN_Len)

For z = 0 To SN_Len

' ok... the reason i've done it this way, is simple:
' if you do the following: SN_Offset + z (witch I begun with)
' the number will of some reason be BACKWARDS!!!
' i.e.: 01234610 will be 10462301!!!!!!
' so instead of reversing the string afterwards, it's
' easier let it start reading at the END of offset,
' and work backwards!

' put the bytes in the respective position in the array!
     Get FileNum, SN_Offset + SN_Len - z, ByteArray(z)
Next

Close FileNum

' Convert bytes to HEX, since Serialnumbers should be in
' HEX format, and put the byte-array into the string...

For z = 0 To SN_Len
Select Case ByteArray(z)
    Case Is = 0: Chicken = Chicken
    Case Is <= 9: Chicken = Chicken & "0" & Hex(ByteArray(z))
    Case Else: Chicken = Chicken & Hex(ByteArray(z))
End Select

' put a "-" in the middle of the string, to make it looks
' better...
Serial_Number = Left(Chicken, Len(Chicken) / 2)
Serial_Number = Serial_Number & "-"
Serial_Number = Serial_Number & Right(Chicken, Len(Chicken) / 2)

Next ' this is faster than using "Next z"

End Function


' Specify witch CD Drive to read from
Public Property Let CD_Drive(sDriveLetter As String)
On Error Resume Next ' just in case... :)

Dim LastCD As String
Dim NewCD As String

LastCD = CD_Drive ' in case we will restore it

If Len(sDriveLetter) <= 0 Then Exit Property

If Len(sDriveLetter) > 1 Then
    NewCD = Left(sDriveLetter, 1)
Else
    NewCD = sDriveLetter
End If

    If Exists(NewCD & ":\track01.cda") Then
        CDDrive = NewCD
    Else
        Exit Property
    End If


'If Not Exists(CDDrive & ":\*.*") Then CDDrive = ""
End Property

Public Property Get CD_Drive() As String
CD_Drive = UCase(CDDrive) & ":"
End Property

Public Function Track_Number(sTrack As Long) As Long
On Error Resume Next ' just in case... :)

' I have this function because sometimes you may want to know if the
' file has been renamed etc. This reads the track number from
' inside the file instead of just the filename.

' make 1 to e:\track01.cda, 2 to e:\track02.cda etc...
CDA_File = GetFileName(sTrack)

' check for file existence
If Not Exists(CDA_File) Then Exit Function

' Let VB give us a free file number :)
FileNum = FreeFile

'open the file
Open CDA_File For Binary Access Read As FileNum


' redim and clear the array that holds the info we gets
ReDim ByteArray(TrackNum_Len)
' use ByteArray for simple... even if we could have used a single
' byte varable... :)
' I'm putting this into a For loop with only 1 loop because
' it will be simple to make support for > 100 CD's, by changing
' TrackNum_Len variable :)

For z = 1 To TrackNum_Len ' 1... ;)

' SEE EXPLANATION ON THIS IN THE SERIAL_NUMBER() FUNCTION
     Get FileNum, TrackNum_Offset + TrackNum_Len - z, ByteArray(z)
Next

Close FileNum

Track_Number = ByteArray(1)

End Function

Public Function CDA_Version(sTrack As Long) As Long
On Error Resume Next ' just in case... :)

Dim Chicken As String ' "temp" variable
' "Clear" the "temp" string... :)
Chicken = ""

' make 1 to e:\track01.cda, 2 to e:\track02.cda etc...
CDA_File = GetFileName(sTrack)

' check for file existence
If Not Exists(CDA_File) Then Exit Function

' Let VB give us a free file number :)
FileNum = FreeFile

'open the file
Open CDA_File For Binary Access Read As FileNum



' redim and clear the array that holds the info we gets
ReDim ByteArray(CDAVersion_Len)

' use ByteArray for simple... even if we could have used a single
' byte varable... :)

For z = 1 To CDAVersion_Len ' 1... ;)
' SEE EXPLANATION ON THIS IN THE SERIAL_NUMBER() FUNCTION
     Get FileNum, CDAVersion_Offset + CDAVersion_Len - z, ByteArray(z)
Next

Close FileNum

Chicken = ByteArray(1)
CDA_Version = Chicken

End Function


Public Function Track_Length(sTrack As Long) As String
On Error Resume Next ' just in case... :)
'''''''''' RED-BOOK TRACK LENGTH ''''''''''''''
Dim Chicken As String ' "temp" variable

' make 1 to e:\track01.cda, 2 to e:\track02.cda etc...
CDA_File = GetFileName(sTrack)

' check for file existence
If Not Exists(CDA_File) Then Exit Function

' Let VB give us a free file number :)
FileNum = FreeFile

'open the file
Open CDA_File For Binary Access Read As FileNum



' redim and clear the array that holds the info we gets
ReDim ByteArray(TrackLen_Len)

For z = 1 To TrackLen_Len - 2  'since last bytes is unknown type.... :(

' SEE EXPLANATION ON THIS IN THE SERIAL_NUMBER() FUNCTION
     Get FileNum, TrackLen_Offset + TrackLen_Len - z, ByteArray(z)
Next

Close FileNum

' Format and put the getten bytes into a string with ":" beetwen
' each number...
For z = 1 To TrackLen_Len

Select Case ByteArray(z)
    Case 0:  'hmmm.... nothing here :)
    Case Else
        Chicken = Chicken & ByteArray(z) '+ 0
End Select

Next

Select Case Len(Trim(Chicken))
    Case 0: Track_Length = "N/A"
    Case 1: Track_Length = Chicken & ":00"
    Case 2: Track_Length = Left(Chicken, 1) & ":02"
    Case 3: Track_Length = Left(Chicken, 1) & ":" & Right(Chicken, 2) ': MsgBox Chicken
    Case 4: Track_Length = Left(Chicken, 2) & ":" & Right(Chicken, 2):
    Case Else: Track_Length = Chicken: ' MsgBox Len(Trim(Chicken))
End Select

End Function


Public Function Track_Start(sTrack As Long) As String
On Error Resume Next ' just in case... :)
Dim Chicken As String ' "temp" variable

' "Clear" the "temp" string... :)
Chicken = ""

' make 1 to e:\track01.cda, 2 to e:\track02.cda etc...
CDA_File = GetFileName(sTrack)

' check for file existence
If Not Exists(CDA_File) Then Exit Function

' Let VB give us a free file number :)
FileNum = FreeFile

'open the file
Open CDA_File For Binary Access Read As FileNum



' redim and clear the array that holds the info we gets
ReDim ByteArray(TrackStart_Len)

For z = 1 To TrackStart_Len - 2 'since last bytes is not in use....
     
' SEE EXPLANATION ON THIS IN THE SERIAL_NUMBER() FUNCTION
     Get FileNum, TrackStart_Offset + TrackStart_Len - z, ByteArray(z)
Next

Close FileNum
' Format and put the getten bytes into a string with ":" beetwen
' each number...
For z = 1 To TrackStart_Len

Select Case ByteArray(z)
    Case 0:
    Case Else
        If Chicken <> "" Then
            Chicken = Chicken & ":" & Format(ByteArray(z), "0#") ' 0#
        Else
            Chicken = Chicken & Format(ByteArray(z), "#") ' 0#
        End If
End Select

Next
    If Len(Chicken) = 2 Then Chicken = Chicken & ":00"
    ' because track 1 always will report wrong number of some reason:
    If Track_Number(sTrack) <= 1 Then Track_Start = "0:02" Else Track_Start = Chicken

End Function


Public Function Total_CD_Length() As String
On Error Resume Next ' just in case... :)


' sorry about not commenting this function better...
' it's too long since I made it ... ;)

Dim LastTrack As Long
Dim LastTrLen As String
Dim LastTrStart As String
Dim Temp As String
Dim Temp2 As String

LastTrack = NumberOfTracks ' number of last track
LastTrLen = Track_Length(LastTrack) ' length of last track
LastTrStart = Me.Track_Start(LastTrack) 'start of last track

Dim hour As Long
Dim min As Long
Dim sec As Long

' "temp" variables...
Temp = ""
Temp2 = ""

For x = 1 To Len(LastTrStart)
    If Mid(LastTrStart, x, 1) <> ":" Then
        Temp = Temp & Mid(LastTrStart, x, 1)
    End If
Next

For x = 1 To Len(LastTrLen)
    If Mid(LastTrLen, x, 1) <> ":" Then
        Temp2 = Temp2 & Mid(LastTrLen, x, 1)
    End If
Next

If Temp <> vbNullString And Temp2 <> vbNullString Then

    Total_CD_Length = CLng(Temp) + CLng(Temp2)
    Total_CD_Length = Format(Total_CD_Length, "##:##")
    
Else
    Total_CD_Length = ""
End If

End Function


'----------------------------------------------------
'Test file existence with error trapping
Private Function Exists(ByVal sSpec As String) As Boolean
Exists = False
Err = 0
On Error Resume Next
  
  Call FileLen(sSpec)
  Exists = (Err = 0)
'If Not Exists Then MsgBox "File not found!", vbCritical
End Function

Private Function GetFileName(sNumber As Long) As String
On Error Resume Next
If sNumber <= 0 Then GetFileName = "": Exit Function

If sNumber <= 9 Then
GetFileName = CDDrive & ":\track0" & sNumber & ".cda"
Else
GetFileName = CDDrive & ":\track" & sNumber & ".cda"
End If

End Function

