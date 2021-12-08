VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Get Audio CD Info"
   ClientHeight    =   3168
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3168
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   2
      Text            =   "E"
      Top             =   2160
      Width           =   372
   End
   Begin VB.HScrollBar track 
      Height          =   372
      Left            =   240
      TabIndex        =   0
      Top             =   2640
      Width           =   3132
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "CD Drive:"
      Height          =   192
      Left            =   1800
      TabIndex        =   3
      Top             =   2280
      Width           =   684
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   192
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   492
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CDA As New clsCDA
Dim aTrack As Long

Private Sub Form_Load()

With CDA
    .CD_Drive = Text1.Text
    track.min = 1
    track.Max = .NumberOfTracks
End With

End Sub

Private Sub Text1_Change()
CDA.CD_Drive = Text1.Text
track.Max = CDA.NumberOfTracks
Call track_Change 'for update
End Sub

Private Sub track_Change()
aTrack = track.Value

' fill the label's or whatever with the CDA info
With CDA
    Label1.Caption = "CD Drive: " & .CD_Drive & vbCrLf & "Serial Number: " & .Serial_Number & vbCrLf _
    & "Number of tracks: " & .NumberOfTracks & vbCrLf _
    & "CDA Version: " & .CDA_Version(aTrack) & vbCrLf _
    & "Track: " & .Track_Number(aTrack) & vbCrLf & "Track Start: " & .Track_Start(aTrack) & vbCrLf _
    & "Track Length: " & .Track_Length(aTrack) & vbCrLf & _
    "Total CD Length: " & .Total_CD_Length
End With
End Sub
