VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "StringFX by Stretch v2.10"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   7950
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame7 
      Caption         =   "FX Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3240
      TabIndex        =   39
      Top             =   1320
      Width           =   3135
      Begin VB.OptionButton Option7 
         Caption         =   "4"
         Enabled         =   0   'False
         Height          =   195
         Index           =   3
         Left            =   2160
         TabIndex        =   43
         ToolTipText     =   "Not Used"
         Top             =   480
         Width           =   375
      End
      Begin VB.OptionButton Option7 
         Caption         =   "3"
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   1520
         TabIndex        =   42
         ToolTipText     =   "Not Used"
         Top             =   480
         Width           =   375
      End
      Begin VB.OptionButton Option7 
         Caption         =   "2"
         Height          =   195
         Index           =   1
         Left            =   880
         TabIndex        =   41
         Top             =   480
         Width           =   375
      End
      Begin VB.OptionButton Option7 
         Caption         =   "1"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   40
         Top             =   480
         Value           =   -1  'True
         Width           =   375
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Speed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3240
      TabIndex        =   35
      Top             =   6720
      Width           =   4215
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1320
         TabIndex        =   38
         Text            =   "1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Fastest"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   37
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Timer"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   36
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Remove/Leave or Changecase to"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   30
      Top             =   5880
      Width           =   4455
      Begin VB.OptionButton Option5 
         Caption         =   "Both"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Lowercase"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   32
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Uppercase"
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   31
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Remove/Leave"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3240
      TabIndex        =   25
      Top             =   4800
      Width           =   3135
      Begin VB.OptionButton Option4 
         Caption         =   "Spaces"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Numbers"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   28
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Letters"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   27
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Punctuation"
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   26
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Direction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   20
      Top             =   3960
      Width           =   3135
      Begin VB.OptionButton Option3 
         Caption         =   "To Left"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "To Right"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   23
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "On/Off Screen"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   19
      Top             =   3000
      Width           =   3135
      Begin VB.OptionButton Option2 
         Caption         =   "On"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Off"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "FX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   360
      TabIndex        =   6
      Top             =   1320
      Width           =   2415
      Begin VB.OptionButton Option1 
         Caption         =   "Reverse"
         Height          =   375
         Index           =   12
         Left            =   360
         TabIndex        =   33
         Top             =   6000
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Remove"
         Height          =   375
         Index           =   10
         Left            =   360
         TabIndex        =   18
         Top             =   5060
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Random"
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   17
         Top             =   1300
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Scroll String"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   16
         Top             =   830
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Rotate String"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   15
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Type"
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   14
         Top             =   1770
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Changecase"
         Height          =   375
         Index           =   9
         Left            =   360
         TabIndex        =   13
         Top             =   4590
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Type Middle"
         Height          =   375
         Index           =   4
         Left            =   360
         TabIndex        =   12
         Top             =   2240
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Black Hole"
         Height          =   375
         Index           =   5
         Left            =   360
         TabIndex        =   11
         Top             =   2710
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Blinds"
         Height          =   375
         Index           =   6
         Left            =   360
         TabIndex        =   10
         Top             =   3180
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Type Slide"
         Height          =   375
         Index           =   7
         Left            =   360
         TabIndex        =   9
         Top             =   3650
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Leave"
         Height          =   375
         Index           =   11
         Left            =   360
         TabIndex        =   8
         Top             =   5530
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Bandit"
         Height          =   375
         Index           =   8
         Left            =   360
         TabIndex        =   7
         Top             =   4120
         Width           =   1455
      End
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   6960
      TabIndex        =   5
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   495
      Left            =   6960
      TabIndex        =   4
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Text            =   "-1.  An example by Stretch!  -2.  An example by Stretch!"
      ToolTipText     =   "no more than 60 characters"
      Top             =   240
      Width           =   6735
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   600
      Width           =   6735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Text"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim StrTemp, StrMask, SingleChr As String
Dim SingleChrType As Byte
Dim FXCounter, OldFXCounter, FXCounter2 As Integer
Dim Loop1 As Integer
Dim TextLen As Integer
Dim ChosenFX, ScreenFX, DirectionFX, RLChrGroupFX, OldRLChrGroupFX, LUCaseFX, ChosenSpeedFX, VersionFX As Integer
Dim RndChosenChr As Integer

Private Sub Command1_Click()
'Find out which FX has been chosen
For Loop1 = 0 To 11
    If Option1(Loop1).Value = True Then ChosenFX = Loop1
Next Loop1
Frame1.Enabled = False
If Frame2.Enabled = True Then
    For Loop1 = 0 To 1
        If Option2(Loop1).Value = True Then ScreenFX = Loop1
    Next Loop1
    Frame2.Enabled = False
End If
If Frame3.Enabled = True Then
    For Loop1 = 0 To 1
        If Option3(Loop1).Value = True Then DirectionFX = Loop1
    Next Loop1
    Frame3.Enabled = False
End If
If Frame4.Enabled = True Then
    For Loop1 = 0 To 3
        If Option4(Loop1).Value = True Then RLChrGroupFX = Loop1
    Next Loop1
    Frame4.Enabled = False
End If
If Frame5.Enabled = True Then
    For Loop1 = 0 To 2
        If Option5(Loop1).Value = True Then LUCaseFX = Loop1
    Next Loop1
    Frame5.Enabled = False
End If
Frame6.Enabled = False
If Frame7.Enabled = True Then
    For Loop1 = 0 To 3
        If Option7(Loop1).Value = True Then VersionFX = Loop1
    Next Loop1
    Frame7.Enabled = False
End If
If Option2(0).Value = False Or Option1(0).Value = True Then
    StrTemp = Text1.Text
    Label3.Caption = StrTemp
Else
    StrTemp = String$(Len(Text1.Text), " ")
    Label3.Caption = ""
End If
StrMask = String$(TextLen, ".")
FXCounter = 0
OldFXCounter = 0
Command1.Enabled = False
Command2.Caption = "Stop"

If Option6(0).Value = True Then
'Start chosen FX with set milliseconds between each change
    Timer1.Interval = Text2.Text
Else
    MainStringFX
End If
End Sub

Private Sub Command2_Click()
'Finish chosen FX, unless it is already finished in which case shutdown the form!
If Command2.Caption = "Exit" Then
' the caption "Exit" is on the stop button, so the user no longer wishes to run the program
    Unload Form1
Else
' the caption "Stop" is on the stop button, so the user wishes to cancel the current operation
    Timer1.Interval = 0
    Command1.Enabled = True
    Command2.Caption = "Exit"
    Frame1.Enabled = True
    Frame1_Click
End If
End Sub

Private Sub Form_Load()
Text1_Change
Frame1_Click
End Sub

Private Sub Frame1_Click()
'The following bit of code is really for asthetics only. Basically each Option in the FX frame only requires certain parameters.
'This piece of code makes sure that the user can only select from those required parameters, by enabling them.
'The others are disabled.
'I have noticed that although the option buttons cannot be clicked when a frame is disabled, they are still shown.
'To make it clear that they are disabled I have created a subroutine called FrameSelector which also enables/disables
'the options in the Frame.
Option5(0).Enabled = True
'Find out which Option has been selected
For Loop1 = 0 To 12
    If Option1(Loop1).Value = True Then ChosenFX = Loop1
Next Loop1
Select Case ChosenFX
    Case 0
        FrameSelector False, True, False, False, False
        FrameSelector7 False
        FrameCaption True
    Case 1
        FrameSelector True, True, False, False, False
        FrameSelector7 False
        FrameCaption True
    Case 2
        FrameSelector True, False, False, False, False
        FrameSelector7 False
        FrameCaption True
    Case 3
        FrameSelector True, True, False, False, False
        FrameSelector7 False
        FrameCaption True
    Case 4
        FrameSelector True, True, False, False, False
        FrameSelector7 False
        FrameCaption False
    Case 5
        FrameSelector True, True, False, False, False
        FrameSelector7 False
        FrameCaption False
    Case 6
        FrameSelector True, True, False, False, False
        FrameSelector7 False
        FrameCaption True
    Case 7
        FrameSelector True, True, False, False, True
        FrameSelector7 False
        FrameCaption True
    Case 8
        FrameSelector True, True, False, False, True
        FrameSelector7 True
        FrameCaption True
        Frame7_Click
    Case 9
        FrameSelector False, False, False, True, True
        FrameSelector7 False
        FrameCaption True
        Option5(0).Enabled = False
        Option5(1).Value = True
    Case 10
        FrameSelector False, False, True, False, True
        FrameSelector7 False
        FrameCaption True
        Frame4_Click
    Case 11
        FrameSelector False, False, True, False, True
        FrameSelector7 False
        FrameCaption True
        Frame4_Click
    Case 12
        FrameSelector False, True, False, False, False
        FrameSelector7 False
        FrameCaption True
End Select
End Sub

Private Sub Frame4_Click()
'Find out which Option has been selected
For Loop1 = 0 To 3
    If Option4(Loop1).Value = True Then RLChrGroupFX = Loop1
Next Loop1
'If RLChrGroupFX (Remove/Leave Character Group FX panel), Option4(2) [Letters] is selected then allow access to Panel 5,
'otherwise dont.
'Although the loop is Zero to Three, it should be remembered that there are in fact 4 options in the array.
Select Case RLChrGroupFX
    Case 2
        FrameSelector False, False, True, True, True
        FrameSelector7 False
        If OldRLChrGroupFX <> 2 Then Option5(0).Value = True
    Case Else
        FrameSelector False, False, True, False, True
        FrameSelector7 False
End Select
OldRLChrGroupFX = RLChrGroupFX
End Sub

Private Sub Frame6_Click()
'If Fastest has been selected then disable the Timer Text box as it is not needed
If Option6(1).Value = True Then
    Text2.Enabled = False
Else
    Text2.Enabled = True
End If
End Sub

Private Sub Frame7_Click()
'Find out which Option has been selected
For Loop1 = 0 To 3
    If Option7(Loop1).Value = True Then VersionFX = Loop1
Next Loop1
Select Case ChosenFX
    Case 8
        Select Case VersionFX
            Case 0
                FrameSelector True, True, False, False, True
            Case 1
                FrameSelector True, False, False, False, True
        End Select
End Select
End Sub

Private Sub Option1_Click(Index As Integer)
'A change has occured in the FX Panel - goto Frame1_Click
'(It saves duplicating the code!)
Frame1_Click
End Sub

Private Sub Option4_Click(Index As Integer)
'A change has occured in the Remove/Leave Panel - goto Frame4_Click
Frame4_Click
End Sub

Private Sub Option6_Click(Index As Integer)
'A change has occured in the Speed Panel - goto Frame4_Click
Frame6_Click
End Sub

Private Sub Option7_Click(Index As Integer)
'A change has occured in the Version Panel - goto Frame7_Click
Frame7_Click
End Sub

Private Sub Text1_Change()
'Make sure the text is 60 characters long - there is no real reason for this length, I just decided it was long enough. :)
If Text1.Text = "" Then Text1.Text = "There should be text here"
TextLen = Len(RTrim(Text1.Text))
If Int(TextLen / 2) * 2 <> TextLen Then TextLen = TextLen + 1
Label3.Caption = TextLen
If Len(Text1.Text) < 60 Then
    Text1.Text = Left$(Text1.Text + String$(60, " "), 60)
Else
    Text1.Text = Left$(Text1.Text, 60)
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
'force text2 to only accept numbers
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
End If
End Sub

Private Sub Timer1_Timer()
StringFX
End Sub

Private Sub FrameSelector(ByVal F2 As Boolean, ByVal F3 As Boolean, ByVal F4 As Boolean, ByVal F5 As Boolean, ByVal F6 As Boolean)
Dim n As Integer
Frame2.Enabled = F2
If F2 Then
    Option2(0).Enabled = True
    Option2(1).Enabled = True
Else
    Option2(0).Enabled = False
    Option2(1).Enabled = False
    Option2(0).Value = True
End If
Frame3.Enabled = F3
If F3 Then
    Option3(0).Enabled = True
    Option3(1).Enabled = True
Else
    Option3(0).Enabled = False
    Option3(1).Enabled = False
End If
Frame4.Enabled = F4
If F4 Then
    For n = 0 To 3
        Option4(n).Enabled = True
    Next n
Else
    For n = 0 To 3
        Option4(n).Enabled = False
    Next n
End If
Frame5.Enabled = F5
If F5 Then
    For n = 0 To 2
        Option5(n).Enabled = True
    Next n
Else
    For n = 0 To 2
        Option5(n).Enabled = False
    Next n
End If
Frame6.Enabled = F6
If F6 Then
    Option6(1).Enabled = True
Else
    Option6(1).Enabled = False
    Option6(0).Value = True
End If
End Sub

Private Sub FrameSelector7(ByVal F7 As Boolean)
Dim n As Integer
Frame7.Enabled = F7
If F7 Then
    For n = 0 To 1
        Option7(n).Enabled = True
    Next n
Else
    For n = 0 To 1
        Option7(n).Enabled = False
    Next n
    Option7(0).Value = True
End If
End Sub

Private Sub FrameCaption(ByVal Horizontal As Boolean)
If Horizontal = True Then
    Option3(0).Caption = "To Left"
    Option3(1).Caption = "To Right"
Else
    Option3(0).Caption = "Into Centre"
    Option3(1).Caption = "Out of Centre"
End If
End Sub


Private Function GetCharType(ByVal char As String) As Byte
Dim CharType As Byte
'0=Space
'1=Number
'2=Letter
'3=Punctuation
'4=Lowercase
'5=Uppercase

CharType = 3
'Is it a space?
If Asc(char) = 32 Then CharType = 0
'Is it a Number?
If Asc(char) >= 48 And Asc(char) <= 57 Then CharType = 1
'Is it lowercase?
If Asc(char) >= 97 And Asc(char) <= 122 Then CharType = 4
'Is it Uppercase?
If Asc(char) >= 65 And Asc(char) <= 90 Then CharType = 5

If CharType >= 4 And LUCaseFX = 0 Then CharType = 2

GetCharType = CharType
End Function

Private Sub MainStringFX()
Do Until Frame1.Enabled = True
    StringFX
Loop
End Sub

Private Sub StringFX()
DoEvents
'Apply chosen calculation each time the Timer hits zero
Select Case ChosenFX
    Case 0
    'Rotate String
        Select Case DirectionFX
            Case 0
                'Take the right hand character and add it to the beginning of the string, then remove it from the end
                StrTemp = Right$(StrTemp, Len(StrTemp) - 1) + Left$(StrTemp, 1)
            Case 1
                'Take the left hand character and add it to the end of the string, then remove it from the beginning
                StrTemp = Right$(StrTemp, 1) + Left$(StrTemp, Len(StrTemp) - 1)
        End Select
        Label3.Caption = StrTemp
    Case 1
    'Scroll String
        FXCounter = FXCounter + 1
        'How about this for a bit of boolean logic!!!
        If DirectionFX + ScreenFX = 1 Then
            If FXCounter > TextLen Then
                    Command2_Click
            Else
                If DirectionFX = 0 Then
                    'take the first character away from the new string
                    'Left - Off
                    StrTemp = Right$(StrTemp, Len(StrTemp) - 1)
                Else
                    'put the last n characters from the original string into the new string
                    'Right - On
                    StrTemp = Right$(Text1.Text, FXCounter + (Len(Text1.Text) - TextLen))
                End If
                Label3.Caption = StrTemp
            End If
        Else
            If FXCounter > Len(Text1.Text) Then
                Command2_Click
            Else
                If DirectionFX = 0 Then
                    'put the first n characters from the original string into the end of the new string
                    'Left - On
                    StrTemp = String$(Len(Text1.Text) - FXCounter, " ") + Left$(Text1.Text, FXCounter)
                Else
                    'take the last character away from the new string and add a space to the beginning
                    'Right - Off
                    StrTemp = " " + Left$(StrTemp, Len(StrTemp) - 1)
                End If
                Label3.Caption = StrTemp
            End If
        End If
    Case 2
    'Random
        'Randomly pick a character from the line that has not been chosen yet...
        RndChosenChr = Int(Rnd * TextLen) + 1
        'This speeds up the choosing process. If it doesnt find anything then look for the first occurance to the right
        'of the randomly chosen position
        If Mid$(StrMask, RndChosenChr, 1) = " " Then
            FXCounter2 = InStr(RndChosenChr, StrMask, ".")
            'If it still hasnt found anything then look from the beginning of the string (to the right)
            If FXCounter2 = 0 Then FXCounter2 = InStr(StrMask, ".")
            RndChosenChr = FXCounter2
        End If
        ' The mask is more a way of counting how many characters have been found
        If Mid$(StrMask, RndChosenChr, 1) = "." Then
            FXCounter = FXCounter + 1
            If ScreenFX = 0 Then
                '... and slot it in position for displaying.
                Mid(StrTemp, RndChosenChr, 1) = Mid$(Text1.Text, RndChosenChr, 1)
            Else
                '... and remove it from the string.
                Mid(StrTemp, RndChosenChr, 1) = " "
            End If
            'Set the character at the same position in the mask to null signifying that another character has been selected,
            'and cannot be chosen again.
            Mid(StrMask, RndChosenChr, 1) = " "
            Label3.Caption = StrTemp
        End If
        If FXCounter = TextLen Then Command2_Click
    Case 3
    'Type 1 character at a time
        Select Case DirectionFX
            Case 0
                FXCounter = FXCounter + 1
                If FXCounter > TextLen Then
                    Command2_Click
                Else
                    If ScreenFX = 0 Then
                        'Take each character from right to left and put it at the end of the new string
                        Mid(StrTemp, TextLen - (FXCounter - 1), FXCounter) = Right$(Text1.Text, (Len(Text1.Text) - TextLen) + FXCounter)
                    Else
                        'Take away each character from right to left of the new string
                        Mid(StrTemp, TextLen - (FXCounter - 1), 1) = " "
                    End If
                    Label3.Caption = StrTemp
                End If
            Case 1
                FXCounter = FXCounter + 1
                If FXCounter > TextLen Then
                    Command2_Click
                Else
                    If ScreenFX = 0 Then
                        'Take each character from left to right and put it at the end of the new string
                        StrTemp = Left$(Text1.Text, FXCounter)
                    Else
                        'Take away each character from left to right of the new string
                        Mid(StrTemp, FXCounter, 1) = " "
                    End If
                    Label3.Caption = StrTemp
                End If
        End Select
    Case 4
    'Wipe Middle
        Select Case DirectionFX
            Case 0
            'Into Middle
                FXCounter = FXCounter + 1
                If FXCounter > (TextLen / 2) + 1 Then
                    Command2_Click
                Else
                    If ScreenFX = 0 Then
                        Mid(StrTemp, FXCounter, 1) = Mid$(Text1.Text, FXCounter, 1)
                        Mid(StrTemp, TextLen - (FXCounter - 1), 1) = Mid$(Text1.Text, TextLen - (FXCounter - 1), 1)
                    Else
                        Mid(StrTemp, FXCounter, 1) = " "
                        Mid(StrTemp, TextLen - (FXCounter - 1), 1) = " "
                    End If
                End If
                Label3.Caption = StrTemp
            Case 1
            'Out of Middle
                FXCounter = FXCounter + 1
                If FXCounter > (TextLen / 2) Then
                    Command2_Click
                Else
                    If ScreenFX = 0 Then
                        Mid(StrTemp, (TextLen / 2) - (FXCounter - 1), 1) = Mid$(Text1.Text, (TextLen / 2) - (FXCounter - 1), 1)
                        Mid(StrTemp, (TextLen / 2) + FXCounter, 1) = Mid$(Text1.Text, (TextLen / 2) + FXCounter, 1)
                    Else
                        Mid(StrTemp, (TextLen / 2) - (FXCounter - 1), 1) = " "
                        Mid(StrTemp, (TextLen / 2) + FXCounter, 1) = " "
                    End If
                End If
                Label3.Caption = StrTemp
        End Select
    Case 5
        If FXCounter = 0 Then
            If ScreenFX = 0 Then
                StrTemp = String$(60, " ")
            Else
                StrTemp = Text1.Text
            End If
        End If
        Select Case DirectionFX
            Case 0
            'Into Middle
                FXCounter = FXCounter + 1
                If FXCounter > 30 Then
                    Command2_Click
                Else
                    StrTemp = Left$(StrTemp, 29) + Right$(StrTemp, 29)
                    If ScreenFX = 0 Then
                        StrTemp = Mid$(Text1.Text, 31 - FXCounter, 1) + StrTemp + Mid$(Text1.Text, 30 + FXCounter, 1)
                    Else
                        StrTemp = " " + StrTemp + " "
                    End If
                End If
                Label3.Caption = StrTemp
            Case 1
            'Out of Middle
                FXCounter = FXCounter + 1
                If FXCounter > 30 Then
                    Command2_Click
                Else
                    If ScreenFX = 0 Then
                        StrTemp = Mid$(StrTemp, 2, 29) + Mid$(Text1.Text, FXCounter, 1) + Mid$(Text1.Text, 61 - FXCounter, 1) + Mid$(StrTemp, 31, 29)
                    Else
                        StrTemp = Mid$(StrTemp, 2, 29) + "  " + Mid$(StrTemp, 31, 29)
                    End If
                End If
                Label3.Caption = StrTemp
        End Select
    Case 6
    'Blinds
        Select Case DirectionFX
            Case 0
                FXCounter = FXCounter + 1
                If FXCounter = 6 Then
                    Command2_Click
                Else
                    For Loop1 = 60 To 5 Step -5
                        If ScreenFX = 0 Then
                            If Loop1 - FXCounter > 0 Then Mid(StrTemp, Loop1 - FXCounter, 1) = Mid$(Text1.Text, Loop1 - FXCounter, 1)
                        Else
                            If Loop1 - FXCounter > 0 Then Mid(StrTemp, Loop1 - FXCounter, 1) = " "
                        End If
                    Next Loop1
                End If
            Case 1
                FXCounter = FXCounter + 1
                If FXCounter = 6 Then
                    Command2_Click
                Else
                    For Loop1 = 0 To 55 Step 5
                        If ScreenFX = 0 Then
                            Mid(StrTemp, Loop1 + FXCounter, 1) = Mid$(Text1.Text, Loop1 + FXCounter, 1)
                        Else
                            Mid(StrTemp, Loop1 + FXCounter, 1) = " "
                        End If
                    Next Loop1
                End If
        End Select
        Label3.Caption = StrTemp
    Case 7
    'Type Slide
        Select Case DirectionFX
            Case 0
                Select Case ScreenFX
                    Case 0
                        If OldFXCounter = FXCounter Then
                            FXCounter = FXCounter + 1
                            FXCounter2 = Len(Text1.Text)
                            SingleChr = Mid$(Text1.Text, FXCounter, 1)
                        End If
                        If FXCounter > TextLen Then
                            Command2_Click
                        Else
                            Mid(StrTemp, FXCounter2, 2) = "   "
                            Mid(StrTemp, FXCounter2, 1) = SingleChr
                            FXCounter2 = FXCounter2 - 1
                            If FXCounter2 < FXCounter Or SingleChr = " " Then OldFXCounter = FXCounter
                            Label3.Caption = StrTemp
                        End If
                    Case 1
                        If OldFXCounter = FXCounter Then
                            FXCounter = FXCounter + 1
                            FXCounter2 = FXCounter
                            SingleChr = Mid$(Text1.Text, FXCounter, 1)
                        End If
                        If FXCounter > TextLen Then
                            Command2_Click
                        Else
                            If FXCounter <> FXCounter2 Then Mid(StrTemp, FXCounter2, 2) = "   "
                            Mid(StrTemp, FXCounter2, 1) = SingleChr
                            FXCounter2 = FXCounter2 - 1
                            If FXCounter2 < 1 Or SingleChr = " " Then
                                OldFXCounter = FXCounter
                                Mid(StrTemp, 1, 1) = " "
                            End If
                            Label3.Caption = StrTemp
                        End If
                End Select
            Case 1
                Select Case ScreenFX
                    Case 0
                        If OldFXCounter = FXCounter Then
                            FXCounter = FXCounter + 1
                            FXCounter2 = 1
                        End If
                        If FXCounter = Len(Text1.Text) Then
                            Command2_Click
                        Else
                            SingleChr = Mid$(Text1.Text, Len(Text1.Text) - FXCounter, 1)
                            If FXCounter2 > 1 Then Mid(StrTemp, FXCounter2 - 1, 2) = "   "
                            Mid(StrTemp, FXCounter2, 1) = SingleChr
                            FXCounter2 = FXCounter2 + 1
                            If FXCounter2 > Len(Text1.Text) - FXCounter Or SingleChr = " " Then OldFXCounter = FXCounter
                            Label3.Caption = StrTemp
                        End If
                    Case 1
                        If OldFXCounter = FXCounter Then
                            FXCounter = FXCounter + 1
                            FXCounter2 = Len(Text1.Text) - FXCounter
                        End If
                        If FXCounter = Len(Text1.Text) Then
                            Command2_Click
                        Else
                            SingleChr = Mid$(Text1.Text, Len(Text1.Text) - FXCounter, 1)
                            If FXCounter2 > Len(Text1.Text) - FXCounter Then Mid(StrTemp, FXCounter2 - 1, 2) = "   "
                            Mid(StrTemp, FXCounter2, 1) = SingleChr
                            FXCounter2 = FXCounter2 + 1
                            If FXCounter2 = Len(Text1.Text) Or SingleChr = " " Then
                                OldFXCounter = FXCounter
                                Mid(StrTemp, 59, 1) = " "
                            End If
                            Label3.Caption = StrTemp
                        End If
                End Select
        End Select
    Case 8
    'Bandit
        Select Case VersionFX
            Case 0
                Select Case DirectionFX
                    Case 0
                        Select Case ScreenFX
                            Case 0
    'loop from character 32 (space) up to the ASCii value of each character. Right to Left
                                If FXCounter = 0 Then
                                    FXCounter2 = 31
                                    FXCounter = 1
                                End If
                                If FXCounter > Len(Text1.Text) Then
                                    Command2_Click
                                Else
                                    If FXCounter2 = 31 Then SingleChr = Mid$(Text1.Text, Len(Text1.Text) - (FXCounter - 1), 1)
                                    FXCounter2 = FXCounter2 + 1
                                    Mid(StrTemp, Len(Text1.Text) - (FXCounter - 1), 1) = Chr$(FXCounter2)
                                    Label3.Caption = StrTemp
                                    If FXCounter2 = Asc(SingleChr) Then
                                        FXCounter2 = 31
                                        FXCounter = FXCounter + 1
                                    End If
                                End If
                            Case 1
    'loop from the ASCii value of each character down to character 32 (space). Right to Left
                                If FXCounter = 0 Then
                                    FXCounter2 = 32
                                    FXCounter = 1
                                End If
                                If FXCounter > Len(Text1.Text) Then
                                    Command2_Click
                                Else
                                    If FXCounter2 = 32 Then FXCounter2 = Asc(Mid$(Text1.Text, Len(Text1.Text) - (FXCounter - 1), 1))
                                    If FXCounter2 > 32 Then FXCounter2 = FXCounter2 - 1
                                    Mid(StrTemp, Len(Text1.Text) - (FXCounter - 1), 1) = Chr$(FXCounter2)
                                    Label3.Caption = StrTemp
                                    If FXCounter2 = 32 Then FXCounter = FXCounter + 1
                                End If
                        End Select
                    Case 1
                        Select Case ScreenFX
                            Case 0
    'loop from character 32 (space) up to the ASCii value of each character. Left to Right
                                If FXCounter = 0 Then
                                    FXCounter2 = 31
                                    FXCounter = 1
                                End If
                                If FXCounter > Len(Text1.Text) Then
                                    Command2_Click
                                Else
                                    If FXCounter2 = 31 Then SingleChr = Mid$(Text1.Text, FXCounter, 1)
                                    FXCounter2 = FXCounter2 + 1
                                    Mid(StrTemp, FXCounter, 1) = Chr$(FXCounter2)
                                    Label3.Caption = StrTemp
                                    If FXCounter2 = Asc(SingleChr) Then
                                        FXCounter2 = 31
                                        FXCounter = FXCounter + 1
                                    End If
                                End If
                            Case 1
    'loop from the ASCii value of each character down to character 32 (space). Left to Right
                                If FXCounter = 0 Then
                                    FXCounter2 = 32
                                    FXCounter = 1
                                End If
                                If FXCounter > Len(Text1.Text) Then
                                    Command2_Click
                                Else
                                    If FXCounter2 = 32 Then FXCounter2 = Asc(Mid$(Text1.Text, FXCounter, 1))
                                    If FXCounter2 > 32 Then FXCounter2 = FXCounter2 - 1
                                    Mid(StrTemp, FXCounter, 1) = Chr$(FXCounter2)
                                    Label3.Caption = StrTemp
                                    If FXCounter2 = 32 Then FXCounter = FXCounter + 1
                                End If
                        End Select
                End Select
            Case 1
                Select Case ScreenFX
                    Case 0
                        If FXCounter = 0 Then
                            StrTemp = String$(Len(Text1.Text), " ")
                            StrMask = String$(Len(StrTemp), ".")
                            FXCounter2 = 0
                        End If
    'loop from character 32 (space) up to the ASCii value of each character. Whole String
                        If FXCounter2 = Len(Text1.Text) Then
                            Command2_Click
                        Else
                            FXCounter = FXCounter + 1
                            For Loop1 = 1 To Len(Text1.Text)
                                If Mid$(StrTemp, Loop1, 1) <> Mid$(Text1.Text, Loop1, 1) Then
                                    Mid(StrTemp, Loop1, 1) = Chr$(32 + FXCounter)
                                Else
                                    If Mid$(StrMask, Loop1, 1) = "." Then
                                        FXCounter2 = FXCounter2 + 1
                                        Mid(StrMask, Loop1, 1) = " "
                                    End If
                                End If
                            Next Loop1
                            Label3.Caption = StrTemp
                        End If
                    Case 1
                        If FXCounter = 0 Then
                            StrTemp = Text1.Text
                            StrMask = String$(Len(StrTemp), ".")
                            FXCounter2 = 0
                        End If
    'loop from the ASCii value of each character down to character 32 (space). Whole String
                        If FXCounter2 = Len(Text1.Text) Then
                            Command2_Click
                        Else
                            FXCounter = FXCounter + 1
                            For Loop1 = 1 To Len(Text1.Text)
                                If Mid$(StrTemp, Loop1, 1) <> " " Then
                                    Mid(StrTemp, Loop1, 1) = Chr$(Asc(Mid$(Text1.Text, Loop1, 1)) - FXCounter)
                                Else
                                    If Mid$(StrMask, Loop1, 1) = "." Then
                                        FXCounter2 = FXCounter2 + 1
                                        Mid(StrMask, Loop1, 1) = " "
                                    End If
                                End If
                            Next Loop1
                            Label3.Caption = StrTemp
                        End If
                End Select
        End Select
    Case 9
    'Changecase
        FXCounter = FXCounter + 1
        If FXCounter > TextLen Then
            Command2_Click
        Else
            SingleChr = Mid$(Text1.Text, FXCounter, 1)
            SingleChrType = GetCharType(SingleChr)
            If LUCaseFX = 1 Then
                If SingleChrType = 5 Then
                    'check the ASCii value of each character. if it is a lowercase letter then deduct 32 - making it uppercase
                    Mid(StrTemp, FXCounter, 1) = Chr$(Asc(SingleChr) + 32)
                Else
                    Mid(StrTemp, FXCounter, 1) = SingleChr
                End If
            Else
                If SingleChrType = 4 Then
                    'check the ASCii value of each character. if it is an uppercase letter then add 32 - making it lowercase
                    Mid(StrTemp, FXCounter, 1) = Chr$(Asc(SingleChr) - 32)
                Else
                    Mid(StrTemp, FXCounter, 1) = SingleChr
                End If
            End If
            Label3.Caption = StrTemp
        End If
    Case 10
    'Remove
    'check each character. if it is not in the chosen grouping then add it to the string
        If FXCounter = 0 Then StrTemp = ""
        FXCounter = FXCounter + 1
        If FXCounter > TextLen Then
            Command2_Click
        Else
            SingleChr = Mid$(Text1.Text, FXCounter, 1)
            SingleChrType = GetCharType(SingleChr)
            If RLChrGroupFX = 2 Then
                If LUCaseFX = 0 Then
                    If SingleChrType <> 2 Then
                        StrTemp = StrTemp + SingleChr
                    End If
                End If
                If LUCaseFX = 1 Then
                    If SingleChrType <> 4 Then
                        StrTemp = StrTemp + SingleChr
                    End If
                End If
                If LUCaseFX = 2 Then
                    If SingleChrType <> 5 Then
                        StrTemp = StrTemp + SingleChr
                    End If
                End If
            Else
                If SingleChrType <> RLChrGroupFX Then
                    StrTemp = StrTemp + SingleChr
                End If
            End If
            Label3.Caption = StrTemp
        End If
    Case 11
    'Leave
    'check each character. if it is in the chosen grouping then add it to the string
        If FXCounter = 0 Then StrTemp = ""
        FXCounter = FXCounter + 1
        If FXCounter > TextLen Then
            Command2_Click
        Else
            SingleChr = Mid$(Text1.Text, FXCounter, 1)
            SingleChrType = GetCharType(SingleChr)
            If RLChrGroupFX = 2 Then
                If LUCaseFX = 0 Then
                    If SingleChrType = 2 Then
                        StrTemp = StrTemp + SingleChr
                    End If
                End If
                If LUCaseFX = 1 Then
                    If SingleChrType = 4 Then
                        StrTemp = StrTemp + SingleChr
                    End If
                End If
                If LUCaseFX = 2 Then
                    If SingleChrType = 5 Then
                        StrTemp = StrTemp + SingleChr
                    End If
                End If
            Else
                If SingleChrType = RLChrGroupFX Then
                    StrTemp = StrTemp + SingleChr
                End If
            End If
            Label3.Caption = StrTemp
        End If
    Case 12
    'Reverse
        FXCounter = FXCounter + 1
        If FXCounter > TextLen Then
            Command2_Click
        Else
            If DirectionFX = 0 Then
                'add each character from left to right to the string
                SingleChr = Mid$(Text1.Text, FXCounter, 1)
                Mid(StrTemp, Len(Text1.Text) - (FXCounter - 1), 1) = SingleChr
            Else
                'add each character from right to left to the string
                SingleChr = Mid$(Text1.Text, TextLen - (FXCounter - 1), 1)
                Mid(StrTemp, (Len(Text1.Text) - TextLen) + FXCounter, 1) = SingleChr
            End If
            Label3.Caption = StrTemp
        End If
End Select
End Sub
