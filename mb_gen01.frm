VERSION 5.00
Begin VB.Form Generator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Message Box Generator"
   ClientHeight    =   7470
   ClientLeft      =   3600
   ClientTop       =   1035
   ClientWidth     =   4260
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "mb_gen01.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   4260
   Begin VB.CommandButton Command4 
      Caption         =   "&About"
      Height          =   345
      Left            =   3000
      TabIndex        =   30
      Top             =   6900
      Width           =   1185
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3000
      TabIndex        =   29
      Top             =   6510
      Width           =   1185
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Copy code"
      Height          =   345
      Left            =   1500
      TabIndex        =   28
      Top             =   6510
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate code"
      Default         =   -1  'True
      Height          =   345
      Left            =   150
      TabIndex        =   27
      Top             =   6510
      Width           =   1275
   End
   Begin VB.Frame Frame6 
      Caption         =   "Code"
      Height          =   1275
      Left            =   1170
      TabIndex        =   25
      Top             =   4920
      Width           =   2955
      Begin VB.TextBox Text5 
         BackColor       =   &H8000000F&
         Height          =   945
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   240
         Width           =   2745
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Title"
      Height          =   615
      Left            =   630
      TabIndex        =   23
      Top             =   4260
      Width           =   3465
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   90
         TabIndex        =   24
         Text            =   "Title of the message box"
         Top             =   210
         Width           =   3285
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Message"
      Height          =   675
      Left            =   630
      TabIndex        =   21
      Top             =   3540
      Width           =   3465
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   90
         TabIndex        =   22
         Text            =   "Your message here"
         Top             =   240
         Width           =   3285
      End
   End
   Begin VB.Frame Frame3 
      Height          =   945
      Left            =   630
      TabIndex        =   15
      Top             =   2520
      Width           =   3465
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   780
         TabIndex        =   20
         Text            =   "ContextID"
         Top             =   570
         Width           =   2115
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   780
         TabIndex        =   18
         Text            =   "HelpFile"
         Top             =   270
         Width           =   2115
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&Help enabled"
         Height          =   195
         Left            =   150
         TabIndex        =   16
         Top             =   0
         Width           =   1245
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Context:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "File:"
         Height          =   195
         Left            =   420
         TabIndex        =   17
         Top             =   300
         Width           =   300
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Buttons"
      Height          =   975
      Left            =   630
      TabIndex        =   8
      Top             =   1440
      Width           =   3465
      Begin VB.OptionButton optButton 
         Caption         =   "Abort Retry Ignore"
         Height          =   195
         Index           =   5
         Left            =   1320
         TabIndex        =   14
         Top             =   690
         Width           =   1725
      End
      Begin VB.OptionButton optButton 
         Caption         =   "Retry Cancel"
         Height          =   195
         Index           =   4
         Left            =   1320
         TabIndex        =   13
         Top             =   480
         Width           =   1275
      End
      Begin VB.OptionButton optButton 
         Caption         =   "Yes No Cancel"
         Height          =   195
         Index           =   3
         Left            =   1320
         TabIndex        =   12
         Top             =   270
         Width           =   1335
      End
      Begin VB.OptionButton optButton 
         Caption         =   "Yes No"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   11
         Top             =   690
         Width           =   915
      End
      Begin VB.OptionButton optButton 
         Caption         =   "Ok Cancel"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   10
         Top             =   480
         Width           =   1065
      End
      Begin VB.OptionButton optButton 
         Caption         =   "Ok"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   9
         Top             =   270
         Value           =   -1  'True
         Width           =   525
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Style"
      Height          =   1005
      Left            =   630
      TabIndex        =   1
      Top             =   390
      Width           =   3465
      Begin VB.OptionButton optStyle 
         Height          =   195
         Index           =   4
         Left            =   2460
         TabIndex        =   6
         Top             =   750
         Value           =   -1  'True
         Width           =   225
      End
      Begin VB.OptionButton optStyle 
         Height          =   195
         Index           =   3
         Left            =   1950
         TabIndex        =   5
         Top             =   750
         Width           =   225
      End
      Begin VB.OptionButton optStyle 
         Height          =   195
         Index           =   2
         Left            =   1350
         TabIndex        =   4
         Top             =   750
         Width           =   225
      End
      Begin VB.OptionButton optStyle 
         Height          =   195
         Index           =   1
         Left            =   780
         TabIndex        =   3
         Top             =   750
         Width           =   225
      End
      Begin VB.OptionButton optStyle 
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   750
         Width           =   225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "None"
         Height          =   195
         Left            =   2400
         TabIndex        =   7
         Top             =   390
         Width           =   375
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   3
         Left            =   1770
         Picture         =   "mb_gen01.frx":08CA
         Top             =   270
         Width           =   480
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   2
         Left            =   1200
         Picture         =   "mb_gen01.frx":0D14
         Top             =   240
         Width           =   480
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   1
         Left            =   630
         Picture         =   "mb_gen01.frx":115E
         Top             =   240
         Width           =   480
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   0
         Left            =   90
         Picture         =   "mb_gen01.frx":15A8
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "By Sam Knight (sbknight@primus.com.au)"
      Height          =   195
      Left            =   60
      TabIndex        =   31
      Top             =   7260
      Width           =   2985
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000014&
      X1              =   150
      X2              =   4200
      Y1              =   6410
      Y2              =   6410
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000010&
      X1              =   4170
      X2              =   150
      Y1              =   6390
      Y2              =   6390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Message Box Generator"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2160
      TabIndex        =   0
      Top             =   60
      Width           =   2025
   End
   Begin VB.Line Line2 
      X1              =   480
      X2              =   2130
      Y1              =   180
      Y2              =   180
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   480
      Y1              =   5130
      Y2              =   180
   End
   Begin VB.Image Image1 
      Height          =   930
      Left            =   60
      Picture         =   "mb_gen01.frx":19F2
      Stretch         =   -1  'True
      Top             =   5220
      Width           =   960
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   -90
      Shape           =   3  'Circle
      Top             =   5190
      Width           =   1185
   End
End
Attribute VB_Name = "Generator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Image3_Click()

End Sub

Private Sub Check1_Click()
'Enable or disable the controls in the help frame
Text1.Enabled = Check1.Value
Text2.Enabled = Check1.Value
End Sub

Private Sub Command1_Click()
'Generate the Message code
'Its an external sub, so call it
GenerateCode
End Sub

Private Sub Command2_Click()
'Generate the code
GenerateCode
'Now copy it
Clipboard.SetText Text5.Text
End Sub

Private Sub Command3_Click()
'Close the window
Unload Me
End Sub

Private Sub Command4_Click()
'Display an about information msgbox
MsgBox "The Message Box Code Generator" & vbCrLf & "Written by Sam Knight, June 2000." & vbCrLf & "You are free to redistribuite or modify this code to your liking." & vbCrLf & "Tell me what you think: " & vbCrLf & "sbknight@primus.com.au", vbInformation, "Message Box Generator"
End Sub

Private Sub imgIcon_Click(Index As Integer)
'Select the option equal to the image clicked
optStyle(Index).Value = True
End Sub

Private Sub Label2_Click()
optStyle(4).Value = True
End Sub

Private Sub GenerateCode()
'Generate the code for the message box

'Find out which style button is checked
If optStyle(0).Value = True Then    'Exclamation
    strAppendButton = "vbExclamation"
ElseIf optStyle(1).Value = True Then    'Question
    strAppendButton = "vbQuestion"
ElseIf optStyle(2).Value = True Then    'Cross
    strAppendButton = "vbCritical"
ElseIf optStyle(3).Value = True Then    'Info
    strAppendButton = "vbInformation"
ElseIf optStyle(4).Value = True Then    'None
    strAppendButton = ""
End If

'Find out what buttons we are using
If optButton(0).Value = True Then
    strButton = "vbOkOnly"
ElseIf optButton(1).Value = True Then
    strButton = "vbOkCancel"
ElseIf optButton(2).Value = True Then
    strButton = "vbYesNo"
ElseIf optButton(3).Value = True Then
    strButton = "vbYesNoCancel"
ElseIf optButton(4).Value = True Then
    strButton = "vbRetryCancel"
ElseIf optButton(5).Value = True Then
    strButton = "vbAbortRetryIgnore"
End If

'Find out if there is help
HelpFile = Text1.Text
HelpContext = Text2.Text

'Find out the message to be displayed
strMessage = Text3.Text

'Msgbox title
strTitle = Text4.Text

'Check to see what code to show (help or not)
If Check1.Value = 1 Then
    Text5.Text = "Dim ReturnValue as vbMsgboxResult" & vbCrLf & "ReturnValue = Msgbox(" & Chr(34) & strMessage & Chr(34) & "," & strButton & " or " & strAppendButton & "," & Chr(34) & strTitle & Chr(34) & "," & Chr(34) & HelpFile & Chr(34) & "," & Chr(34) & HelpContext & Chr(34) & ")" & vbCrLf & "'Add your event handlers here" & vbCrLf & "Select case ReturnValue" & vbCrLf & "   'add select statements here" & vbCrLf & "End Select"
Else
    Text5.Text = "Dim ReturnValue as vbMsgboxResult" & vbCrLf & "ReturnValue = Msgbox(" & Chr(34) & strMessage & Chr(34) & "," & strButton & " or " & strAppendButton & "," & Chr(34) & strTitle & Chr(34) & ")" & vbCrLf & "'Add your event handlers here" & vbCrLf & "Select case ReturnValue" & vbCrLf & "   'add select statements here" & vbCrLf & "End Select"
End If
End Sub

