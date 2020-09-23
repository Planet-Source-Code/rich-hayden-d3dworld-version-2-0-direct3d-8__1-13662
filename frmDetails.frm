VERSION 5.00
Begin VB.Form frmDetails 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Details"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      TabIndex        =   19
      Text            =   "0.2"
      Top             =   7080
      Width           =   375
   End
   Begin VB.CheckBox Check2 
      Caption         =   "House Light?"
      Height          =   255
      Left            =   2280
      TabIndex        =   16
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Time Of Day/Lighting"
      Height          =   2895
      Left            =   2160
      TabIndex        =   8
      Top             =   4080
      Width           =   2055
      Begin VB.CheckBox Check5 
         Caption         =   "Eerie green glow on church?"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   2400
         Width           =   1815
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Search Light?"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Sun/Moon?"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Early Morning"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Evening"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Night"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Dawn"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Midday"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   480
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Dusk"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   10000
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fill Mode"
      Height          =   2895
      Left            =   0
      TabIndex        =   7
      Top             =   4080
      Width           =   2055
      Begin VB.OptionButton Option1 
         Caption         =   "Points"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Wireframe"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Fill"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Search Light Intensity:"
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Label Label2 
      Height          =   1935
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   4095
   End
   Begin VB.Label lblFps 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   4095
   End
   Begin VB.Label Label1 
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "frmDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    Command1.SetFocus
End Sub

Private Sub Check2_Click()
    Command1.SetFocus
End Sub

Private Sub Check3_Click()
    Command1.SetFocus
End Sub

Private Sub Check4_Click()
    Command1.SetFocus
End Sub

Private Sub Check5_Click()
    Command1.SetFocus
End Sub

Private Sub Form_Load()
    Label2.Caption = "Controls: " & Chr(13) & "TURN LEFT=LEFT ARROW" & Chr(13) & "TURN RIGHT=RIGHT ARROW" & Chr(13) & "MOVE FORWARD=UP ARROW" & Chr(13) & "MOVE BACKWARD=DOWN ARROW" & Chr(13) & "JUMP=SPACE BAR" & Chr(13) & "RUN=HOLD DOWN SHIFT" & Chr(13) & "CROUCH=HOLD DOWN CTRL"
End Sub

Private Sub Option1_Click(Index As Integer)
    On Error Resume Next
    If Option1(0).Value = True Then
        lngFillMode = D3DFILL_SOLID
    ElseIf Option1(1).Value = True Then
        lngFillMode = D3DFILL_WIREFRAME
    ElseIf Option1(2).Value = True Then
        lngFillMode = D3DFILL_POINT
    End If
    
    Command1.SetFocus
End Sub

Private Sub Option2_Click(Index As Integer)
    On Error Resume Next
    If Option2(0).Value = True Then
        lngLightType = DAWN_LIGHT
    ElseIf Option2(1).Value = True Then
        lngLightType = MIDDAY_LIGHT
    ElseIf Option2(2).Value = True Then
        lngLightType = DUSK_LIGHT
    ElseIf Option2(3).Value = True Then
        lngLightType = EVENING_LIGHT
    ElseIf Option2(4).Value = True Then
        lngLightType = NIGHT_LIGHT
    ElseIf Option2(5).Value = True Then
        lngLightType = EARLYMORNING_LIGHT
    End If
    
    Command1.SetFocus
End Sub

Private Sub Text1_Change()
    If IsNumeric(Text1.Text) Then
        searchLightIntensity = CSng(Text1.Text)
    End If
End Sub

Private Sub Text1_LostFocus()
    If Not IsNumeric(Text1.Text) Then
        Text1.Text = "0.2"
    End If
    searchLightIntensity = 0.2
End Sub

Private Sub Text2_Change()
    If Not IsNumeric(Text2.Text) Then
        Text2.Text = "0.2"
    End If
    torchLightIntensity = 0.2
End Sub

Private Sub Text2_LostFocus()
    If Not IsNumeric(Text2.Text) Then
        Text2.Text = "0.2"
    End If
    torchLightIntensity = 0.2
End Sub
