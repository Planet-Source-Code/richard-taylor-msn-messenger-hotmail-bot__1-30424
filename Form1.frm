VERSION 5.00
Object = "{ED7CC03B-082B-11D6-B02C-00A0242DDE13}#2.0#0"; "MSN Bot.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MSN Bot v3.0 Due to Popular Demand... Back in OCX!"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   8
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Sign Out"
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Options"
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Media"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3000
      Top             =   720
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change name"
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   630
      Left            =   820
      TabIndex        =   1
      Top             =   2640
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1111
      _Version        =   393216
      Max             =   6
      TextPosition    =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin Project1.MSN MSN 
      Left            =   -2760
      Top             =   2040
      _ExtentX        =   6165
      _ExtentY        =   5345
   End
   Begin VB.Label lblName 
      Caption         =   "Friendly name"
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   480
      Width           =   2775
   End
   Begin VB.Image Image7 
      Height          =   780
      Left            =   5520
      Picture         =   "Form1.frx":0000
      Top             =   1695
      Width           =   210
   End
   Begin VB.Image Image6 
      Height          =   1320
      Left            =   4755
      Picture         =   "Form1.frx":0782
      Top             =   1185
      Width           =   225
   End
   Begin VB.Image Image5 
      Height          =   1320
      Left            =   3915
      Picture         =   "Form1.frx":1844
      Top             =   1185
      Width           =   240
   End
   Begin VB.Image Image4 
      Height          =   525
      Left            =   3180
      Picture         =   "Form1.frx":2906
      Top             =   1980
      Width           =   225
   End
   Begin VB.Image Image3 
      Height          =   1335
      Left            =   2385
      Picture         =   "Form1.frx":2FD8
      Top             =   1200
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   435
      Left            =   1635
      Picture         =   "Form1.frx":40CA
      Top             =   2040
      Width           =   225
   End
   Begin VB.Image Image1 
      Height          =   690
      Left            =   840
      Picture         =   "Form1.frx":467C
      Top             =   1845
      Width           =   270
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End
End Sub
Private Sub Command2_Click()
On Error GoTo Error:
MSN.SetNickName txtName.Text
Timer1_Timer
txtName.Text = ""
txtName.SetFocus
Exit Sub
Error:
MsgBox "Please make sure you do not use brackets in your name", vbCritical, "Error"
End Sub

Private Sub Command3_Click()
MSN.ShowAudioTuning
End Sub

Private Sub Command4_Click()
MSN.ShowOptionsGeneral
End Sub

Private Sub Command5_Click()
MSN.SignOut
End Sub

Private Sub Command6_Click()
Me.WindowState = 1
End Sub

Private Sub Form_Load()
Timer1_Timer
End Sub

Private Sub Slider1_Change()
If Slider1.Value = 0 Then
MSN.SetStatusOnline
ElseIf Slider1.Value = 1 Then
MSN.SetStatusBusy
ElseIf Slider1.Value = 2 Then
MSN.SetStatusBeRightBack
ElseIf Slider1.Value = 3 Then
MSN.SetStatusAway
ElseIf Slider1.Value = 4 Then
MSN.SetStatusPhone
ElseIf Slider1.Value = 5 Then
MSN.SetStatusLunch
ElseIf Slider1.Value = 6 Then
MSN.SetStatusInvisible
End If
End Sub

Private Sub Slider1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Timer1.Enabled = False
End Sub

Private Sub Slider1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
lblName.Caption = MSN.NickName
If MSN.GetStatus = "Online" Then
Slider1.Value = 0
ElseIf MSN.GetStatus = "Away" Then
Slider1.Value = 3
ElseIf MSN.GetStatus = "Busy" Then
Slider1.Value = 1
ElseIf MSN.GetStatus = "Be Right Back" Then
Slider1.Value = 2
ElseIf MSN.GetStatus = "Your on the phone" Then
Slider1.Value = 4
ElseIf MSN.GetStatus = "Out to lunch" Then
Slider1.Value = 5
ElseIf MSN.GetStatus = "Invisible" Then
Slider1.Value = 6
End If
End Sub
