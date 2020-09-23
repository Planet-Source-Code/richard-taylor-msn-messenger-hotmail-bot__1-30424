VERSION 5.00
Begin VB.UserControl MSN 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3495
   InvisibleAtRuntime=   -1  'True
   Picture         =   "MSN.ctx":0000
   ScaleHeight     =   3030
   ScaleWidth      =   3495
   ToolboxBitmap   =   "MSN.ctx":0102
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "New MSN Only"
      BeginProperty Font 
         Name            =   "Niagara Engraved"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "MSN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Created only by Richard Taylor Apart from the
'Change background function, as it says in the
'section who made it - kegham.
'This control is freeware to use compiled
'but to use any source code seperate you
'must only upload the control and link to
'the control in my section
'
'If  you use the source code without my
'autority you will be reported to
'PSCode.com and your code will be deleted
'This code was uploaded 1:42pm GMT TIME
'13/1/02
'
'Happy coding
'
'
'Please vote on the site

Option Explicit
Public Function ShowAudioTuning()
'Show the audio tuning wizard dialogue
messengerapi.MediaWizard 0 'hwndparent must be 0 everytime in MSN
End Function
Private Sub UserControl_Resize()
'Keep the control the same size everytim
UserControl.Height = 3030
UserControl.Width = 3495
End Sub
Public Function NickName()
'Get your nickname that is used in conversations
NickName = messengerapi.MyFriendlyName
End Function
Public Function GetStatus()
'Gets your status and places it as the function
Dim strStatus As String
If messengerapi.MyStatus = MISTATUS_ONLINE Then
strStatus = "Online"
End If
If messengerapi.MyStatus = MISTATUS_AWAY Then
strStatus = "Away"
End If
If messengerapi.MyStatus = MISTATUS_BE_RIGHT_BACK Then
strStatus = "Be Right Back"
End If
If messengerapi.MyStatus = MISTATUS_BUSY Then
strStatus = "Busy"
End If
If messengerapi.MyStatus = MISTATUS_IDLE Then
strStatus = "Idle"
End If
If messengerapi.MyStatus = MISTATUS_INVISIBLE Then
strStatus = "Invisible"
End If
If messengerapi.MyStatus = MISTATUS_LOCAL_CONNECTING_TO_SERVER Then
strStatus = "Connecting to MSN"
End If
If messengerapi.MyStatus = MISTATUS_ON_THE_PHONE Then
strStatus = "Your on the phone"
End If
If messengerapi.MyStatus = MISTATUS_OUT_TO_LUNCH Then
strStatus = "Out to lunch"
End If
If messengerapi.MyStatus = MISTATUS_LOCAL_DISCONNECTING_FROM_SERVER Then
strStatus = "Disconnected"
End If
If messengerapi.MyStatus = MISTATUS_LOCAL_FINDING_SERVER Then
strStatus = "Finding the server"
End If
If messengerapi.MyStatus = MISTATUS_LOCAL_SYNCHRONIZING_WITH_SERVER Then
strStatus = "Synchronizing with the server"
End If
If messengerapi.MyStatus = MISTATUS_UNKNOWN Then
strStatus = "Unknown"
End If
GetStatus = strStatus
End Function
Public Function SignIn()
'Take you to the sign in screen
messengerapi.SignIn 0, "", "" 'The first set of speechmarks is the username and the second is the password
End Function
Public Function AutoSignIn()
'Uses the auto sign in function
messengerapi.AutoSignIn
End Function
Public Function ChangeBackGroundGIF(Filename As String)
'Changes the background in the real MSN messenger
'My good friend Keghams code for this function
'Kegham_D@Hotmail.com
FileCopy Filename, "C:\Program Files\Messenger\lvback.gif"
End Function
Public Function EmailAddress()
'Gets your email address
EmailAddress = messengerapi.MySigninName
End Function
Public Function ViewProfile(Email As String)
'Views the profile of anyone
messengerapi.ViewProfile Email
End Function
Public Function SendEmail(Email As String)
'Sends an email to anyone on the internet
messengerapi.SendMail Email
End Function
Public Function SendSingleMessage(Email As String, Message As String)
'Sends a message and then instantly quits
messengerapi.InstantMessage Email
SendKeys Message
SendKeys "{Enter}"
SendKeys "{Esc}"
End Function
Public Function UnReadMailCount()
'Gets the unread email count in your inbox
UnReadMailCount = messengerapi.UnreadEmailCount(MUAFOLDER_INBOX)
End Function
Public Function GotoInbox()
'Takes your web browser to your inbox
messengerapi.OpenInbox
End Function
Public Function SignOut()
'Signs you out of hotmail
messengerapi.SignOut
End Function
Public Function ShowOptionsGeneral()
'Shows the options screen
messengerapi.OptionsPages 0, MOPT_GENERAL_PAGE
End Function
Public Function ShowOptionsAccounts()
'Shows the accounts option screen
messengerapi.OptionsPages 0, MOPT_ACCOUNTS_PAGE
End Function
Public Function ShowOptionsConnections()
'Shows the connections screen
messengerapi.OptionsPages 0, MOPT_CONNECTION_PAGE
End Function
Public Function ShowOptionsExchange()
messengerapi.OptionsPages 0, MOPT_EXCHANGE_PAGE
End Function
Public Function ShowOptionsPhone()
messengerapi.OptionsPages 0, MOPT_PHONE_PAGE
End Function
Public Function ShowOptionsPreferences()
messengerapi.OptionsPages 0, MOPT_PREFERENCES_PAGE
End Function
Public Function ShowOptionsPrivacy()
messengerapi.OptionsPages 0, MOPT_PRIVACY_PAGE
End Function
Public Function ShowOptionsServices()
messengerapi.OptionsPages 0, MOPT_SERVICES_PAGE
End Function
Public Function SetStatusOnline()
'set status to online
On Error Resume Next
messengerapi.MyStatus = MISTATUS_ONLINE
End Function
Public Function SetStatusBusy()
On Error Resume Next
messengerapi.MyStatus = MISTATUS_BUSY
End Function
Public Function SetStatusBeRightBack()
On Error Resume Next
messengerapi.MyStatus = MISTATUS_BE_RIGHT_BACK
End Function
Public Function SetStatusAway()
On Error Resume Next
messengerapi.MyStatus = MISTATUS_AWAY
End Function
Public Function SetStatusPhone()
On Error Resume Next
messengerapi.MyStatus = MISTATUS_ON_THE_PHONE
End Function
Public Function SetStatusLunch()
On Error Resume Next
messengerapi.MyStatus = MISTATUS_OUT_TO_LUNCH
End Function
Public Function SetStatusInvisible()
On Error Resume Next
messengerapi.MyStatus = MISTATUS_INVISIBLE
End Function
Public Function SendFile(Filename, Email As String)
'Sends a file to an email address via, msn
messengerapi.SendFile Email, Filename
End Function
Public Function StartConversation(Email As String)
'Starts a conversation window
messengerapi.InstantMessage Email
End Function
Public Function SendMessageStartConversation(Email, Message As String)
'Starts a conversation window AND sends one message
messengerapi.InstantMessage Email
SendKeys Message
End Function
Public Function AddContact(Email As String)
'Adds a contact to your list
messengerapi.AddContact 0, Email
End Function
Public Function RecievedFilesFolder()
'Gets the recieved files folder
RecievedFilesFolder = messengerapi.ReceiveFileDirectory
End Function
Public Function CreateGroup(NewName, Service As String)
'Creates a new group
messengerapi.CreateGroup NewName, Service
End Function
Public Function FindContact(Firstname, Surname) As String
'Finds a contact anywhere in the world
messengerapi.FindContact 0, Firstname, Surname, "", "", ""
End Function
Public Function GetContact(Email, ServiceID) As String
'Gets a contact name
messengerapi.GetContact Email, ServiceID
End Function
Public Function SetGroupOrderByGROUPS()
'Set the group orders
messengerapi.ContactsSortOrder = MUASORT_GROUPS
End Function
Public Function SetGroupOrderByONLINE_OFFLINE()
messengerapi.ContactsSortOrder = MUASORT_ONOFFLINE
End Function
Public Function SetNickName(NewName As String)
'Sets the nickname
messengerapi.OptionsPages 0, MOPT_GENERAL_PAGE
'Send the keys below
SendKeys NewName
SendKeys "{Enter}"
End Function

