VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Enter PassWord"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2445
   BeginProperty Font 
      Name            =   "Tempus Sans ITC"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   2445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   960
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   0
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "The Password is          text1"
      Top             =   0
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Try PassWord"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      MaskColor       =   &H000000C0&
      TabIndex        =   0
      Tag             =   "0"
      ToolTipText     =   "The Password is          text1"
      Top             =   600
      UseMaskColor    =   -1  'True
      Width           =   2445
   End
   Begin VB.CheckBox chkSound 
      Caption         =   "Sounds: On"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "The Password is          text1"
      Top             =   360
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   960
      TabIndex        =   3
      Text            =   "stop"
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'feel free to change any part of this project
'just give credit where credit is due
'the password is        text1
'feeback or comments
'email:  u_tech@ excite.com


Public Sub PW()
'check to see if the sounds are on or off and plays the sound if they are on
If chkSound.Value = 1 Then
PlayWav (App.Path + "\type.wav")
Else
End If
End Sub

Public Sub PW_rite()
'check to see if the sounds are on or off and plays the sound if they are on

If chkSound.Value = 1 Then
PlayWav (App.Path + "\pw-rite.wav")
Else
End If
End Sub

Public Sub PW_wrong()
'check to see if the sounds are on or off and plays the sound if they are on

If chkSound.Value = 1 Then
PlayWav (App.Path + "\denied.wav")
Else
End If
End Sub



Private Sub chkSound_Click()
If chkSound.Caption = "Sounds: On" Then
chkSound.Caption = "Sounds: Off"
chkSound.ForeColor = &H6130E0 'change the checkbox forecolor to red
Else
chkSound.Caption = "Sounds: On"
chkSound.ForeColor = &HFF00& 'change the checkbox forecolor to green
End If
End Sub
Private Sub Command1_Click()
'this was made up in about 20 minutes so it may be a little buggy
Dim t6, t1, t2, t3, t4, t5 As String
Dim daP As String
Text1.PasswordChar = "*" ' this sets the textbox passwordcharactor to  *
        t1 = "t" 'set the 1st letter of your password
        t2 = "e" 'set the 2nd letter of your password
        t3 = "x" 'set the 3rd letter of your password
        t4 = "t" 'set the 4th letter of your password
        t5 = "1" 'set the 5th letter of your password
        t6 = t1 + t2 + t3 + t4 + t5 'put all the pw letters together
        
' write the password to the registry
Call SetStringValue(HKEY_LOCAL_MACHINE, "Software\u-tech\system\u-tech\pen", "Currency", "" & t6)

' gets  the password from the registry
daP = getstring(HKEY_LOCAL_MACHINE, "Software\u-tech\system\u-tech\pen", "Currency")

If LCase(Text1.Text) = LCase(daP) Then 'check the pass from the registry to see if it is correct

Text1.PasswordChar = "" '  sets the textbox passwordcharactor to  nothing
Call COLOR_GRN 'the password was correct so change the color to green
Text1.FontBold = True 'make the font bold
Text1 = "PW CORRECT" 'show user that the pw was correct

'delete the password from the registry so unwanted people wont find it
Call DeleteValue(HKEY_LOCAL_MACHINE, "Software\u-tech\system\u-tech\pen", "Currency")
'play sound letting user know pw was correct
Call PW_rite
pause 1 'pause for 1 second
Unload frm_bg 'unload the background form
Form2.Show 'this would be where you put the password protected area of your app
Text2.Text = "" 'make the form unloadable
Unload Me 'unload the form
Else ' if the pw was wrong
'delete the password from the registry
Call DeleteValue(HKEY_LOCAL_MACHINE, "Software\u-tech\system\u-tech\pen", "Currency")
Call COLOR_RED 'change the forecolor of the textbox to red
Text1.FontBold = True 'make the font bold
Text1.PasswordChar = "" 'set the passwordchactor to nothing
Text2.Text = "" 'make the form unloadable
Text1 = "PW WRONG" 'let the user know that the pw was wrong
Timer1.Enabled = True 'start the ending app procedure by turning on the timer
End If ' duhhhhh.......
End Sub


Private Sub Form_Activate()
Text1.SetFocus 'puts the mouse cursor in the textbox
End Sub

Sub COLOR_GRN()
Text1.ForeColor = &HC000&
End Sub

Sub COLOR_RED()
Text1.ForeColor = &H6130E0
End Sub

Private Sub Form_Load()
Call CenterForm(Me)
Call BeTop(Me, 1)
End Sub

Private Sub Form_QueryUnload(p1 As Integer, p2 As Integer)
' this make it so the user can't unload the form
p1 = True 'cancel
If Text2.Text = "stop" Then p1 = True: Exit Sub 'check to see if the user has pressed the decode button
p1 = False 'unloadmode
Unload Me 'unload the form
Exit Sub ' duhhh.....
End Sub

Private Sub Form_Resize()
Me.Height = 1240 'make the form stay the same height
Me.Width = 2515 'make the form stay the same width
End Sub

Private Sub Text1_Change()
Call PW 'play the typing .wav
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
' this checks to see if the user presses the enter key to try the password
If KeyAscii = 13 Then
Command1_Click ' if the key was enter click the decode button
End If
End Sub

Private Sub Timer1_Timer()
Call PW_wrong 'plays the wrong password .wav
Timer1.Enabled = False 'turns the timer off
pause 2 'pause for 2 seconds
End 'close the app
End Sub
