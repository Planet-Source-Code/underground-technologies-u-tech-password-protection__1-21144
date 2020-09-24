VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form2"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5340
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   5415
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2D544&
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form2.frx":0000
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2D544&
      Height          =   4335
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'feel free to change any part of this project
'just give credit where credit is due
'feeback or comments
'email:  u_tech@ excite.com

Private Sub Form_Activate()
CenterForm Me
Call CycleText("I made this project cause I was trying to figure out a way to password protect something in one of my apps where nobody could figure out the password by finding an INI, TXT or and kind of trace of the password. I have not seen anything like on PSC so I made up my own. This is the first project I have released to PSC. If I get a good responce from this one then I will release some more.", Label1)
Call CycleText("comments or feedback mailto:      u_tech@excite.com", Label2)
Label3.Enabled = True

End Sub

Private Sub Form_Load()
'CenterForm Me
End Sub

Private Sub Form_Resize()
If Me.WindowState = 1 Then Me.WindowState = 0: Exit Sub
If Me.WindowState = 2 Then Me.WindowState = 0: Exit Sub
'Me.Height = Picture1.Height
'Me.Width = Picture1.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub CycleText(txt As String, lbl1 As Label)
Dim jay
Dim startTime
lbl1.Caption = ""
For jay = 1 To Len(txt)
    lbl1.Caption = Mid$(txt, 1, jay)
    startTime = Timer
    Do While Timer - startTime < Val("0.00001")
    DoEvents
    Loop
Next jay
End Sub

Private Sub Label3_Click()
Unload Me
End Sub
