VERSION 5.00
Begin VB.Form frm_bg 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   6690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8850
   ClipControls    =   0   'False
   Icon            =   "bg_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "bg_frm.frx":0442
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frm_bg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'feel free to change any part of this project
'just give credit where credit is due
'feeback or comments
'email:  u_tech@ excite.com

Private Sub Form_Load()
Me.WindowState = 2
Form1.Show
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.MousePointer = 99
End Sub
