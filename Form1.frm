VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CD Identification Number"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3435
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   3435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get CD ID"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CD ID:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   405
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

' If The CD Drive Isn't Available, Then We Better Quit!

mciSendString "close all", 0, 0, 0
If (SendMCIString("open cdaudio alias cd69 wait shareable", True) = False) Then

End
End If

' If The CD Drive IS Ready, Then Let's Go On...

SendMCIString "set cd69 time format tmsf wait", True
readcdtoc
Text1.Text = cddbdiscid(totaltr)

End Sub

Private Sub Command2_Click()

Unload Me

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

' This Piece Of Code, Close the CD Audio.
' So That Other Programs, Can Get Access to the CD Audio.

mciSendString "close all", 0, 0, hWnd

End Sub

