VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Simple Virtual Keyboard"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4065
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      ScaleHeight     =   705
      ScaleWidth      =   3105
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   3135
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ENTER"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   27
         Left            =   2160
         TabIndex        =   29
         Top             =   480
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SPACE"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   26
         Left            =   0
         TabIndex        =   28
         Top             =   480
         Width           =   2145
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Z"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   25
         Left            =   2880
         TabIndex        =   27
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Y"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   24
         Left            =   2640
         TabIndex        =   26
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "X"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   23
         Left            =   2400
         TabIndex        =   25
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "W"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   22
         Left            =   2160
         TabIndex        =   24
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "V"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   21
         Left            =   1920
         TabIndex        =   23
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "U"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   20
         Left            =   1680
         TabIndex        =   22
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "T"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   19
         Left            =   1440
         TabIndex        =   21
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   18
         Left            =   1200
         TabIndex        =   20
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   17
         Left            =   960
         TabIndex        =   19
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Q"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   16
         Left            =   720
         TabIndex        =   18
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "P"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   15
         Left            =   480
         TabIndex        =   17
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "O"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   14
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "N"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   13
         Left            =   0
         TabIndex        =   15
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   12
         Left            =   2880
         TabIndex        =   14
         Top             =   0
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "L"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   11
         Left            =   2640
         TabIndex        =   13
         Top             =   0
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "K"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   10
         Left            =   2400
         TabIndex        =   12
         Top             =   0
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "J"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   9
         Left            =   2160
         TabIndex        =   11
         Top             =   0
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "I"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   8
         Left            =   1920
         TabIndex        =   10
         Top             =   0
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "H"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   7
         Left            =   1680
         TabIndex        =   9
         Top             =   0
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "G"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   6
         Left            =   1440
         TabIndex        =   8
         Top             =   0
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   5
         Left            =   1200
         TabIndex        =   7
         Top             =   0
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "E"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   4
         Left            =   960
         TabIndex        =   6
         Top             =   0
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "D"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   3
         Left            =   720
         TabIndex        =   5
         Top             =   0
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   2
         Left            =   480
         TabIndex        =   4
         Top             =   0
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "B"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   0
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   225
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1065
      Left            =   0
      Picture         =   "Form1.frx":000C
      ScaleHeight     =   1005
      ScaleWidth      =   4005
      TabIndex        =   31
      Top             =   2910
      Width           =   4065
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "stormdev@golden.net"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   2280
      TabIndex        =   33
      Top             =   2520
      Width           =   1560
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":8A7E
      Height          =   735
      Left            =   120
      TabIndex        =   32
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":8B08
      Height          =   855
      Left            =   120
      TabIndex        =   30
      Top             =   720
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3360
      Picture         =   "Form1.frx":8BBE
      Top             =   45
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This code is copyright(c) 2000, 2001 Stormdev Software Development.
' You are hereby granted rights to use/modify this code as you see fit,
' for commercial or personal use. The only stipulation is that I
' ask for some feedback regarding the code contained herein.
'
' Send feedback to: stormdev@golden.net
'      Code Author: Jonathan Roach
'  Purpose of Code: Demonstrate a virtual keyboard
'    Level of Code: Beginner
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Image1_Click()
' The keyboard icon which toggles display of the
' virtual keyboard.
Picture1.Visible = Not Picture1.Visible
End Sub

Private Sub Label1_Click(Index As Integer)
' These are the label controls that represent each key
' on the virtual keyboard.
Select Case Index
    Case 26 ' Space was clicked
        Text1.SetFocus
        SendKeys " "
    Case 27 ' Enter key was clicked
        Text1.SetFocus
        SendKeys "~"
        Picture1.Visible = False
    Case Else
        Text1.SetFocus
        SendKeys Label1(Index).Caption
End Select
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
' The following line simply disables the
' enter key beep when the enter key is
' pressed in a textbox.
If KeyAscii = 13 Then KeyAscii = 0
End Sub
