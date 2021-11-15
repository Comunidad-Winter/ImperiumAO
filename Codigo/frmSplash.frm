VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9675
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   480
      Top             =   360
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   690
      Left            =   5640
      Picture         =   "frmSplash.frx":0000
      ScaleHeight     =   630
      ScaleWidth      =   735
      TabIndex        =   0
      Top             =   2280
      Width           =   795
   End
   Begin VB.Label lblAvanceDesc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H008080FF&
      Height          =   195
      Left            =   1200
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblAvance 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H008080FF&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   3840
      Left            =   6120
      Picture         =   "frmSplash.frx":056C
      Top             =   0
      Width           =   3840
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ImperiumAO TileEngine "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   435
      Left            =   480
      TabIndex        =   1
      Top             =   1560
      Width           =   4200
   End
   Begin VB.Image Image2 
      Height          =   1155
      Left            =   0
      Picture         =   "frmSplash.frx":109AE
      Top             =   0
      Width           =   6360
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   120
      Picture         =   "frmSplash.frx":28888
      Top             =   2160
      Width           =   5700
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Picture1_DblClick()
Timer1.Enabled = Not Timer1.Enabled
End Sub

Private Sub Timer1_Timer()
Me.Tag = "0"
Timer1.Enabled = False
Unload Me
End Sub
