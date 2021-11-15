VERSION 5.00
Begin VB.Form frmGuildLeader 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "$433"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   6165
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
   Icon            =   "frmGuildLeader.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Cancel          =   -1  'True
      Caption         =   "$25"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3090
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   6030
      Width           =   2955
   End
   Begin VB.CommandButton Command6 
      Caption         =   "$50"
      Height          =   495
      Left            =   3090
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   4650
      Width           =   2955
   End
   Begin VB.CommandButton Command5 
      Caption         =   "$49"
      Height          =   495
      Left            =   3090
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   4110
      Width           =   2955
   End
   Begin VB.Frame Frame3 
      Caption         =   "$34"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   90
      TabIndex        =   9
      Top             =   90
      Width           =   2895
      Begin VB.ListBox guildslist 
         Height          =   1425
         ItemData        =   "frmGuildLeader.frx":000C
         Left            =   120
         List            =   "frmGuildLeader.frx":000E
         TabIndex        =   11
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton Command4 
         Caption         =   "$35"
         Height          =   375
         Left            =   120
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   1800
         Width           =   2655
      End
   End
   Begin VB.Frame txtnews 
      Caption         =   "$52"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   90
      TabIndex        =   6
      Top             =   2460
      Width           =   5955
      Begin VB.CommandButton Command3 
         Caption         =   "$48"
         Height          =   375
         Left            =   120
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   1080
         Width           =   5715
      End
      Begin VB.TextBox txtguildnews 
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   5715
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "$53"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   3060
      TabIndex        =   3
      Top             =   90
      Width           =   2985
      Begin VB.CommandButton Command2 
         Caption         =   "$35"
         Height          =   375
         Left            =   120
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   1800
         Width           =   2745
      End
      Begin VB.ListBox members 
         Height          =   1425
         ItemData        =   "frmGuildLeader.frx":0010
         Left            =   120
         List            =   "frmGuildLeader.frx":0012
         TabIndex        =   4
         Top             =   240
         Width           =   2745
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "$51"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   90
      TabIndex        =   0
      Top             =   4110
      Width           =   2895
      Begin VB.CommandButton Command1 
         Caption         =   "$35"
         Height          =   375
         Left            =   120
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   1560
         Width           =   2655
      End
      Begin VB.ListBox solicitudes 
         Height          =   1230
         ItemData        =   "frmGuildLeader.frx":0014
         Left            =   120
         List            =   "frmGuildLeader.frx":0016
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Miembros 
         Alignment       =   2  'Center
         Caption         =   "El clan cuenta con x miembros"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmGuildLeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmGuildLeader - ImperiumAO - v1.4.5 R5
'*****************************************************************
'Respective portions copyrighted by contributors listed below.
'
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation version 2.1 of
'the License
'
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'Lesser General Public License for more details.
'
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'*****************************************************************

'*****************************************************************
'Pablo Ignacio Márquez (morgolock@speedy.com.ar)
'   - First Relase
'*****************************************************************

Private Sub Command1_Click()

frmCharInfo.frmsolicitudes = True
Call ClientTCP.Send_Data(Guild_Char_Info_Cl, solicitudes.List(solicitudes.ListIndex))

End Sub

Private Sub Command2_Click()

frmCharInfo.frmmiembros = True
Call ClientTCP.Send_Data(Guild_Char_Info_Cl, members.List(members.ListIndex))

End Sub

Private Sub Command3_Click()

Dim k$
k$ = Replace(txtguildnews, vbCrLf, "º")
Call ClientTCP.Send_Data(Guild_News_Set, k$)

End Sub

Private Sub Command4_Click()

frmGuildBrief.EsLeader = True
Call ClientTCP.Send_Data(Guild_Details_Request, guildslist.List(guildslist.ListIndex))

End Sub

Private Sub Command5_Click()

frmGuildDetails.framAlign.Visible = False
Call frmGuildDetails.Show(vbModeless, frmMain)
Unload Me

End Sub

Private Sub Command6_Click()
Call frmGuildURL.Show(vbModeless, frmGuildLeader)
End Sub

Private Sub Command8_Click()
Unload Me
End Sub


Public Sub ParseLeaderInfo(ByVal Data As String)

If Me.Visible Then Exit Sub

Dim r%, t%

r% = Val(General_Field_Read(1, Data, "¬"))

For t% = 1 To r%
    guildslist.AddItem General_Field_Read(1 + t%, Data, "¬")
Next t%

r% = Val(General_Field_Read(t% + 1, Data, "¬"))
Miembros.Caption = IIf(r% > 1, Locale_GUI_Frase(168) & " " & r% & " " & Locale_GUI_Frase(169) & ".", "El clan cuenta con un miembro.")

Dim k%

For k% = 1 To r%
    members.AddItem General_Field_Read(t% + 1 + k%, Data, "¬")
Next k%

txtguildnews = Replace(General_Field_Read(t% + k% + 1, Data, "¬"), "º", vbCrLf)

t% = t% + k% + 2

r% = Val(General_Field_Read(t%, Data, "¬"))

For k% = 1 To r%
    solicitudes.AddItem General_Field_Read(t% + k%, Data, "¬")
Next k%

Me.Show vbModeless, frmMain

End Sub

Private Sub Form_Deactivate()
On Error Resume Next
If Me.Visible And Not frmGuildURL.Visible _
And Not frmGuildBrief.Visible _
And Not frmCharInfo.Visible _
Then Me.SetFocus
End Sub

Private Sub Form_Load()
Call FormParser.Parse_Form(Me)
End Sub
