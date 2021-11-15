VERSION 5.00
Begin VB.Form frmGuildAdm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "$436"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   4065
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
   Icon            =   "frmGuildAdm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "$25"
      Height          =   375
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   3690
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "$36"
      Height          =   375
      Left            =   1080
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3690
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "$35"
      Height          =   375
      Left            =   2640
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3690
      Width           =   1335
   End
   Begin VB.Frame Frame1 
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
      Height          =   3435
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.ListBox GuildsList 
         Height          =   2790
         ItemData        =   "frmGuildAdm.frx":000C
         Left            =   240
         List            =   "frmGuildAdm.frx":000E
         TabIndex        =   1
         Top             =   360
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmGuildAdm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmGuildAdm - ImperiumAO - v1.4.5 R5
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

Private cargado As Boolean

Private Sub Command1_Click()

Call ClientTCP.Send_Data(Guild_Details_Request, guildslist.List(guildslist.ListIndex))

End Sub

Private Sub Command2_Click()
If guildslist.ListIndex = -1 Then Exit Sub
Call frmGuildSol.RecieveSolicitud(guildslist.List(guildslist.ListIndex))
Call frmGuildSol.Show(vbModeless, frmGuildBrief)
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Public Sub ParseGuildList(ByVal rData As String)

Dim j As Integer, k As Integer

k = CInt(General_Field_Read(1, rData, ","))

For j = 1 To k
    guildslist.AddItem General_Field_Read(1 + j, rData, ",")
Next j

Me.Show vbModeless, frmMain

End Sub

Private Sub Form_Deactivate()

If Not frmGuildBrief.Visible And Not frmGuildSol.Visible And Me.Visible Then Me.SetFocus

End Sub

Private Sub Form_Load()
Call FormParser.Parse_Form(Me)
End Sub
