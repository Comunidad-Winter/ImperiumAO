VERSION 5.00
Begin VB.Form frmGuildNews 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Novedades del clan"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   4845
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkShow 
      Caption         =   "Mostrar la próxima vez"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   5880
      Width           =   4575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Clanes aliados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   4320
      Width           =   4575
      Begin VB.ListBox aliados 
         Height          =   1035
         ItemData        =   "frmGuildNews.frx":0000
         Left            =   120
         List            =   "frmGuildNews.frx":0002
         TabIndex        =   6
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Clanes con los que estamos en guerra"
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   4575
      Begin VB.ListBox guerra 
         Height          =   1035
         ItemData        =   "frmGuildNews.frx":0004
         Left            =   120
         List            =   "frmGuildNews.frx":0006
         TabIndex        =   4
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Noticias"
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4575
      Begin VB.TextBox news 
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   6300
      Width           =   4575
   End
End
Attribute VB_Name = "frmGuildNews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmGuildNews - ImperiumAO - v1.3.0
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
Unload Me
End Sub

Public Sub ParseGuildNews(ByVal s As String)

news = Replace(General_Field_Read(1, s, "¬"), "º", vbCrLf)

Dim h%, j%

h% = Val(General_Field_Read(2, s, "¬"))

For j% = 1 To h%
    
    guerra.AddItem General_Field_Read(j% + 2, s, "¬")
    
Next j%

j% = j% + 2

h% = Val(General_Field_Read(j%, s, "¬"))

For j% = j% + 1 To j% + h%
    
    aliados.AddItem General_Field_Read(j%, s, "¬")
    
Next j%

Me.Show vbModeless, frmMain

End Sub

Private Sub Form_Deactivate()
Me.SetFocus
End Sub

