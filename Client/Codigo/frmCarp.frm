VERSION 5.00
Begin VB.Form frmCarp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "$444"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   4380
   ControlBox      =   0   'False
   Icon            =   "frmCarp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCantidad 
      Height          =   285
      Left            =   180
      TabIndex        =   3
      Text            =   "1"
      Top             =   2280
      Width           =   4035
   End
   Begin VB.ListBox lstArmas 
      Height          =   1815
      Left            =   150
      TabIndex        =   2
      Top             =   150
      Width           =   4080
   End
   Begin VB.CommandButton Command3 
      Caption         =   "$1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2520
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2670
      Width           =   1710
   End
   Begin VB.CommandButton Command4 
      Caption         =   "$2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   180
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   2670
      Width           =   1710
   End
   Begin VB.Label lblCantidad 
      BackStyle       =   0  'Transparent
      Caption         =   "$22"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   180
      TabIndex        =   4
      Top             =   2040
      Width           =   1395
   End
End
Attribute VB_Name = "frmCarp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmCarp - ImperiumAO - v1.4.5 R5
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

Private Sub Command3_Click()

If lstArmas.ListIndex = -1 Then Exit Sub
Call ClientTCP.Send_Data(Carpenter_Build, Integer_To_String(ObjCarpintero(lstArmas.ListIndex + 1)) & Long_To_String(CLng(txtCantidad.Text)))
Unload Me

End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call FormParser.Parse_Form(Me)
End Sub

Private Sub txtCantidad_Change()

If Val(txtCantidad.Text) < 0 Then
    txtCantidad.Text = 1
End If

End Sub
