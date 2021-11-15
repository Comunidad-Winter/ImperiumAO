VERSION 5.00
Begin VB.Form frmPeaceProp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ofertas de paz"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   4980
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
   ScaleHeight     =   2895
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "&Rechazar"
      Height          =   495
      Left            =   3720
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   2520
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Detalles"
      Height          =   495
      Left            =   1320
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   495
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2280
      Width           =   975
   End
   Begin VB.ListBox lista 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      ItemData        =   "frmPeaceProp.frx":0000
      Left            =   120
      List            =   "frmPeaceProp.frx":0002
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmPeaceProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmPeaceProp - ImperiumAO - v1.3.0
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

Public Sub ParsePeaceOffers(ByVal s As String)

Dim t%, r%

t% = Val(General_Field_Read(1, s, ","))

For r% = 1 To t%
    Call lista.AddItem(General_Field_Read(r% + 1, s, ","))
Next r%

Me.Show vbModeless, frmMain

End Sub

Private Sub Command2_Click()
'Me.Visible = False
Call SendData("PEACEDET" & lista.List(lista.ListIndex))
End Sub

Private Sub Command3_Click()
'Me.Visible = False
Call SendData("ACEPPEAT" & lista.List(lista.ListIndex))
Unload Me
End Sub

Private Sub Form_Deactivate()
Me.SetFocus
End Sub

