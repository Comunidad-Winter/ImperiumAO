VERSION 5.00
Begin VB.Form frmHerrero 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "$429"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   5265
   ControlBox      =   0   'False
   Icon            =   "frmHerrero.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstEscudos 
      Height          =   2205
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   5085
   End
   Begin VB.ListBox lstCascos 
      Height          =   2205
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   5085
   End
   Begin VB.CommandButton Command6 
      Caption         =   "$61"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   150
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "$60"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2700
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   150
      Width           =   1215
   End
   Begin VB.TextBox txtCantidad 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Text            =   "1"
      Top             =   3000
      Width           =   5055
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
      Height          =   495
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   3390
      Width           =   1815
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
      Height          =   495
      Left            =   3360
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3420
      Width           =   1815
   End
   Begin VB.ListBox lstArmas 
      Height          =   2205
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   5085
   End
   Begin VB.CommandButton Command2 
      Caption         =   "$59"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   150
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "$58"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   150
      Width           =   1245
   End
   Begin VB.ListBox lstArmaduras 
      Height          =   2205
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   5085
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "$22"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2730
      Width           =   4815
   End
End
Attribute VB_Name = "frmHerrero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmHerrero - ImperiumAO - v1.4.5 R5
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
lstArmaduras.Visible = False
lstArmas.Visible = True
lstCascos.Visible = False
lstEscudos.Visible = False
End Sub

Private Sub Command2_Click()
lstArmaduras.Visible = True
lstArmas.Visible = False
lstCascos.Visible = False
lstEscudos.Visible = False
End Sub

Private Sub Command3_Click()

If lstArmas.Visible Then
    If lstArmas.ListIndex = -1 Then Exit Sub
    Call ClientTCP.Send_Data(Herrero_Build, Integer_To_String(ArmasHerrero(lstArmas.ListIndex + 1)) & Long_To_String(CLng(txtCantidad.Text)))
ElseIf lstArmaduras.Visible Then
    If lstArmaduras.ListIndex = -1 Then Exit Sub
    Call ClientTCP.Send_Data(Herrero_Build, Integer_To_String(ArmadurasHerrero(lstArmaduras.ListIndex + 1)) & Long_To_String(CLng(txtCantidad.Text)))
ElseIf lstCascos.Visible Then
    If lstCascos.ListIndex = -1 Then Exit Sub
    Call ClientTCP.Send_Data(Herrero_Build, Integer_To_String(CascosHerrero(lstCascos.ListIndex + 1)) & Long_To_String(CLng(txtCantidad.Text)))
ElseIf lstEscudos.Visible Then
    If lstEscudos.ListIndex = -1 Then Exit Sub
    Call ClientTCP.Send_Data(Herrero_Build, Integer_To_String(EscudosHerrero(lstEscudos.ListIndex + 1)) & Long_To_String(CLng(txtCantidad.Text)))
End If

Unload Me

End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
lstArmaduras.Visible = False
lstArmas.Visible = False
lstCascos.Visible = True
lstEscudos.Visible = False
End Sub

Private Sub Command6_Click()
lstArmaduras.Visible = False
lstArmas.Visible = False
lstCascos.Visible = False
lstEscudos.Visible = True
End Sub

Private Sub Form_Load()
Call FormParser.Parse_Form(Me)
End Sub

Private Sub txtCantidad_Change()

If Val(txtCantidad.Text) < 0 Then
    txtCantidad.Text = 1
End If

End Sub
