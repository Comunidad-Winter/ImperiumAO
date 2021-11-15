VERSION 5.00
Begin VB.Form frmGuildDetails 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "$435"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   6780
   ClipControls    =   0   'False
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
   Icon            =   "frmGuildDetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame framAlign 
      Caption         =   "$41"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   14
      Top             =   6450
      Width           =   6495
      Begin VB.ComboBox cmbAlineamiento 
         Height          =   315
         ItemData        =   "frmGuildDetails.frx":000C
         Left            =   1800
         List            =   "frmGuildDetails.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "$42"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   6255
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "$1"
      Height          =   375
      Index           =   1
      Left            =   5160
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   8250
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "$2"
      Height          =   375
      Index           =   0
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   8250
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Codex"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      TabIndex        =   2
      Top             =   2130
      Width           =   6495
      Begin VB.TextBox txtCodex1 
         Height          =   285
         Index           =   7
         Left            =   360
         TabIndex        =   11
         Top             =   3720
         Width           =   5655
      End
      Begin VB.TextBox txtCodex1 
         Height          =   285
         Index           =   6
         Left            =   360
         TabIndex        =   10
         Top             =   3360
         Width           =   5655
      End
      Begin VB.TextBox txtCodex1 
         Height          =   285
         Index           =   5
         Left            =   360
         TabIndex        =   9
         Top             =   3000
         Width           =   5655
      End
      Begin VB.TextBox txtCodex1 
         Height          =   285
         Index           =   4
         Left            =   360
         TabIndex        =   8
         Top             =   2640
         Width           =   5655
      End
      Begin VB.TextBox txtCodex1 
         Height          =   285
         Index           =   3
         Left            =   360
         TabIndex        =   7
         Top             =   2280
         Width           =   5655
      End
      Begin VB.TextBox txtCodex1 
         Height          =   285
         Index           =   2
         Left            =   360
         TabIndex        =   6
         Top             =   1920
         Width           =   5655
      End
      Begin VB.TextBox txtCodex1 
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   1560
         Width           =   5655
      End
      Begin VB.TextBox txtCodex1 
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   1200
         Width           =   5655
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "$40"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   6255
      End
   End
   Begin VB.Frame frmDesc 
      Caption         =   "$39"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   6495
      Begin VB.TextBox txtDesc 
         Height          =   1455
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   6135
      End
   End
End
Attribute VB_Name = "frmGuildDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmGuildDetails - ImperiumAO - v1.4.5 R5
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

Option Explicit


Private Sub Command1_Click(Index As Integer)

Dim k As Integer
Dim Cont As Integer
Dim chunk$

Select Case Index

Case 0
    Unload Me
Case 1
    Dim fdesc$
    fdesc$ = Replace(txtDesc, vbCrLf, "º", , , vbBinaryCompare)
    
    For k = 0 To txtCodex1.UBound
        If Len(txtCodex1(k).Text) > 0 Then Cont = Cont + 1
    Next k
    
    If Cont < 4 Then
        MensajeAdvertencia Locale_GUI_Frase(167)
        Exit Sub
    End If
        
    If CurrentUser.CreandoClan Then
            
        If LenB(cmbAlineamiento.Text) = 0 Then
            MensajeAdvertencia Locale_GUI_Frase(166)
            Exit Sub
        End If
        
        chunk$ = fdesc$
        chunk$ = chunk$ & "¬" & CurrentUser.ClanName & "¬" & CurrentUser.Site & "¬" & Cont
        '[Barrin]
        If UCase$(cmbAlineamiento.Text) = "REPUBLICANO" Then
            chunk$ = chunk$ & "¬" & Republicano
        ElseIf UCase$(cmbAlineamiento.Text) = "LEGAL" Then
            chunk$ = chunk$ & "¬" & Legal
        ElseIf UCase$(cmbAlineamiento.Text) = "CAOTICO" Then
            chunk$ = chunk$ & "¬" & Caotico
        End If
        '[/Barrin]
    
        For k = 0 To txtCodex1.UBound
            chunk$ = chunk$ & "¬" & txtCodex1(k)
        Next k
        
        Call ClientTCP.Send_Data(Guild_Create, chunk$)
    Else
        chunk$ = fdesc$ & "¬" & Cont
            
        For k = 0 To txtCodex1.UBound
            chunk$ = chunk$ & "¬" & txtCodex1(k)
        Next k
        
        Call ClientTCP.Send_Data(Guild_Desc_Codec_Update, chunk$)
    End If
            
    CurrentUser.CreandoClan = False
    
    Unload Me
    
End Select

End Sub

Private Sub Form_Load()
Call FormParser.Parse_Form(Me)
End Sub
