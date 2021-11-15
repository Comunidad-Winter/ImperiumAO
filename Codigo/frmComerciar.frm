VERSION 5.00
Begin VB.Form frmComerciar 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6975
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
   ScaleHeight     =   509
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrNumber 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   30
      Top             =   30
   End
   Begin VB.TextBox Cantidad 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Text            =   "1"
      Top             =   6960
      Width           =   510
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   900
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   1620
      Width           =   480
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   3960
      Index           =   1
      Left            =   3720
      TabIndex        =   1
      Top             =   2580
      Width           =   2490
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   3960
      Index           =   0
      Left            =   780
      TabIndex        =   0
      Top             =   2580
      Width           =   2460
   End
   Begin VB.Image cmdMasMenos 
      Height          =   420
      Index           =   0
      Left            =   2940
      Tag             =   "1"
      Top             =   6870
      Width           =   195
   End
   Begin VB.Image cmdMasMenos 
      Height          =   420
      Index           =   1
      Left            =   3840
      Tag             =   "1"
      Top             =   6870
      Width           =   195
   End
   Begin VB.Image Command2 
      Height          =   345
      Left            =   6480
      Tag             =   "1"
      Top             =   180
      Width           =   345
   End
   Begin VB.Image Image1 
      Height          =   450
      Index           =   1
      Left            =   4230
      Tag             =   "1"
      Top             =   6855
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   450
      Index           =   0
      Left            =   585
      Tag             =   "1"
      Top             =   6855
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   3
      Left            =   1560
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   5520
      TabIndex        =   5
      Top             =   1530
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   5100
      TabIndex        =   4
      Top             =   1890
      Width           =   1035
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   1560
      TabIndex        =   3
      Top             =   1530
      Width           =   2985
   End
End
Attribute VB_Name = "frmComerciar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmComerciar - ImperiumAO - v1.3.0
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

Public LastIndex1 As Integer
Public LastIndex2 As Integer

Private m_Number As Integer
Private m_Increment As Integer
Private m_Interval As Integer

Private Sub cantidad_Change()

If Val(Cantidad.Text) < 0 Then
    Cantidad.Text = 1
    m_Number = 1
ElseIf Val(Cantidad.Text) > MAX_INVENTORY_OBJS Then
    Cantidad.Text = 1
    m_Number = 1
Else
    m_Number = Val(Cantidad.Text)
End If

End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
    If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub cantidad_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call Form_MouseMove(Button, Shift, x, y)
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call Sound.Sound_Play(SND_CLICK)
Command2.Picture = General_Load_Picture_From_Resource("salir-down.bmp")
Command2.Tag = "1"
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If Command2.Tag = "0" Then
    Command2.Picture = General_Load_Picture_From_Resource("salir-over.bmp")
    Command2.Tag = "1"
End If

End Sub

Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Call Form_MouseMove(Button, Shift, x, y)
SendData "FINCOM"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If (Button = vbLeftButton) Then
    Call Auto_Drag(Me.hwnd)
Else
    SendData "FINCOM"
End If

End Sub

Private Sub Form_Load()
Me.Picture = General_Load_Picture_From_Resource("comercio.bmp")
m_Number = 1
m_Interval = 30
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If Image1(0).Tag = "1" Then
    Image1(0).Picture = Nothing
    Image1(0).Tag = "0"
End If

If Image1(1).Tag = "1" Then
    Image1(1).Picture = Nothing
    Image1(1).Tag = "0"
End If

If cmdMasMenos(0).Tag = "1" Then
    cmdMasMenos(0).Picture = Nothing
    cmdMasMenos(0).Tag = "0"
End If

If cmdMasMenos(1).Tag = "1" Then
    cmdMasMenos(1).Picture = Nothing
    cmdMasMenos(1).Tag = "0"
End If

If Command2.Tag = "1" Then
    Command2.Picture = Nothing
    Command2.Tag = "0"
End If

End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If Index = 0 Then
    Image1(Index).Picture = General_Load_Picture_From_Resource("comprar-down.bmp")
    Image1(Index).Tag = "1"
ElseIf Index = 1 Then
    Image1(Index).Picture = General_Load_Picture_From_Resource("vender-down.bmp")
    Image1(Index).Tag = "1"
End If

End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If Index = 0 Then
    If Image1(Index).Tag = "0" Then
        Image1(Index).Picture = General_Load_Picture_From_Resource("comprar-over.bmp")
        Image1(Index).Tag = "1"
    End If
ElseIf Index = 1 Then
    If Image1(Index).Tag = "0" Then
        Image1(Index).Picture = General_Load_Picture_From_Resource("vender-over.bmp")
        Image1(Index).Tag = "1"
    End If
End If

End Sub

Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Call Sound.Sound_Play(SND_CLICK)
Call Form_MouseMove(Button, Shift, x, y)

If List1(Index).List(List1(Index).ListIndex) = "Nada" Or _
   List1(Index).ListIndex < 0 Then Exit Sub

Select Case Index
    Case 0
        frmComerciar.List1(0).SetFocus
        LastIndex1 = List1(0).ListIndex
        If CurrentUser.UserGLD >= NPCInventory(List1(0).ListIndex + 1).Valor * Val(Cantidad) Then
            SendData ("COMP" & "," & List1(0).ListIndex + 1 & "," & Cantidad.Text)
        Else
            AddtoRichTextBox frmMain.RecTxt, "No tenés suficiente oro.", 2, 51, 223, 1, 1
            Exit Sub
        End If
    Case 1
        LastIndex2 = List1(1).ListIndex
        SendData ("VEND" & "," & List1(1).ListIndex + 1 & "," & Cantidad.Text)
                
End Select

List1(0).Clear
List1(1).Clear

NPCInvDim = 0

End Sub

Private Sub list1_Click(Index As Integer)

Select Case Index
    Case 0
        Label1(0).Caption = NPCInventory(List1(0).ListIndex + 1).Name
        Label1(1).Caption = NPCInventory(List1(0).ListIndex + 1).Valor
        Label1(2).Caption = NPCInventory(List1(0).ListIndex + 1).Amount
        
        Select Case NPCInventory(List1(0).ListIndex + 1).ObjType
            Case 2
                Label1(3).Caption = "Golpe: " & NPCInventory(List1(0).ListIndex + 1).MinHIT & "/" & NPCInventory(List1(0).ListIndex + 1).MaxHIT
                Label1(3).Visible = True
            Case 3
                Label1(3).Caption = "Defensa: " & NPCInventory(List1(0).ListIndex + 1).Def
                Label1(3).Visible = True
            Case Else
                Label1(3).Visible = False
        End Select
        
        If NPCInventory(List1(0).ListIndex + 1).GrhIndex <> 0 Then
            Call Engine.Grh_Render_To_Hdc(NPCInventory(List1(0).ListIndex + 1).GrhIndex, Picture1.hDC, 0, 0)
        Else
            Picture1.Picture = Nothing
        End If
        
    Case 1
        Label1(0).Caption = UserInventory(List1(1).ListIndex + 1).Name
        Label1(1).Caption = UserInventory(List1(1).ListIndex + 1).Valor
        Label1(2).Caption = UserInventory(List1(1).ListIndex + 1).Amount
        Select Case UserInventory(List1(1).ListIndex + 1).ObjType
            Case 2
                Label1(3).Caption = "Golpe: " & UserInventory(List1(1).ListIndex + 1).MinHIT & "/" & UserInventory(List1(1).ListIndex + 1).MaxHIT
                Label1(3).Visible = True
            Case 3
                Label1(3).Caption = "Defensa: " & UserInventory(List1(1).ListIndex + 1).Def
                Label1(3).Visible = True
            Case Else
                Label1(3).Visible = False
        End Select
        
        If UserInventory(List1(1).ListIndex + 1).GrhIndex <> 0 Then
            Call Engine.Grh_Render_To_Hdc(UserInventory(List1(1).ListIndex + 1).GrhIndex, Picture1.hDC, 0, 0)
        Else
            Picture1.Picture = Nothing
        End If
End Select

Picture1.Refresh

End Sub

Private Sub List1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Call Form_MouseMove(Button, Shift, x, y)
End Sub

Private Sub cmdMasMenos_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Call Sound.Sound_Play(SND_CLICK)

Select Case Index
    Case 0
        cmdMasMenos(Index).Picture = General_Load_Picture_From_Resource("menos-down.bmp")
        cmdMasMenos(Index).Tag = "1"
        Cantidad.Text = Str((Val(Cantidad.Text) - 1))
        m_Increment = -1
    Case 1
        cmdMasMenos(Index).Picture = General_Load_Picture_From_Resource("mas-down.bmp")
        cmdMasMenos(Index).Tag = "1"
        m_Increment = 1
End Select

tmrNumber.Interval = 30
tmrNumber.Enabled = True

End Sub

Private Sub cmdMasMenos_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Select Case Index
    Case 0
        If cmdMasMenos(Index).Tag = "0" Then
            cmdMasMenos(Index).Picture = General_Load_Picture_From_Resource("menos-over.bmp")
            cmdMasMenos(Index).Tag = "1"
        End If
    Case 1
        If cmdMasMenos(Index).Tag = "0" Then
            cmdMasMenos(Index).Picture = General_Load_Picture_From_Resource("mas-over.bmp")
            cmdMasMenos(Index).Tag = "1"
        End If
End Select

End Sub

Private Sub cmdMasMenos_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Call Form_MouseMove(Button, Shift, x, y)
tmrNumber.Enabled = False
End Sub

Private Sub tmrNumber_Timer()

Const MIN_NUMBER = 1
Const MAX_NUMBER = 10000

    m_Number = m_Number + m_Increment
    If m_Number < MIN_NUMBER Then
        m_Number = MIN_NUMBER
    ElseIf m_Number > MAX_NUMBER Then
        m_Number = MAX_NUMBER
    End If

    Cantidad.Text = format$(m_Number)
    
    If m_Interval > 1 Then
        m_Interval = m_Interval - 1
        tmrNumber.Interval = m_Interval
    End If

End Sub
