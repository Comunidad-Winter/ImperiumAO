VERSION 5.00
Begin VB.Form frmCharList 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "ImperiumAO 1.3"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   FillColor       =   &H00000040&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCharList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   9
      Left            =   8445
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   9
      Top             =   5730
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   8
      Left            =   6945
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   8
      Top             =   5730
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   7
      Left            =   5445
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   7
      Top             =   5730
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   6
      Left            =   3945
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   6
      Top             =   5730
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   5
      Left            =   2445
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   5
      Top             =   5730
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   4
      Left            =   8445
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   4
      Top             =   3930
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   3
      Left            =   6945
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   3
      Top             =   3930
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   2
      Left            =   5445
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   2
      Top             =   3930
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   1
      Left            =   3945
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   1
      Top             =   3930
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   0
      Left            =   2445
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   0
      Top             =   3930
      Width           =   1140
   End
   Begin VB.Label lblCharData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clase"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   6210
      TabIndex        =   23
      Top             =   7920
      Width           =   390
   End
   Begin VB.Label lblCharData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ubicaci�n"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   6210
      TabIndex        =   22
      Top             =   7770
      Width           =   675
   End
   Begin VB.Label lblCharData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel"
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   0
      Left            =   6210
      TabIndex        =   21
      Top             =   7620
      Width           =   1605
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   10
      Left            =   8340
      TabIndex        =   20
      Top             =   5370
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   9
      Left            =   6840
      TabIndex        =   19
      Top             =   5370
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   8
      Left            =   5325
      TabIndex        =   18
      Top             =   5370
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   7
      Left            =   3840
      TabIndex        =   17
      Top             =   5370
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   6
      Left            =   2325
      TabIndex        =   16
      Top             =   5370
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   5
      Left            =   8340
      TabIndex        =   15
      Top             =   3570
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   4
      Left            =   6840
      TabIndex        =   14
      Top             =   3570
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   3
      Left            =   5325
      TabIndex        =   13
      Top             =   3570
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   2
      Left            =   3840
      TabIndex        =   12
      Top             =   3570
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   1
      Left            =   2325
      TabIndex        =   11
      Top             =   3570
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de la cuenta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   10
      Top             =   2370
      Width           =   3705
   End
   Begin VB.Image imgAccion 
      Height          =   300
      Index           =   6
      Left            =   11640
      Tag             =   "0"
      Top             =   60
      Width           =   300
   End
   Begin VB.Image imgAccion 
      Height          =   300
      Index           =   5
      Left            =   11325
      Tag             =   "0"
      Top             =   60
      Width           =   300
   End
   Begin VB.Image WebLink 
      Height          =   345
      Left            =   8490
      MousePointer    =   99  'Custom
      Top             =   8550
      Width           =   3405
   End
   Begin VB.Image imgAccion 
      Height          =   615
      Index           =   4
      Left            =   8025
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   7575
      Width           =   1755
   End
   Begin VB.Image imgAccion 
      Height          =   615
      Index           =   3
      Left            =   4155
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   7575
      Width           =   1755
   End
   Begin VB.Image imgAccion 
      Height          =   615
      Index           =   2
      Left            =   8025
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   2085
      Width           =   1755
   End
   Begin VB.Image imgAccion 
      Height          =   615
      Index           =   0
      Left            =   2235
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   7575
      Width           =   1755
   End
   Begin VB.Image imgAccion 
      Height          =   615
      Index           =   1
      Left            =   6180
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   2085
      Width           =   1755
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   9
      Left            =   8280
      Top             =   5310
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   8
      Left            =   6780
      Top             =   5310
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   7
      Left            =   5280
      Top             =   5310
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   6
      Left            =   3780
      Top             =   5310
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   5
      Left            =   2280
      Top             =   5310
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   4
      Left            =   8280
      Top             =   3510
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   3
      Left            =   6780
      Top             =   3510
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   2
      Left            =   5280
      Top             =   3510
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   1
      Left            =   3780
      Top             =   3510
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   0
      Left            =   2280
      Top             =   3510
      Width           =   1455
   End
End
Attribute VB_Name = "frmCharList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmConnect - ImperiumAO - v1.4.5 R5
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
'Augusto Jos� Rando (barrin@imperiumao.com.ar)
'   - First Relase
'*****************************************************************

Option Explicit

Private intSelChar As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then
    frmMain.MainWinsock.Close
    Call FormParser.Parse_Form(frmConnect)
    frmConnect.Visible = True
    Me.Visible = False
End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If (Button = vbLeftButton) And (RunWindowed = 1) Then Call Auto_Drag(Me.hwnd)
End Sub

Private Sub Form_Load()

Me.Caption = Form_Caption
Me.Picture = General_Load_Picture_From_Resource_Ex("cuenta")

Call FormParser.Parse_Form(Me)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim i As Integer

For i = 0 To imgAccion.UBound
    If imgAccion(i).Tag = "0" Then
        If (i <> 3 And i <> 4) Or (intSelChar > 0 And LenB(lblAccData(intSelChar)) > 0) Then
            imgAccion(i).Picture = Nothing
            imgAccion(i).Tag = "1"
        End If
    End If
Next i

End Sub

Private Sub imgAcc_Click(Index As Integer)

On Error Resume Next

If Index + 1 <> intSelChar Then
    If intSelChar > 0 Then imgAcc(intSelChar - 1) = Nothing
    intSelChar = Index + 1
    imgAcc(Index) = General_Load_Picture_From_Resource_Ex("slot" & intSelChar)
    
    If LenB(lblAccData(intSelChar)) > 0 Then
        imgAccion(3).Picture = Nothing
        imgAccion(4).Picture = Nothing
        lblCharData(0) = "Nivel " & ListaPersonajes(intSelChar).char_level 'Nivel
        lblCharData(1) = Map_NameLoad(ListaPersonajes(intSelChar).char_map) 'Ubicacion

        If ListaPersonajes(intSelChar).char_clase > NUMCLASES Then
            lblCharData(2) = "GM"
        Else
            lblCharData(2) = ListaClases(ListaPersonajes(intSelChar).char_clase) 'Clase raza
        End If

    Else
        imgAccion(3).Picture = General_Load_Picture_From_Resource_Ex("accborrardes")
        imgAccion(4).Picture = General_Load_Picture_From_Resource_Ex("acccondes")
        lblCharData(0) = vbNullString
        lblCharData(1) = vbNullString
        lblCharData(2) = vbNullString
    End If
    
End If

End Sub

Private Sub lblAccData_Click(Index As Integer)

On Error Resume Next

If Index <> intSelChar Then
    If intSelChar > 0 Then imgAcc(intSelChar - 1) = Nothing
    intSelChar = Index
    imgAcc(Index - 1) = General_Load_Picture_From_Resource_Ex("slot" & intSelChar)
    
    If LenB(lblAccData(intSelChar)) > 0 Then
        imgAccion(3).Picture = Nothing
        imgAccion(4).Picture = Nothing
        lblCharData(0) = "Nivel " & ListaPersonajes(intSelChar).char_level 'Nivel
        lblCharData(1) = Map_NameLoad(ListaPersonajes(intSelChar).char_map) 'Ubicacion
        
        If ListaPersonajes(intSelChar).char_clase > NUMCLASES Then
            lblCharData(2) = "GM"
        Else
            lblCharData(2) = ListaClases(ListaPersonajes(intSelChar).char_clase) 'Clase raza
        End If
        
    Else
        imgAccion(3).Picture = General_Load_Picture_From_Resource_Ex("accborrardes")
        imgAccion(4).Picture = General_Load_Picture_From_Resource_Ex("acccondes")
        lblCharData(0) = vbNullString
        lblCharData(1) = vbNullString
        lblCharData(2) = vbNullString
    End If
    
End If

End Sub

Private Sub picChar_Click(Index As Integer)

On Error Resume Next

If Index + 1 <> intSelChar Then
    If intSelChar > 0 Then imgAcc(intSelChar - 1) = Nothing
    intSelChar = Index + 1
    imgAcc(Index) = General_Load_Picture_From_Resource_Ex("slot" & intSelChar)

    If LenB(lblAccData(intSelChar)) > 0 Then
        imgAccion(3).Picture = Nothing
        imgAccion(4).Picture = Nothing
        lblCharData(0) = "Nivel " & ListaPersonajes(intSelChar).char_level 'Nivel
        lblCharData(1) = Map_NameLoad(ListaPersonajes(intSelChar).char_map) 'Ubicacion
        
        If ListaPersonajes(intSelChar).char_clase > NUMCLASES Then
            lblCharData(2) = "GM"
        Else
            lblCharData(2) = ListaClases(ListaPersonajes(intSelChar).char_clase) 'Clase raza
        End If
        
    Else
        imgAccion(3).Picture = General_Load_Picture_From_Resource_Ex("accborrardes")
        imgAccion(4).Picture = General_Load_Picture_From_Resource_Ex("acccondes")
        lblCharData(0) = vbNullString
        lblCharData(1) = vbNullString
        lblCharData(2) = vbNullString
    End If

End If

End Sub

Private Sub imgAccion_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If (Index = 3 Or Index = 4) And (intSelChar <= 0 Or LenB(ListaPersonajes(intSelChar).char_name) <= 0) Then Exit Sub

Call imgAccionRestaurar
Call Sound.Sound_Play(SND_CLICK)

Select Case Index
    
    Case 0 'Crear personaje
        If sMusica <> CONST_DESHABILITADA Then
            If sMusica <> CONST_DESHABILITADA Then
                Sound.NextMusic = MUS_CrearPersonaje
                Sound.Fading = 200
            End If
        End If
        
        CurrentUser.CurrentCharIndex = 0
        
        Me.Visible = False
        frmCrearPersonaje.Show

    Case 1 'Cambiar pass
        ShellExecute Me.hwnd, "open", Chr$(34) & "http://www.imperiumao.com.ar/" & GameLocale & "/cambiopass.php" & Chr$(34), vbNullString, vbNullString, 1
        
    Case 2 'Salir
        frmMain.MainWinsock.Close
        Call FormParser.Parse_Form(frmConnect)
        frmConnect.Visible = True
        Me.Visible = False

    Case 3 'Borrar
        
        If ListaPersonajes(intSelChar).char_level >= 13 Then
            MensajeAdvertencia Locale_GUI_Frase(231)
            Exit Sub
        End If
        
        If InputBox(Locale_GUI_Frase(232) & " " & Chr$(34) & ListaPersonajes(intSelChar).char_name & Chr$(34) & Locale_GUI_Frase(233), Locale_GUI_Frase(234)) = Locale_GUI_Frase(235) Then
            Call ClientTCP.Send_Data(Char_Erase, ListaPersonajes(intSelChar).char_name)
        End If

    Case 4 'Conectar
        
        If FormParser.GetDefaultCursor(Me) = E_WAIT Then Exit Sub
        Call FormParser.Parse_Form(Me, E_WAIT)
        CurrentUser.UserName = ListaPersonajes(intSelChar).char_name
        CurrentUser.CurrentCharIndex = intSelChar
        Call ClientTCP.Send_Data(Char_Login, CurrentUser.UserName)
        
    Case 5
        Me.WindowState = vbMinimized
    Case 6
        Call EndGame(True)
    
End Select

End Sub

Private Sub imgAccion_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Select Case Index
    Case 0 'Crear personaje
        imgAccion(0).Picture = General_Load_Picture_From_Resource_Ex("acccredown")
        imgAccion(0).Tag = "0"
    Case 1 'Cambiar pass
        imgAccion(1).Picture = General_Load_Picture_From_Resource_Ex("acccambiardown")
        imgAccion(1).Tag = "0"
    Case 2 'Salir
        imgAccion(2).Picture = General_Load_Picture_From_Resource_Ex("accsaldown")
        imgAccion(2).Tag = "0"
    Case 3 'Borrar
        If intSelChar <= 0 Or LenB(lblAccData(intSelChar)) <= 0 Then Exit Sub
        imgAccion(3).Picture = General_Load_Picture_From_Resource_Ex("accborrardown")
        imgAccion(3).Tag = "0"
    Case 4 'Conectar
        If intSelChar <= 0 Or LenB(lblAccData(intSelChar)) <= 0 Then Exit Sub
        imgAccion(4).Picture = General_Load_Picture_From_Resource_Ex("acccondown")
        imgAccion(4).Tag = "0"
    Case 5 'Minimizar
        imgAccion(5).Picture = General_Load_Picture_From_Resource_Ex("conmindown")
        imgAccion(5).Tag = "0"
    Case 6 'Cerrar X
        imgAccion(6).Picture = General_Load_Picture_From_Resource_Ex("concedown")
        imgAccion(6).Tag = "0"
End Select

'Call imgAccionRestaurar(Index)

End Sub

Private Sub imgAccion_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Select Case Index
    Case 0 'Crear personaje
        If imgAccion(0).Tag = "1" Then
            imgAccion(0).Picture = General_Load_Picture_From_Resource_Ex("acccreover")
            imgAccion(0).Tag = "0"
            Call Sound.Sound_Play(SND_OVER)
        End If
        
    Case 1 'Cambiar pass
        If imgAccion(1).Tag = "1" Then
            imgAccion(1).Picture = General_Load_Picture_From_Resource_Ex("acccambiarover")
            imgAccion(1).Tag = "0"
            Call Sound.Sound_Play(SND_OVER)
        End If
        
    Case 2 'Salir
        If imgAccion(2).Tag = "1" Then
            imgAccion(2).Picture = General_Load_Picture_From_Resource_Ex("accsaover")
            imgAccion(2).Tag = "0"
            Call Sound.Sound_Play(SND_OVER)
        End If
        
    Case 3 'Borrar
        If intSelChar <= 0 Or LenB(lblAccData(intSelChar)) <= 0 Then Exit Sub
        
        If imgAccion(3).Tag = "1" Then
            imgAccion(3).Picture = General_Load_Picture_From_Resource_Ex("accborrarover")
            imgAccion(3).Tag = "0"
            Call Sound.Sound_Play(SND_OVER)
        End If
        
    Case 4 'Conectar
        If intSelChar <= 0 Or LenB(lblAccData(intSelChar)) <= 0 Then Exit Sub
        
        If imgAccion(4).Tag = "1" Then
            imgAccion(4).Picture = General_Load_Picture_From_Resource_Ex("accconover")
            imgAccion(4).Tag = "0"
            Call Sound.Sound_Play(SND_OVER)
        End If
        
    Case 5 'Minimizar
        If imgAccion(5).Tag = "1" Then
            imgAccion(5).Picture = General_Load_Picture_From_Resource_Ex("conminover")
            imgAccion(5).Tag = "0"
            Call Sound.Sound_Play(SND_OVER)
        End If
        
    Case 6 'Cerrar X
        If imgAccion(6).Tag = "1" Then
            imgAccion(6).Picture = General_Load_Picture_From_Resource_Ex("conceover")
            imgAccion(6).Tag = "0"
            Call Sound.Sound_Play(SND_OVER)
        End If
End Select

Call imgAccionRestaurar(Index)

End Sub

Private Sub imgAccionRestaurar(Optional ByVal NoIndex As Integer = 1000)

Dim i As Integer

For i = 0 To imgAccion.UBound
    If i <> NoIndex Then
        If (i <> 3 And i <> 4) Or (intSelChar > 0 And LenB(lblAccData(intSelChar)) > 0) Then
            imgAccion(i).Picture = Nothing
            imgAccion(i).Tag = "1"
        End If
    End If
Next i

End Sub

Private Sub picChar_DblClick(Index As Integer)

intSelChar = Index + 1

If LenB(lblAccData(intSelChar)) > 0 Then
    If FormParser.GetDefaultCursor(Me) = E_WAIT Then Exit Sub
    Call FormParser.Parse_Form(Me, E_WAIT)
    CurrentUser.UserName = ListaPersonajes(intSelChar).char_name
    CurrentUser.CurrentCharIndex = intSelChar
    Call ClientTCP.Send_Data(Char_Login, CurrentUser.UserName)
End If

End Sub

Private Sub lblAccData_DblClick(Index As Integer)

intSelChar = Index

If LenB(lblAccData(intSelChar)) > 0 Then
    If FormParser.GetDefaultCursor(Me) = E_WAIT Then Exit Sub
    Call FormParser.Parse_Form(Me, E_WAIT)
    CurrentUser.UserName = ListaPersonajes(intSelChar).char_name
    CurrentUser.CurrentCharIndex = intSelChar
    Call ClientTCP.Send_Data(Char_Login, CurrentUser.UserName)
End If

End Sub

Private Sub imgAcc_DblClick(Index As Integer)

intSelChar = Index + 1

If LenB(lblAccData(intSelChar)) > 0 Then
    If FormParser.GetDefaultCursor(Me) = E_WAIT Then Exit Sub
    Call FormParser.Parse_Form(Me, E_WAIT)
    CurrentUser.UserName = ListaPersonajes(intSelChar).char_name
    CurrentUser.CurrentCharIndex = intSelChar
    Call ClientTCP.Send_Data(Char_Login, CurrentUser.UserName)
End If

End Sub

Private Sub picChar_Paint(Index As Integer)

Dim i As Long

'For i = 1 To UBound(ListaPersonajes)
    If LenB(ListaPersonajes(Index + 1).char_name) > 0 Then
        frmMain.Engine.Char_Render_Start
        
        Call frmMain.Engine.Char_Render_To_HWnd(ListaPersonajes(Index + 1).char_head, ListaPersonajes(Index + 1).char_body, ListaPersonajes(Index + 1).char_weapon, ListaPersonajes(Index + 1).char_shield, ListaPersonajes(Index + 1).char_helmet, ListaPersonajes(Index + 1).char_familiar)
        
        frmMain.Engine.Char_Render_End frmCharList.picChar(Index).hwnd
    End If
'Next

End Sub

Private Sub WebLink_Click()
ShellExecute Me.hwnd, "open", Chr$(34) & "http://www.imperiumao.com.ar/" & GameLocale & "/" & Chr$(34), vbNullString, vbNullString, 1
End Sub

Public Sub ClearChars()

Dim i As Long

intSelChar = 0

For i = 0 To picChar.UBound
    picChar(i).Picture = Nothing
    imgAcc(i).Picture = Nothing
    lblAccData(i + 1).Caption = vbNullString
Next i

imgAccion(3).Picture = General_Load_Picture_From_Resource_Ex("accborrardes")
imgAccion(4).Picture = General_Load_Picture_From_Resource_Ex("acccondes")
lblCharData(0) = vbNullString
lblCharData(1) = vbNullString
lblCharData(2) = vbNullString

End Sub

Public Sub PaintChars()

Dim i As Integer

For i = 0 To picChar.UBound
    Call picChar_Paint(i)
Next i

End Sub
