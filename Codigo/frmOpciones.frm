VERSION 5.00
Begin VB.Form frmOpciones 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6885
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
   ForeColor       =   &H00000000&
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdControles 
      Caption         =   "Con&figuración de controles"
      Height          =   360
      Left            =   150
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   5700
      Width           =   3255
   End
   Begin VB.Frame Frame4 
      Caption         =   "General"
      Height          =   3375
      Left            =   3480
      TabIndex        =   9
      Top             =   3090
      Width           =   3285
      Begin VB.ListBox lstIgnore 
         Height          =   2010
         Left            =   180
         Sorted          =   -1  'True
         TabIndex        =   26
         Top             =   1140
         Width           =   2895
      End
      Begin VB.CheckBox chkop 
         Caption         =   "Habilitar mensajes globales"
         Height          =   285
         Index           =   9
         Left            =   180
         TabIndex        =   17
         Top             =   570
         Width           =   2715
      End
      Begin VB.CheckBox chkop 
         Caption         =   "Uso inteligente de consola"
         Height          =   285
         Index           =   8
         Left            =   180
         TabIndex        =   16
         Top             =   300
         Width           =   2715
      End
      Begin VB.Label Label4 
         Caption         =   "Lista de ignorados (click derecho: menú)"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   27
         Top             =   900
         Width           =   2925
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Apariencia y performance"
      Height          =   2865
      Left            =   3480
      TabIndex        =   7
      Top             =   120
      Width           =   3285
      Begin VB.ListBox lstSkin 
         Height          =   1230
         Left            =   180
         Sorted          =   -1  'True
         TabIndex        =   29
         Top             =   1410
         Width           =   2895
      End
      Begin VB.CheckBox chkop 
         Caption         =   "Ver nombres de los jugadores"
         Height          =   285
         Index           =   3
         Left            =   180
         TabIndex        =   28
         Top             =   840
         Width           =   2715
      End
      Begin VB.CheckBox chkop 
         Caption         =   "Ver diálogos en la consola"
         Height          =   285
         Index           =   7
         Left            =   180
         TabIndex        =   15
         Top             =   570
         Width           =   2715
      End
      Begin VB.CheckBox chkop 
         Caption         =   "Ver nombre del mapa"
         Height          =   285
         Index           =   4
         Left            =   180
         TabIndex        =   14
         Top             =   300
         Width           =   2715
      End
      Begin VB.Label Label4 
         Caption         =   "Skins instalados"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   30
         Top             =   1200
         Width           =   2925
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Información"
      Height          =   1335
      Left            =   135
      TabIndex        =   6
      Top             =   4290
      Width           =   3255
      Begin VB.CommandButton cmdWeb 
         Caption         =   "&www.imperiumao.com.ar"
         Height          =   345
         Left            =   180
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   360
         Width           =   2895
      End
      Begin VB.CommandButton cmdAyuda 
         Caption         =   "¿Necesitás &ayuda?"
         Height          =   345
         Left            =   180
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   840
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Audio"
      Height          =   4065
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3255
      Begin VB.CheckBox chkInvertir 
         Caption         =   "Invertir canales de audio (L / R)"
         Height          =   255
         Left            =   180
         TabIndex        =   31
         Top             =   1530
         Width           =   2775
      End
      Begin VB.HScrollBar scrMidi 
         Height          =   315
         LargeChange     =   15
         Left            =   150
         Max             =   0
         Min             =   -4000
         SmallChange     =   2
         TabIndex        =   23
         Top             =   3570
         Width           =   2895
      End
      Begin VB.HScrollBar scrAmbient 
         Enabled         =   0   'False
         Height          =   315
         LargeChange     =   15
         Left            =   150
         Max             =   0
         Min             =   -4000
         SmallChange     =   2
         TabIndex        =   22
         Top             =   3000
         Width           =   2895
      End
      Begin VB.HScrollBar scrVolume 
         Height          =   315
         LargeChange     =   15
         Left            =   150
         Max             =   0
         Min             =   -4000
         SmallChange     =   2
         TabIndex        =   20
         Top             =   2400
         Width           =   2895
      End
      Begin VB.CheckBox chkop 
         Caption         =   "Sonido habilitado"
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   18
         Top             =   600
         Width           =   2715
      End
      Begin VB.CheckBox chkop 
         Caption         =   "Efectos de navegación"
         Height          =   285
         Index           =   2
         Left            =   180
         TabIndex        =   13
         Top             =   900
         Width           =   2715
      End
      Begin VB.CheckBox chkop 
         Caption         =   "Música habilitada"
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   12
         Top             =   300
         Width           =   2715
      End
      Begin VB.TextBox txtMidi 
         Height          =   285
         Left            =   2385
         TabIndex        =   1
         Top             =   1845
         Width           =   345
      End
      Begin VB.CheckBox chkMidi 
         Caption         =   "Reproducir midi default de la zona"
         Height          =   255
         Left            =   180
         TabIndex        =   0
         Top             =   1230
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Volúmen de música"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   25
         Top             =   3360
         Width           =   2865
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Volúmen de sonidos ambientales"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   24
         Top             =   2790
         Width           =   2865
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Volúmen de audio"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   21
         Top             =   2190
         Width           =   2835
      End
      Begin VB.Label lblNextMidi 
         Caption         =   "»"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2760
         TabIndex        =   11
         Top             =   1875
         Width           =   135
      End
      Begin VB.Label lblBackMidi 
         Caption         =   "«"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2265
         TabIndex        =   10
         Top             =   1875
         Width           =   135
      End
      Begin VB.Label lblMidi 
         BackStyle       =   0  'Transparent
         Caption         =   "Reproduciendo midi número"
         Height          =   255
         Left            =   195
         TabIndex        =   8
         Top             =   1875
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   150
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   6090
      Width           =   3255
   End
   Begin VB.Menu mnuIgnore 
      Caption         =   "Ignorar"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuQuitarIgnorado 
         Caption         =   "Quitar"
      End
      Begin VB.Menu mnuAgregarIgnorado 
         Caption         =   "Agregar"
      End
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmOpciones - ImperiumAO - v1.3.0
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
'Augusto José Rando (barrin@imperiumao.com.ar)
'   - First Relase
'*****************************************************************

Option Explicit

Private Sub chkInvertir_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

InvertirSonido = (chkInvertir.Value = 1)
Sound.InvertirSonido = InvertirSonido

End Sub

Private Sub chkMidi_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

If chkMidi.Value = 1 Then
    txtMidi.Enabled = False
    lblNextMidi.Enabled = False
    lblBackMidi.Enabled = False
    
    If CurrentUser.Logged Then
        Call SendData("ENVIAMID")
    Else
        If frmConnect.Visible Then
            If Musica <> CONST_DESHABILITADA Then
                Sound.NextMusic = MUS_VolverInicio
                Sound.Fading = 200
            End If
        End If
    End If

Else
    txtMidi.Enabled = True
    lblNextMidi.Enabled = True
    lblBackMidi.Enabled = True
End If

DefMidi = chkMidi.Value

End Sub

Private Sub chkOp_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Call Sound.Sound_Play(SND_CLICK)

Select Case Index
    Case 0

        If Musica <> CONST_DESHABILITADA Then
            Sound.Music_Stop
            Musica = CONST_DESHABILITADA
            txtMidi.Enabled = False
            chkMidi.Enabled = False
            lblNextMidi.Enabled = False
            lblBackMidi.Enabled = False
            scrMidi.Enabled = False
        Else
            Musica = CONST_MP3
            chkMidi.Enabled = True
            scrMidi.Enabled = True
            
            If chkMidi.Value = 1 Then
                txtMidi.Enabled = False
                lblNextMidi.Enabled = False
                lblBackMidi.Enabled = False
            Else
                txtMidi.Enabled = True
                lblNextMidi.Enabled = True
                lblBackMidi.Enabled = True
            End If
            
            If Sound.Music_Load(Val(txtMidi.Text), Sound.VolumenActualMusic) Then
                Sound.Music_Stop
                Sound.Music_Play
            End If
        End If

        chkop(Index).Value = IIf((Musica > 0), 1, 0)
    
    Case 1

        If fx = 1 Then
            fx = 0
            chkop(2).Enabled = False
            'scrAmbient.Enabled = False
            scrVolume.Enabled = False
            Call Sound.Sound_Stop_All
        Else
            fx = 1
            chkop(2).Enabled = True
            'scrAmbient.Enabled = True
            scrVolume.Enabled = True
            Sound.Ambient_Play
        End If

        chkop(Index).Value = fx

    Case 2

        If FxNavega = 1 Then
            FxNavega = 0
        Else
            FxNavega = 1
        End If

        chkop(Index).Value = FxNavega

    Case 3
        Engine.Engine_Label_Render_Set
        chkop(Index).Value = IIf(Engine.Engine_Label_Render_Get = True, 1, 0)
        
    Case 4
    
        If VerLugar = 0 Then
            VerLugar = 1
            frmMain.Label2(0).Visible = True
        Else
            VerLugar = 0
            frmMain.Label2(0).Visible = False
        End If

        chkop(Index).Value = VerLugar

    Case 7
    
        If CopiarDialogos = 1 Then
            CopiarDialogos = 0
        Else
            CopiarDialogos = 1
        End If

        chkop(Index).Value = CopiarDialogos

    Case 8
    
        If GuardarEXP = 1 Then
            GuardarEXP = 0
            frmMain.tmrExp.Enabled = False
        Else
            GuardarEXP = 1
            frmMain.tmrExp.Enabled = True
        End If

        chkop(Index).Value = GuardarEXP

    Case 9
    
        If MensajesGlobales = 1 Then
            MensajesGlobales = 0
        Else
            MensajesGlobales = 1
        End If

        If CurrentUser.Logged Then Call SendData("GLO" & MensajesGlobales)
        chkop(Index).Value = MensajesGlobales
        

End Select

End Sub

Private Sub cmdAyuda_Click()
frmHlp.Show vbModeless, frmOpciones
End Sub

Private Sub cmdControles_Click()
frmReBind.Show vbModeless, frmOpciones
End Sub

Private Sub cmdWeb_Click()
ShellExecute Me.hwnd, "open", "http://www.imperiumao.com.ar", "", "", 1
End Sub

Private Sub cmdCerrar_Click()
Me.Visible = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If (Button = vbLeftButton) Then
    Call Auto_Drag(Me.hwnd)
Else
    Unload Me
End If

End Sub

Public Sub AgregarIgnorado(ByVal Nick As String)

Dim i As Integer

Nick = UCase$(Nick)

For i = 0 To lstIgnore.ListCount
    If UCase$(lstIgnore.List(i)) = Nick Then
        If CurrentUser.Logged Then
            Call AddtoRichTextBox(frmMain.RecTxt, "Servidor> " & Nick & " ya había sido ignorado.", 0, 0, 0, 0, 0, 0, 8)
        Else
            Call MensajeAdvertencia(Nick & " ya había sido ignorado.")
        End If
        
        Exit Sub
    End If
Next i

lstIgnore.AddItem Nick
If CurrentUser.Logged Then Call AddtoRichTextBox(frmMain.RecTxt, "Servidor> " & Nick & " ha sido ignorado.", 0, 0, 0, 0, 0, 0, 8)

End Sub

Public Sub Init()

On Error Resume Next

Dim t() As String, i As Integer, file_name As String

chkop(0).Value = IIf((Musica > 0), 1, 0)
chkop(1).Value = fx
chkop(2).Value = FxNavega
chkop(3).Value = IIf(Engine.Engine_Label_Render_Get = True, 1, 0)
chkop(4).Value = VerLugar
chkop(7).Value = CopiarDialogos
chkop(8).Value = GuardarEXP
chkop(9).Value = MensajesGlobales
chkMidi.Value = DefMidi
chkInvertir.Value = IIf(InvertirSonido = True, 1, 0)
txtMidi.Text = Sound.MusicActual
scrVolume.Value = FXVolume
scrMidi.Value = MusicVolume

t = Split(ListaIgnorados, "¬")

lstIgnore.Clear

For i = 0 To UBound(t)
    lstIgnore.AddItem t(i)
Next i

lstSkin.Clear
lstSkin.AddItem "Principal"

file_name = Dir$(App.Path & "\Skins")
Do While Len(file_name) > 0
    If Not _
        (file_name = ".") Or _
        (file_name = "..") Or _
        (LCase(Right(Trim(file_name), 3)) <> "iao") _
    Then
        lstSkin.AddItem General_Field_Read(1, file_name, ".")
    End If
    
    file_name = Dir$()
Loop

If Musica <> CONST_DESHABILITADA Then
    chkMidi.Enabled = True
    
    If chkMidi.Value = 1 Then
        txtMidi.Enabled = False
        lblNextMidi.Enabled = False
        lblBackMidi.Enabled = False
    Else
        txtMidi.Enabled = True
        lblNextMidi.Enabled = True
        lblBackMidi.Enabled = True
    End If

    If Musica = CONST_MP3 Then
        lblMidi.Caption = "Reproduciendo MP3 número"
        chkMidi.Caption = "Reproducir MP3 default de la zona"
    End If

Else
    chkMidi.Enabled = False
    txtMidi.Enabled = False
    lblNextMidi.Enabled = False
    lblBackMidi.Enabled = False
    scrMidi.Enabled = False
End If

If Not CurrentUser.Logged Then
    Me.Show vbModeless, frmConnect
Else
    Me.Show vbModeless, frmMain
End If

End Sub

Private Sub lblBackMidi_Click()

If Val(txtMidi.Text) <= 1 Then
    Beep
ElseIf Sound.Music_Load(Val(txtMidi.Text) - 1, Sound.VolumenActualMusic) Then
    txtMidi.Text = Val(txtMidi.Text) - 1
    Sound.Music_Stop
    Sound.Music_Play
Else
    txtMidi.Text = Val(txtMidi.Text) - 1
End If

End Sub

Private Sub lblNextMidi_Click()

If Musica = CONST_MIDI Then
    If Val(txtMidi.Text) > 70 Then
        Beep
    ElseIf Sound.Music_Load(Val(txtMidi.Text) + 1, Sound.VolumenActualMusic) Then
        txtMidi.Text = Val(txtMidi.Text) + 1
        Sound.Music_Stop
        Sound.Music_Play
    End If
Else
    If Val(txtMidi.Text) > 70 Then
        Beep
    ElseIf Sound.Music_Load(Val(txtMidi.Text) + 1, Sound.VolumenActualMusic) Then
        Sound.Music_Stop
        Sound.Music_Play
        txtMidi.Text = Val(txtMidi.Text) + 1
    End If
End If

End Sub

Private Sub lstIgnore_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = vbRightButton Then
    mnuQuitarIgnorado.Enabled = (lstIgnore.ListIndex <> -1)
    PopupMenu mnuIgnore
End If

End Sub

Private Sub mnuAgregarIgnorado_Click()

Dim Resp As String
Resp = InputBox("Escriba el nombre del usuario que desea ignorar (también puede usar el comando /IGNORAR nick)", "Ignorar usuario")
If Resp <> "" Then Call frmOpciones.AgregarIgnorado(Resp)

End Sub

Private Sub mnuQuitarIgnorado_Click()

If lstIgnore.ListIndex = -1 Then Exit Sub
lstIgnore.RemoveItem lstIgnore.ListIndex

End Sub

Private Sub scrMidi_Change()

Sound.Music_Volume_Set scrMidi.Value
Sound.VolumenActualMusicMax = scrMidi.Value
MusicVolume = Sound.VolumenActualMusicMax

End Sub

Private Sub scrVolume_Change()

Sound.VolumenActual = scrVolume.Value
FXVolume = Sound.VolumenActual

End Sub

Private Sub txtMidi_Change()

If Val(txtMidi.Text) > 0 And (Val(txtMidi.Text) <> Sound.MusicActual) Then
    If Not Sound.Music_Load(Val(txtMidi.Text), Sound.VolumenActualMusic) Then
        txtMidi.Text = Sound.MusicActual
    Else
        Sound.Music_Stop
        Sound.Music_Play
    End If
End If

End Sub
