VERSION 5.00
Begin VB.Form frmOpciones 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6870
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
   ScaleHeight     =   7725
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmIdioma 
      Caption         =   "$465"
      Height          =   705
      Left            =   120
      TabIndex        =   37
      Top             =   150
      Width           =   3255
      Begin VB.ComboBox cmbLanguage 
         Height          =   315
         ItemData        =   "frmOpciones.frx":0152
         Left            =   180
         List            =   "frmOpciones.frx":015C
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdControles 
      Caption         =   "$69"
      Height          =   360
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   6870
      Width           =   3255
   End
   Begin VB.Frame Frame4 
      Caption         =   "$68"
      Height          =   4065
      Left            =   3510
      TabIndex        =   9
      Top             =   3570
      Width           =   3285
      Begin VB.CheckBox chkop 
         Caption         =   "$84"
         Height          =   285
         Index           =   11
         Left            =   180
         TabIndex        =   36
         Top             =   1350
         Width           =   2715
      End
      Begin VB.CheckBox chkop 
         Caption         =   "$85"
         Height          =   285
         Index           =   10
         Left            =   180
         TabIndex        =   35
         Top             =   1080
         Width           =   2715
      End
      Begin VB.CheckBox chkop 
         Caption         =   "$86"
         Height          =   285
         Index           =   6
         Left            =   180
         TabIndex        =   34
         Top             =   810
         Width           =   2715
      End
      Begin VB.CheckBox chkop 
         Caption         =   "$87"
         Height          =   285
         Index           =   5
         Left            =   180
         TabIndex        =   31
         Top             =   540
         Width           =   2715
      End
      Begin VB.ListBox lstIgnore 
         Height          =   2010
         Left            =   180
         Sorted          =   -1  'True
         TabIndex        =   25
         Top             =   1890
         Width           =   2895
      End
      Begin VB.CheckBox chkop 
         Caption         =   "$88"
         Height          =   285
         Index           =   9
         Left            =   180
         TabIndex        =   16
         Top             =   270
         Width           =   2715
      End
      Begin VB.Label Label4 
         Caption         =   "$83"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   26
         Top             =   1650
         Width           =   2925
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "$66"
      Height          =   3345
      Left            =   3480
      TabIndex        =   7
      Top             =   120
      Width           =   3285
      Begin VB.ListBox lstSkin 
         Height          =   1635
         Left            =   180
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   28
         Top             =   1350
         Width           =   2895
      End
      Begin VB.CheckBox chkop 
         Caption         =   "$82"
         Height          =   285
         Index           =   3
         Left            =   180
         TabIndex        =   27
         Top             =   840
         Width           =   2715
      End
      Begin VB.CheckBox chkop 
         Caption         =   "$81"
         Height          =   285
         Index           =   7
         Left            =   180
         TabIndex        =   15
         Top             =   570
         Width           =   2715
      End
      Begin VB.CheckBox chkop 
         Caption         =   "$80"
         Height          =   285
         Index           =   4
         Left            =   180
         TabIndex        =   14
         Top             =   300
         Width           =   2715
      End
      Begin VB.Label lblSkinData 
         BackStyle       =   0  'Transparent
         Caption         =   "$90: Desconocido"
         Height          =   195
         Left            =   180
         TabIndex        =   33
         Top             =   3030
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "$75"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   29
         Top             =   1140
         Width           =   2925
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "$67"
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   5190
      Width           =   3255
      Begin VB.CommandButton cmdWeb 
         Caption         =   "$70"
         Height          =   345
         Index           =   1
         Left            =   180
         MousePointer    =   99  'Custom
         TabIndex        =   32
         Top             =   690
         Width           =   2895
      End
      Begin VB.CommandButton cmdWeb 
         Caption         =   "&www.imperiumao.com.ar"
         Height          =   345
         Index           =   0
         Left            =   180
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   300
         Width           =   2895
      End
      Begin VB.CommandButton cmdAyuda 
         Caption         =   "$71"
         Height          =   345
         Left            =   180
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   1080
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "$65"
      Height          =   4095
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   990
      Width           =   3255
      Begin VB.CheckBox chkInvertir 
         Caption         =   "$76"
         Height          =   255
         Left            =   180
         TabIndex        =   30
         Top             =   1530
         Width           =   2985
      End
      Begin VB.HScrollBar scrMidi 
         Height          =   315
         LargeChange     =   15
         Left            =   150
         Max             =   0
         Min             =   -4000
         SmallChange     =   2
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   19
         Top             =   2400
         Width           =   2895
      End
      Begin VB.CheckBox chkop 
         Caption         =   "$78"
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   17
         Top             =   600
         Width           =   2985
      End
      Begin VB.CheckBox chkop 
         Caption         =   "$77"
         Height          =   285
         Index           =   2
         Left            =   180
         TabIndex        =   13
         Top             =   900
         Width           =   2955
      End
      Begin VB.CheckBox chkop 
         Caption         =   "$79"
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   12
         Top             =   300
         Width           =   2985
      End
      Begin VB.TextBox txtMidi 
         Height          =   285
         Left            =   2385
         TabIndex        =   1
         Top             =   1845
         Width           =   345
      End
      Begin VB.CheckBox chkMidi 
         Caption         =   "$91"
         Height          =   255
         Left            =   180
         TabIndex        =   0
         Top             =   1230
         Width           =   2985
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "$72"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   24
         Top             =   3360
         Width           =   2865
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "$73"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   23
         Top             =   2790
         Width           =   2865
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "$74"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   20
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
         Caption         =   "$92"
         Height          =   255
         Left            =   195
         TabIndex        =   8
         Top             =   1875
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "$25"
      Height          =   360
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   7260
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
'frmOpciones - ImperiumAO - v1.4.5 R5
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

Private bLoading As Boolean

Private Sub chkInvertir_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

InvertirSonido = (chkInvertir.value = 1)
Sound.InvertirSonido = InvertirSonido

End Sub

Private Sub chkMidi_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

If chkMidi.value = 1 Then
    txtMidi.Enabled = False
    lblNextMidi.Enabled = False
    lblBackMidi.Enabled = False
    
    If CurrentUser.Logged Then
        If sMusica <> CONST_DESHABILITADA Then
            Sound.NextMusic = Sound.LastMapMusic
            Sound.Fading = 200
        End If
    Else
        If frmConnect.Visible Then
            If sMusica <> CONST_DESHABILITADA Then
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

DefMidi = chkMidi.value

End Sub

Private Sub chkOp_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Dim map_x As Integer, map_y As Integer

Call Sound.Sound_Play(SND_CLICK)

Select Case Index
    Case 0
                
        If sMusica <> CONST_DESHABILITADA Then
            Sound.Music_Stop
            sMusica = CONST_DESHABILITADA
            txtMidi.Enabled = False
            chkMidi.Enabled = False
            lblNextMidi.Enabled = False
            lblBackMidi.Enabled = False
            scrMidi.Enabled = False
        Else
            
            If Sound.Engine_Running = False Then
                
                sMusica = CONST_MP3
                frmPregunta.msg.Caption = Locale_GUI_Frase(335)
                frmPregunta.Unload_And_Update = True
                frmPregunta.Show vbModal, Me
            
            Else
                sMusica = CONST_MP3
                chkMidi.Enabled = True
                scrMidi.Enabled = True
                
                If chkMidi.value = 1 Then
                    txtMidi.Enabled = False
                    lblNextMidi.Enabled = False
                    lblBackMidi.Enabled = False
                Else
                    txtMidi.Enabled = True
                    lblNextMidi.Enabled = True
                    lblBackMidi.Enabled = True
                End If
                
                If Sound.Music_Load(Val(txtMidi.Text), Sound.VolumenActualMusic) Then
                    'Sound.Music_Stop
                    Sound.Music_Play
                End If
            End If
            
        End If

        chkop(Index).value = IIf((sMusica > 0), 1, 0)
    
    Case 1

        If Audio = 1 Then
            chkop(2).Enabled = False
            'scrAmbient.Enabled = False
            scrVolume.Enabled = False
            Call Sound.Sound_Stop_All
            Audio = 0
        Else
        
            If Sound.Engine_Running = False Then
                Audio = 1
                frmPregunta.msg.Caption = Locale_GUI_Frase(336)
                frmPregunta.Unload_And_Update = True
                frmPregunta.Show vbModal, Me
            Else
                Audio = 1
                chkop(2).Enabled = True
                'scrAmbient.Enabled = True
                scrVolume.Enabled = True
                Call Sound.Ambient_Load(Sound.AmbienteActual, Sound.VolumenActualAmbient)
                Call Sound.Ambient_Play
            End If
        
        End If

        chkop(Index).value = Audio

    Case 2

        If FxNavega = 1 Then
            FxNavega = 0
        Else
            FxNavega = 1
        End If

        chkop(Index).value = FxNavega

    Case 3
        frmMain.Engine.Engine_Label_Render_Set
        chkop(Index).value = IIf(frmMain.Engine.Engine_Label_Render_Get = True, 1, 0)
        
    Case 4
    
        If VerLugar = 0 Then
            VerLugar = 1
            frmMain.Label2(0).Caption = frmMain.Engine.Map_Name_Get
        Else
            VerLugar = 0
            Call frmMain.Engine.Char_Pos_Get(CurrentUser.CurrentChar, map_x, map_y)
            frmMain.Label2(0).Caption = Locale_GUI_Frase(170) & ": " & CurrentUser.MapNum & ", " & map_x & ", " & map_y
        End If

        chkop(Index).value = VerLugar

    Case 5
        If NombresSimples = 1 Then
            NombresSimples = 0
        Else
            NombresSimples = 1
        End If
        
        frmMain.Engine.Engine_Label_Simple_Set
        chkop(Index).value = NombresSimples
    
    Case 6
    
        If Publicidad_Contenido = 1 Then
            Publicidad_Contenido = 0
        Else
            Publicidad_Contenido = 1
        End If
        
        chkop(Index).value = Publicidad_Contenido
    
    Case 7
    
        If CopiarDialogos = 1 Then
            CopiarDialogos = 0
        Else
            CopiarDialogos = 1
        End If

        chkop(Index).value = CopiarDialogos

    Case 9
        
        If MensajesGlobales = 1 Then
            MensajesGlobales = 0
        Else
            MensajesGlobales = 1
        End If

        If CurrentUser.Logged Then Call ClientTCP.Send_Data(Global_Option, Byte_To_String(MensajesGlobales) & Byte_To_String(MensajesFaccionarios))
        chkop(Index).value = MensajesGlobales
        
    Case 10
        
        If MsgBox(Locale_GUI_Frase(338), vbQuestion + vbYesNo, Locale_GUI_Frase(339)) = vbYes Then
            If CursoresStandar = 1 Then
                CursoresStandar = 0
            Else
                CursoresStandar = 1
            End If
            
            Call EndGame(True, True)
        End If
        
    Case 11
        
        If MensajesFaccionarios = 1 Then
            MensajesFaccionarios = 0
        Else
            MensajesFaccionarios = 1
        End If

        If CurrentUser.Logged Then Call ClientTCP.Send_Data(Global_Option, Byte_To_String(MensajesGlobales) & Byte_To_String(MensajesFaccionarios))
        chkop(Index).value = MensajesFaccionarios
        
End Select

End Sub

Private Sub cmbLanguage_Click()

If bLoading Then Exit Sub

If MsgBox(Locale_GUI_Frase(338), vbQuestion + vbYesNo, Locale_GUI_Frase(339)) = vbYes Then
    Select Case cmbLanguage.ListIndex
        Case 0
            GameLocale = "es"
        Case 1
            GameLocale = "en"
    End Select
    
    Call EndGame(True, True)
End If

End Sub

Private Sub cmdAyuda_Click()
ShellExecute Me.hwnd, "open", Chr$(34) & "http://www.imperiumao.com.ar/" & GameLocale & "/manual/" & Chr$(34), vbNullString, vbNullString, 1
End Sub

Private Sub cmdControles_Click()
frmReBind.Show vbModeless, frmOpciones
End Sub

Private Sub cmdCerrar_Click()
Me.Visible = False
End Sub

Private Sub cmdWeb_Click(Index As Integer)

Select Case Index
    Case 0
        ShellExecute Me.hwnd, "open", Chr$(34) & "http://www.imperiumao.com.ar/" & GameLocale & "/" & Chr$(34), vbNullString, vbNullString, 1
    Case 1
        ShellExecute Me.hwnd, "open", Chr$(34) & "http://www.imperiumao.com.ar/" & GameLocale & "/control.php" & Chr$(34), vbNullString, vbNullString, 1
End Select

End Sub

Private Sub Form_Load()

Call FormParser.Parse_Form(Me)

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If (Button = vbLeftButton) Then
    Call Auto_Drag(Me.hwnd)
Else
    Me.Visible = False
End If

End Sub

Public Sub AgregarIgnorado(ByVal Nick As String)

On Error Resume Next

Dim i As Long

Nick = UCase$(Nick)

For i = 0 To lstIgnore.ListCount
    If UCase$(lstIgnore.List(i)) = Nick Then
        
        Call lstIgnore.RemoveItem(i)
        
        If CurrentUser.Logged Then
            Call PrintToConsole(Nick & " " & Locale_GUI_Frase(262), 0, 0, 0, 0, 0, 0, 8)
        Else
            Call MensajeAdvertencia(Nick & " " & Locale_GUI_Frase(262))
        End If
        
        Exit Sub
    End If
Next i

lstIgnore.AddItem Nick
If CurrentUser.Logged Then Call PrintToConsole(Nick & " " & Locale_GUI_Frase(263), 0, 0, 0, 0, 0, 0, 8)

End Sub

Public Sub Init()

On Error Resume Next

Dim t() As String, i As Integer, file_name As String, tBtArr() As Byte

bLoading = True

If sMusica = CONST_DESHABILITADA Then
    chkop(0).value = 0
    chkop(0).Enabled = False
    scrMidi.value = MusicVolume
    scrMidi.Enabled = False
    chkMidi.value = DefMidi
    chkMidi.Enabled = False
    txtMidi.Text = 0
    txtMidi.Enabled = False
Else
    chkop(0).value = 1
    scrMidi.value = MusicVolume
    chkMidi.value = DefMidi
    txtMidi.Text = Sound.MusicActual
End If

If Audio = 1 Then
    chkop(1).value = 1
    chkop(2).value = FxNavega
    chkInvertir.value = IIf(InvertirSonido = True, 1, 0)
    scrVolume.value = FXVolume
Else
    chkop(1).value = 0
    chkop(2).value = FxNavega
    chkop(2).Enabled = False
    chkInvertir.value = IIf(InvertirSonido = True, 1, 0)
    chkInvertir.Enabled = False
    scrVolume.value = FXVolume
    scrVolume.Enabled = False
End If

chkop(3).value = IIf(frmMain.Engine.Engine_Label_Render_Get = True, 1, 0)
chkop(4).value = VerLugar
chkop(5).value = NombresSimples
chkop(6).value = Publicidad_Contenido
chkop(7).value = CopiarDialogos
chkop(9).value = MensajesGlobales
chkop(10).value = CursoresStandar
chkop(11).value = MensajesFaccionarios

If lstIgnore.ListCount = 0 Then
    t = Split(ListaIgnorados, "¬")
    
    lstIgnore.Clear
    
    For i = 0 To UBound(t)
        lstIgnore.AddItem t(i)
    Next i
End If

lstSkin.Clear

file_name = dir$(App.Path & "\Skins\")
Do While Len(file_name) > 0
    If Not _
        (file_name = ".") Or _
        (file_name = "..") Or _
        (Right$(file_name, 3) <> "ias") _
    Then
        If (LenB(General_Field_Read(2, file_name, "_")) = 0 And GameLocale = "es") Or (mid$(LCase$(General_Field_Read(2, file_name, "_")), 1, 2) = GameLocale) Then
            If Resource_File_Exists(App.Path & "\Skins\" & file_name, "todo.jpg") Then
                lstSkin.AddItem IIf(InStr(1, file_name, "_") = 0, mid$(file_name, 1, Len(file_name) - 4), General_Field_Read(1, file_name, "_"))
            End If
        End If
    End If
    
    file_name = dir$()
Loop

For i = 0 To lstSkin.ListCount - 1
    If lstSkin.List(i) = NombreSkin Or lstSkin.List(i) = General_Field_Read(1, NombreSkin, "_") Then
        lstSkin.Selected(i) = True
        lblSkinData.Caption = Locale_GUI_Frase(194) & ": " & General_Get_Skin_Author
        Exit For
    End If
Next i

If sMusica <> CONST_DESHABILITADA Then
    chkMidi.Enabled = True
    
    If chkMidi.value = 1 Then
        txtMidi.Enabled = False
        lblNextMidi.Enabled = False
        lblBackMidi.Enabled = False
    Else
        txtMidi.Enabled = True
        lblNextMidi.Enabled = True
        lblBackMidi.Enabled = True
    End If

Else
    chkMidi.Enabled = False
    txtMidi.Enabled = False
    lblNextMidi.Enabled = False
    lblBackMidi.Enabled = False
    scrMidi.Enabled = False
End If

If Audio = 0 Then
    scrVolume.Enabled = False
End If

If GameLocale = "en" Then
    cmbLanguage.ListIndex = 1
Else
    cmbLanguage.ListIndex = 0
End If

If Not CurrentUser.Logged Then
    Me.Show vbModeless, frmConnect
Else
    Me.Show vbModeless, frmMain
End If

bLoading = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
End Sub

Private Sub lblBackMidi_Click()

If Val(txtMidi.Text) <= 1 Then
    Beep
ElseIf Sound.Music_Load(Val(txtMidi.Text) - 1, Sound.VolumenActualMusic) Then
    txtMidi.Text = Val(txtMidi.Text) - 1
    'Sound.Music_Stop
    Sound.Music_Play
Else
    txtMidi.Text = Val(txtMidi.Text) - 1
End If

End Sub

Private Sub lblNextMidi_Click()

If sMusica = CONST_MIDI Then
    If Val(txtMidi.Text) > 70 Then
        Beep
    ElseIf Sound.Music_Load(Val(txtMidi.Text) + 1, Sound.VolumenActualMusic) Then
        txtMidi.Text = Val(txtMidi.Text) + 1
        'Sound.Music_Stop
        Sound.Music_Play
    End If
Else
    If Val(txtMidi.Text) > 70 Then
        Beep
    ElseIf Sound.Music_Load(Val(txtMidi.Text) + 1, Sound.VolumenActualMusic) Then
        'Sound.Music_Stop
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

Private Sub lstSkin_ItemCheck(Item As Integer)

Dim i As Long

For i = 0 To lstSkin.ListCount - 1
    If i <> Item Then
        lstSkin.Selected(i) = False
    Else
            
        If frmMain.LoadedSkin <> lstSkin.List(Item) Then
            
            If Not frmMain.Picture Is Nothing Then
                lstSkin.Selected(Item) = True
                
                If Not General_File_Exists(App.Path & "\Skins\" & lstSkin.List(Item) & "_" & GameLocale & ".ias", vbNormal) Then
                    NombreSkin = lstSkin.List(Item)
                Else
                    NombreSkin = lstSkin.List(Item) & "_" & GameLocale
                End If
                
                Set_Skin_Name (NombreSkin)
                frmMain.LoadedSkin = NombreSkin
                frmMain.Picture = General_Load_Skin_Picture_From_Resource_Ex("todo")
                frmMain.RecTxt.Refresh
                Call frmMain.CambiaCentro(frmMain.CentroActual)
                lblSkinData.Caption = Locale_GUI_Frase(194) & ": " & General_Get_Skin_Author
            Else
            
            End If
            
        End If
        
    End If
Next i

End Sub

Private Sub mnuAgregarIgnorado_Click()

Dim Resp As String
Resp = InputBox("Escriba el nombre del usuario que desea ignorar (también puede usar el comando /IGNORAR nick)", "Ignorar usuario")
If Resp <> vbNullString Then Call AgregarIgnorado(Resp)

End Sub

Private Sub mnuQuitarIgnorado_Click()

If lstIgnore.ListIndex = -1 Then Exit Sub
lstIgnore.RemoveItem lstIgnore.ListIndex

End Sub

Private Sub scrMidi_Change()

If sMusica <> CONST_DESHABILITADA Then
    Sound.Music_Volume_Set scrMidi.value
    Sound.VolumenActualMusicMax = scrMidi.value
    MusicVolume = Sound.VolumenActualMusicMax
End If

End Sub

Private Sub scrVolume_Change()

If Audio = 1 Then
    Sound.VolumenActual = scrVolume.value
    FXVolume = Sound.VolumenActual
End If

End Sub

Private Sub txtMidi_Change()

If sMusica = CONST_DESHABILITADA Then Exit Sub

If Val(txtMidi.Text) > 0 And (Val(txtMidi.Text) <> Sound.MusicActual) Then
    If Not Sound.Music_Load(Val(txtMidi.Text), Sound.VolumenActualMusic) Then
        txtMidi.Text = Sound.MusicActual
    Else
        'Sound.Music_Stop
        Sound.Music_Play
    End If
End If

End Sub
