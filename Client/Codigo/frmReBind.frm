VERSION 5.00
Begin VB.Form frmReBind 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "$396"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7980
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReBind.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txConfig 
      Enabled         =   0   'False
      Height          =   315
      Index           =   20
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   47
      Text            =   "Bloq. Num."
      Top             =   3270
      Width           =   2415
   End
   Begin VB.TextBox txtMSens 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7470
      TabIndex        =   46
      Text            =   "20"
      Top             =   3960
      Width           =   375
   End
   Begin VB.HScrollBar scrSens 
      Height          =   345
      LargeChange     =   15
      Left            =   5400
      Max             =   20
      Min             =   1
      TabIndex        =   44
      Top             =   3960
      Value           =   1
      Width           =   2025
   End
   Begin VB.TextBox txConfig 
      Enabled         =   0   'False
      Height          =   315
      Index           =   19
      Left            =   2730
      Locked          =   -1  'True
      TabIndex        =   42
      Text            =   "ALT1"
      Top             =   5460
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Enabled         =   0   'False
      Height          =   315
      Index           =   18
      Left            =   2730
      Locked          =   -1  'True
      TabIndex        =   40
      Text            =   "*"
      Top             =   4710
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Enabled         =   0   'False
      Height          =   315
      Index           =   17
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   38
      Text            =   "Impr. Pant."
      Top             =   5460
      Width           =   2415
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "$420"
      Height          =   315
      Left            =   5430
      TabIndex        =   37
      Top             =   4830
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   16
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   2550
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   15
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   1830
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   14
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   1110
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   13
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   390
      Width           =   2415
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "$25"
      Height          =   315
      Index           =   2
      Left            =   5430
      TabIndex        =   30
      Top             =   5490
      Width           =   2415
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "$421"
      Height          =   315
      Index           =   1
      Left            =   5430
      TabIndex        =   34
      Top             =   4500
      Width           =   2415
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "$419"
      Height          =   315
      Index           =   0
      Left            =   5430
      TabIndex        =   32
      Top             =   5160
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   12
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   3990
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   11
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   3270
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   10
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   2550
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   9
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   1830
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   8
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1110
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   7
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   390
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   6
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   4710
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   5
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   3990
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   4
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   3270
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   3
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2550
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   2
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1830
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   1
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1110
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   390
      Width           =   2415
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$417"
      Height          =   195
      Index           =   21
      Left            =   5400
      TabIndex        =   48
      Top             =   2970
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$418"
      Height          =   195
      Index           =   20
      Left            =   5400
      TabIndex        =   45
      Top             =   3660
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$412"
      Height          =   195
      Index           =   19
      Left            =   2730
      TabIndex        =   43
      Top             =   5160
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$411"
      Height          =   195
      Index           =   18
      Left            =   2730
      TabIndex        =   41
      Top             =   4410
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$404"
      Height          =   195
      Index           =   17
      Left            =   120
      TabIndex        =   39
      Top             =   5160
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$416"
      Height          =   195
      Index           =   16
      Left            =   5400
      TabIndex        =   36
      Top             =   2250
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$415"
      Height          =   195
      Index           =   15
      Left            =   5400
      TabIndex        =   35
      Top             =   1530
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$414"
      Height          =   195
      Index           =   14
      Left            =   5400
      TabIndex        =   33
      Top             =   810
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$413"
      Height          =   195
      Index           =   13
      Left            =   5400
      TabIndex        =   31
      Top             =   90
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$410"
      Height          =   195
      Index           =   12
      Left            =   2760
      TabIndex        =   25
      Top             =   3690
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$409"
      Height          =   195
      Index           =   11
      Left            =   2760
      TabIndex        =   23
      Top             =   2970
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$408"
      Height          =   195
      Index           =   10
      Left            =   2760
      TabIndex        =   21
      Top             =   2250
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$407"
      Height          =   195
      Index           =   9
      Left            =   2760
      TabIndex        =   19
      Top             =   1530
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$406"
      Height          =   195
      Index           =   8
      Left            =   2760
      TabIndex        =   17
      Top             =   810
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$405"
      Height          =   195
      Index           =   7
      Left            =   2760
      TabIndex        =   15
      Top             =   90
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$403"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   13
      Top             =   4410
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$402"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   11
      Top             =   3690
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$401"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   2970
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$400"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   2250
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$399"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1530
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$398"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   810
      Width           =   360
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$397"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   90
      Width           =   360
   End
End
Attribute VB_Name = "frmReBind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TempVars(0 To 16) As Integer

Private Sub cmdAccion_Click(Index As Integer)

Dim i As Integer
Dim bCambio As Boolean
Dim Resultado As VbMsgBoxResult

Select Case Index
    
    Case 0
        Call GuardaConfigEnVariables
    Case 1
        Call LoadDefaultBinds
        Call CargaConfigEnForm
    Case 2
    
        For i = 1 To NUMBINDS
            If TempVars(i - 1) <> BindKeys(i).KeyCode Then
                bCambio = True
                Exit For
            End If
        Next

        If bCambio Then
            Resultado = MsgBox(Locale_GUI_Frase(341), vbQuestion + vbYesNoCancel, Locale_GUI_Frase(342))
            If Resultado = vbYes Then Call GuardaConfigEnVariables
        End If
        
        If Resultado <> vbCancel Then Unload Me

End Select

End Sub

Private Sub cmdReset_Click()

Dim i As Integer

If MsgBox(Locale_GUI_Frase(343), vbYesNo + vbQuestion, Locale_GUI_Frase(344)) = vbYes Then
    For i = 1 To NUMBOTONES
        MacroKeys(i).TipoAccion = 0
        MacroKeys(i).hlist = 0
        MacroKeys(i).SendString = vbNullString
        MacroKeys(i).invslot = 0
    Next i
End If

End Sub

Private Sub GuardaConfigEnVariables()

Dim i As Integer

For i = 1 To NUMBINDS
    BindKeys(i).name = txConfig(i - 1).Text
    BindKeys(i).KeyCode = TempVars(i - 1)
    BindKeys(i).VirtualKey = MapVirtualKey(BindKeys(i).KeyCode, 0)
    
    If BindKeys(i).VirtualKey = DIK_NUMPAD4 Then
        BindKeys(i).VirtualKey = DIK_LEFTARROW
    ElseIf BindKeys(i).VirtualKey = DIK_NUMPAD6 Then
        BindKeys(i).VirtualKey = DIK_RIGHTARROW
    ElseIf BindKeys(i).VirtualKey = DIK_NUMPAD8 Then
        BindKeys(i).VirtualKey = DIK_UPARROW
    ElseIf BindKeys(i).VirtualKey = DIK_NUMPAD2 Then
        BindKeys(i).VirtualKey = DIK_DOWNARROW
    End If
    
Next

End Sub

Private Sub CargaConfigEnForm()

Dim i As Integer

For i = 1 To NUMBINDS
    txConfig(i - 1).Text = BindKeys(i).name
    TempVars(i - 1) = BindKeys(i).KeyCode
Next

scrSens.value = MouseS
txtMSens.Text = MouseS

End Sub

Private Sub Form_Load()
Call CargaConfigEnForm
Call FormParser.Parse_Form(Me)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Dim i As Integer
Dim bCambio As Boolean
Dim Resultado As VbMsgBoxResult

For i = 1 To NUMBINDS
    If TempVars(i - 1) <> BindKeys(i).KeyCode Then
        bCambio = True
        Exit For
    End If
Next

If bCambio Then
    Resultado = MsgBox(Locale_GUI_Frase(341), vbQuestion + vbYesNoCancel, Locale_GUI_Frase(342))
    If Resultado = vbYes Then Call GuardaConfigEnVariables
End If

If Resultado = vbCancel Then Cancel = 1

End Sub

Private Sub scrSens_Change()

MouseS = scrSens.value
Call General_Set_Mouse_Speed(MouseS)
txtMSens.Text = scrSens.value

End Sub

Private Sub txConfig_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim name As String
name = txConfig(Index).Text

If KeyCode > 0 Then
    
    If AlreadyBinded(KeyCode) Then
        txConfig(Index).ForeColor = vbRed
        Exit Sub
    Else
        txConfig(Index).ForeColor = vbBlack
    End If
    
    If KeyCode = vbKeyShift Then
        name = "Shift"
    ElseIf KeyCode = vbKeyLeft Then
        name = "Flecha Izquierda"
    ElseIf KeyCode = vbKeyRight Then
        name = "Flecha Derecha"
    ElseIf KeyCode = vbKeyDown Then
        name = "Flecha Abajo"
    ElseIf KeyCode = vbKeyUp Then
        name = "Flecha Arriba"
    ElseIf KeyCode = vbKeyControl Then
        name = "Control"
    ElseIf KeyCode = vbKeyPageDown Then
        name = "Page Down"
    ElseIf KeyCode = vbKeyPageUp Then
        name = "Page Up"
    ElseIf KeyCode = vbKeySeparator Then 'Enter teclado numerico
        name = "Intro"
    ElseIf KeyCode = vbKeySpace Then
        name = "Barra Espaciadora"
    ElseIf KeyCode = vbKeyDelete Then
        name = "Delete"
    ElseIf KeyCode = vbKeyEnd Then
        name = "Fin"
    ElseIf KeyCode = vbKeyHome Then
        name = "Inicio"
    ElseIf KeyCode = vbKeyInsert Then
        name = "Insert"
    Else
        name = Chr$(KeyCode)
    End If
    
    Call Change_TempKey(Index, KeyCode, name)

End If

End Sub

Private Sub Change_TempKey(ByVal Index As Integer, ByVal KeyCode As Integer, ByVal name As String)
TempVars(Index) = KeyCode
txConfig(Index).Text = name
End Sub

Function AlreadyBinded(KeyCode As Integer) As Boolean

Dim i As Integer

If (KeyCode >= vbKeyF1 And KeyCode <= vbKeyF12) Or (KeyCode = 44) Or (KeyCode = 106) Then
    AlreadyBinded = True
    Exit Function
End If

For i = 1 To NUMBINDS
    If (TempVars(i - 1) = KeyCode) Then
        AlreadyBinded = True
        Exit Function
    End If
Next i

End Function
