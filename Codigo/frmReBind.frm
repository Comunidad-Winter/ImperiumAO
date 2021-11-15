VERSION 5.00
Begin VB.Form frmReBind 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuraración de controles"
   ClientHeight    =   5160
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txConfig 
      Enabled         =   0   'False
      Height          =   315
      Index           =   18
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   40
      Text            =   "*"
      Top             =   3270
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Enabled         =   0   'False
      Height          =   315
      Index           =   17
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   38
      Text            =   "Impr. Pantalla"
      Top             =   2550
      Width           =   2415
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reconfigurar macros"
      Height          =   315
      Left            =   5400
      TabIndex        =   37
      Top             =   4740
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   16
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   1830
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   15
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   1110
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   14
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   390
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   13
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   4710
      Width           =   2415
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   5400
      TabIndex        =   30
      Top             =   3750
      Width           =   2415
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "&Cargar defaults"
      Height          =   315
      Index           =   1
      Left            =   5400
      TabIndex        =   34
      Top             =   4410
      Width           =   2415
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "&Guardar"
      Height          =   315
      Index           =   0
      Left            =   5400
      TabIndex        =   32
      Top             =   4080
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
      Caption         =   "Ver estadísticas del engine"
      Height          =   195
      Index           =   18
      Left            =   5400
      TabIndex        =   41
      Top             =   2970
      Width           =   1905
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tomar screenshot"
      Height          =   195
      Index           =   17
      Left            =   5400
      TabIndex        =   39
      Top             =   2250
      Width           =   1290
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Moverse hacia la derecha"
      Height          =   195
      Index           =   16
      Left            =   5400
      TabIndex        =   36
      Top             =   1530
      Width           =   1830
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Moverse hacia la izquierda"
      Height          =   195
      Index           =   15
      Left            =   5400
      TabIndex        =   35
      Top             =   810
      Width           =   1890
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Moverse hacia abajo"
      Height          =   195
      Index           =   14
      Left            =   5400
      TabIndex        =   33
      Top             =   90
      Width           =   1485
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Moverse hacia arriba"
      Height          =   195
      Index           =   13
      Left            =   2760
      TabIndex        =   31
      Top             =   4410
      Width           =   1500
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Modo rol"
      Height          =   195
      Index           =   12
      Left            =   2760
      TabIndex        =   25
      Top             =   3690
      Width           =   615
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Modo combate"
      Height          =   195
      Index           =   11
      Left            =   2760
      TabIndex        =   23
      Top             =   2970
      Width           =   1050
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ocultarse"
      Height          =   195
      Index           =   10
      Left            =   2760
      TabIndex        =   21
      Top             =   2250
      Width           =   690
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Actualizar posición"
      Height          =   195
      Index           =   9
      Left            =   2760
      TabIndex        =   19
      Top             =   1530
      Width           =   1320
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Robar"
      Height          =   195
      Index           =   8
      Left            =   2760
      TabIndex        =   17
      Top             =   810
      Width           =   435
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Domar"
      Height          =   195
      Index           =   7
      Left            =   2760
      TabIndex        =   15
      Top             =   90
      Width           =   465
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mostrar / Ocultar Nicknames"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   13
      Top             =   4410
      Width           =   2025
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Activar / Desactivar Seguro"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   11
      Top             =   3690
      Width           =   1980
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Equipar objeto"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   2970
      Width           =   1050
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usar objeto"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   2250
      Width           =   840
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tirar objeto"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1530
      Width           =   840
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tomar objeto"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   810
      Width           =   960
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Atacar"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   90
      Width           =   480
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
            Resultado = MsgBox("Realizo cambios en la configuración ¿desea guardar antes de salir?", vbQuestion + vbYesNoCancel, "Guardar cambios")
            If Resultado = vbYes Then Call GuardaConfigEnVariables
        End If
        
        If Resultado <> vbCancel Then Unload Me

End Select

End Sub

Private Sub cmdReset_Click()

Dim i As Integer

If MsgBox("¿Esta seguro que desea borrar toda la configuración actual de las teclas?", vbYesNo + vbQuestion, "Borrar macros") = vbYes Then
    For i = 1 To NUMBOTONES
        MacroKeys(i).TipoAccion = 0
        MacroKeys(i).hlist = 0
        MacroKeys(i).SendString = ""
        MacroKeys(i).invslot = 0
    Next i
End If

End Sub

Private Sub GuardaConfigEnVariables()

Dim i As Integer

For i = 1 To NUMBINDS
    BindKeys(i).Name = txConfig(i - 1).Text
    BindKeys(i).KeyCode = TempVars(i - 1)
Next

End Sub

Private Sub CargaConfigEnForm()

Dim i As Integer

For i = 1 To NUMBINDS
    txConfig(i - 1).Text = BindKeys(i).Name
    TempVars(i - 1) = BindKeys(i).KeyCode
Next

End Sub

Private Sub Form_Load()
Call CargaConfigEnForm
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
    Resultado = MsgBox("Realizo cambios en la configuración ¿desea guardar antes de salir?", vbQuestion + vbYesNoCancel, "Guardar cambios")
    If Resultado = vbYes Then Call GuardaConfigEnVariables
End If

If Resultado = vbCancel Then Cancel = 1

End Sub

Private Sub txConfig_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim Name As String
Name = txConfig(Index).Text

If KeyCode > 0 Then
    
    If AlreadyBinded(KeyCode) Then
        Beep
        txConfig(Index).ForeColor = vbRed
        Exit Sub
    End If
    
    If KeyCode = vbKeyShift Then
        Name = "Shift"
    ElseIf KeyCode = vbKeyLeft Then
        Name = "Flecha Izquierda"
    ElseIf KeyCode = vbKeyRight Then
        Name = "Flecha Derecha"
    ElseIf KeyCode = vbKeyDown Then
        Name = "Flecha Abajo"
    ElseIf KeyCode = vbKeyUp Then
        Name = "Flecha Arriba"
    ElseIf KeyCode = vbKeyControl Then
        Name = "Control"
    ElseIf KeyCode = vbKeyPageDown Then
        Name = "Page Down"
    ElseIf KeyCode = vbKeyPageUp Then
        Name = "Page Up"
    ElseIf KeyCode = vbKeySeparator Then 'Enter teclado numerico
        Name = "Intro"
    ElseIf KeyCode = vbKeySpace Then
        Name = "Barra Espaciadora"
    ElseIf KeyCode = vbKeyDelete Then
        Name = "Delete"
    ElseIf KeyCode = vbKeyEnd Then
        Name = "Fin"
    ElseIf KeyCode = vbKeyHome Then
        Name = "Inicio"
    ElseIf KeyCode = vbKeyInsert Then
        Name = "Insert"
    Else
        Name = Chr(KeyCode)
    End If
    
    Call Change_TempKey(Index, KeyCode, Name)

End If

End Sub

Sub Change_TempKey(Index As Integer, KeyCode As Integer, Name As String)
TempVars(Index) = KeyCode
txConfig(Index).Text = Name
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
