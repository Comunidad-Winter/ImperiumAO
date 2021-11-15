VERSION 5.00
Begin VB.Form frmInicio 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "ImperiumAO"
   ClientHeight    =   7110
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   4920
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
   Icon            =   "frmInicio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmInicio.frx":0ECA
   ScaleHeight     =   7110
   ScaleWidth      =   4920
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "v1.3"
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
      Height          =   255
      Left            =   4440
      TabIndex        =   5
      Top             =   6840
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1140
      TabIndex        =   4
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Image Image2 
      Height          =   810
      Index           =   3
      Left            =   1095
      Top             =   3390
      Width           =   2760
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   1140
      TabIndex        =   3
      Top             =   2775
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1140
      TabIndex        =   2
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Image Image2 
      Height          =   810
      Index           =   2
      Left            =   1095
      Top             =   2550
      Width           =   2760
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   1155
      TabIndex        =   1
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Image Image2 
      Height          =   810
      Index           =   1
      Left            =   1095
      MousePointer    =   99  'Custom
      Top             =   1710
      Width           =   2760
   End
   Begin VB.Image Image2 
      Height          =   810
      Index           =   0
      Left            =   1095
      Top             =   880
      Width           =   2760
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   4680
   End
End
Attribute VB_Name = "frmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum tPantalla
    PRINCIPAL = 0
    CONFIGURACION = 1
    INFORMACION = 2
    
    TOTAL_PANTALLAS
End Enum

Private Type tPantInfo
    Titulo As String
    Botones(0 To 3) As String
End Type

Dim UltPos As Long
Dim PantAct As tPantalla
Dim Pantallas(0 To TOTAL_PANTALLAS - 1) As tPantInfo

Private Sub Image2_Click(Index As Integer)

On Error Resume Next

Call PlayWaveAPI(App.Path & "\WAV\CLICK.wav")

Select Case PantAct
Case PRINCIPAL
    Select Case Index
        Case 0
            frmUpdater.Show
            IniciarUpdates
            Call StartGame
        Case 1
            Call CambiaPantalla(CONFIGURACION)
        Case 2
            Call CambiaPantalla(INFORMACION)
        Case 3
            Unload Me
    End Select
Case CONFIGURACION
    Select Case Index
        Case 0
            'ShellExecute Me.hWnd, "open", App.Path & "\AoSetup.exe", "", "", 1
            MsgBox ("Lo sentimos aún se está trabajando en esta opción.")
        Case 1
            If MsgBox("Se registrarán las dependencias requeridas por el juego a fin de evitar errores en tiempo de ejecución ¿Continuar?", vbYesNo, "Solución de problemas") = vbYes Then
                Call RegistrarTodasLasDependencias
                MsgBox ("Todas las dependencias han sido registradas correctamente")
            End If
        Case 2
            ShellExecute Me.hWnd, "open", App.Path & "\Documentos\Notas de version.txt", "", "", 1
        Case 3
            Call CambiaPantalla(PRINCIPAL)
    End Select
Case INFORMACION
    Select Case Index
        Case 0
            ShellExecute Me.hWnd, "open", "http://www.imperiumao.com.ar", "", "", 1
        Case 1
            ShellExecute Me.hWnd, "open", "http://foro.imperiumao.com.ar", "", "", 1
        Case 2
            ShellExecute Me.hWnd, "open", "http://www.imperiumao.com.ar/manual/", "", "", 1
        Case 3
            Call CambiaPantalla(PRINCIPAL)
    End Select
End Select

End Sub

Private Sub Form_Deactivate()
On Error Resume Next
Me.Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If UltPos >= 0 Then
    Image2(UltPos).Picture = LoadPicture(App.Path & "\Interface\BotonInicio.jpg")
    UltPos = -1
End If

End Sub

Private Sub Form_Load()

Image2(0).Picture = LoadPicture(App.Path & "\Interface\BotonInicio.jpg")
Image2(1).Picture = LoadPicture(App.Path & "\Interface\BotonInicio.jpg")
Image2(2).Picture = LoadPicture(App.Path & "\Interface\BotonInicio.jpg")
Image2(3).Picture = LoadPicture(App.Path & "\Interface\BotonInicio.jpg")

With Pantallas(PRINCIPAL)
    .Titulo = ""
    .Botones(0) = "Jugar"
    .Botones(1) = "Configuración"
    .Botones(2) = "Información"
    .Botones(3) = "Salir"
End With

With Pantallas(CONFIGURACION)
    .Titulo = "Configuración"
    .Botones(0) = "Ejecutar el Setup"
    .Botones(1) = "Solución de problemas"
    .Botones(2) = "Notas de versión"
    .Botones(3) = "Volver"
End With

With Pantallas(INFORMACION)
    .Titulo = "Información"
    .Botones(0) = "Visitar el sitio oficial"
    .Botones(1) = "Foros de discusión"
    .Botones(2) = "Manual"
    .Botones(3) = "Volver"
End With

Call CambiaPantalla(PRINCIPAL)

End Sub

Private Sub Image2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If UltPos <> Index Then
    If UltPos >= 0 Then Image2(UltPos).Picture = LoadPicture(App.Path & "\Interface\BotonInicio.jpg")
    Image2(Index).Picture = LoadPicture(App.Path & "\Interface\BotonInicioApretado.jpg")
    UltPos = Index
End If

End Sub

Private Sub Label2_Click(Index As Integer)
Call Image2_Click(Index)
End Sub

Private Sub Label2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Image2_MouseMove(Index, Button, Shift, X, Y)
End Sub

Private Sub CambiaPantalla(Nueva As tPantalla)
Dim i As Long

PantAct = Nueva

For i = LBound(Pantallas(PantAct).Botones) To UBound(Pantallas(PantAct).Botones)
    Label2(i).Caption = Pantallas(PantAct).Botones(i)
Next i
Label1.Caption = Pantallas(PantAct).Titulo

End Sub

Private Sub RegistrarTodasLasDependencias()

On Error Resume Next

Call FileCopy(App.Path & "\MSWINSCK.OCX", fGetWinDir & "\SYSTEM32\MSWINSCK.OCX")
Call ShellExecute(frmInicio.hWnd, "open", "regsvr32", "MSWINSCK.OCX" & " /s", "", 1)

Call FileCopy(App.Path & "\RICHTX32.OCX", fGetWinDir & "\SYSTEM32\RICHTX32.OCX")
Call ShellExecute(frmInicio.hWnd, "open", "regsvr32", "RICHTX32.OCX" & " /s", "", 1)

Call FileCopy(App.Path & "\MSINET.OCX", fGetWinDir & "\SYSTEM32\MSINET.OCX")
Call ShellExecute(frmInicio.hWnd, "open", "regsvr32", "MSINET.OCX" & " /s", "", 1)

Call FileCopy(App.Path & "\Mscomctl.ocx", fGetWinDir & "\SYSTEM32\Mscomctl.ocx")
Call ShellExecute(frmInicio.hWnd, "open", "regsvr32", "Mscomctl.ocx" & " /s", "", 1)

Call FileCopy(App.Path & "\msstdfmt.dll", fGetWinDir & "\SYSTEM32\msstdfmt.dll")
Call ShellExecute(frmInicio.hWnd, "open", "regsvr32", "msstdfmt.dll" & " /s", "", 1)

End Sub
