VERSION 5.00
Begin VB.Form fmrGMAyuda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Llamando a un Game Master"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2558
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar consulta"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   2558
      Width           =   1695
   End
   Begin VB.TextBox txtMotivo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   233
      TabIndex        =   1
      Top             =   840
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"fmrGMAyuda.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "fmrGMAyuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If txtMotivo.Text = "" Then
    MsgBox ("Debes escribir el motivo de tu consulta")
    Exit Sub
Else
    SendData ("/GM " & txtMotivo.Text)
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
