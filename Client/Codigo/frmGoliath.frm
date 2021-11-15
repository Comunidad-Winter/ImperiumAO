VERSION 5.00
Begin VB.Form frmGoliath 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "$297"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4590
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
   Icon            =   "frmGoliath.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "$2"
      Height          =   345
      Left            =   120
      TabIndex        =   5
      Top             =   3090
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "$1"
      Height          =   345
      Left            =   3000
      TabIndex        =   4
      Top             =   3090
      Width           =   1455
   End
   Begin VB.TextBox txtDatos 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   4335
   End
   Begin VB.ListBox lstBanco 
      Height          =   840
      ItemData        =   "frmGoliath.frx":000C
      Left            =   90
      List            =   "frmGoliath.frx":000E
      TabIndex        =   1
      Top             =   1230
      Width           =   4395
   End
   Begin VB.Label lblDatos 
      Caption         =   "$29"
      Height          =   405
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   4335
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGoliath.frx":0010
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4395
   End
End
Attribute VB_Name = "frmGoliath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmGoliath - ImperiumAO - v1.4.5 R5
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

Private Oro As Long
Private Items As Long
Private CantTransferencia As Long
Private NombreTransferencia As String
Private EtapaTransferencia As Byte

Public Sub ParseBancoInfo(ByVal rData As String)

On Error GoTo Error_Handler

Oro = Val(General_Field_Read(1, rData, ","))
Items = Val(General_Field_Read(2, rData, ","))

If Val(Oro) > 0 And Val(Items) > 0 Then
    lblInfo.Caption = Locale_GUI_Frase(214) & " " & Items & " " & Locale_GUI_Frase(218) & " " & Oro & " " & Locale_GUI_Frase(216)
ElseIf Val(Oro) <= 0 And Val(Items) > 0 Then
    lblInfo.Caption = Locale_GUI_Frase(214) & " " & Items & " " & Locale_GUI_Frase(217)
ElseIf Val(Oro) > 0 And Val(Items) <= 0 Then
    lblInfo.Caption = Locale_GUI_Frase(219) & " " & Oro & " " & Locale_GUI_Frase(216)
ElseIf Val(Oro) <= 0 And Val(Items) <= 0 Then
    lblInfo.Caption = Locale_GUI_Frase(215)
End If

Me.Show vbModeless, frmMain

Exit Sub

Error_Handler:
    'Error vite'

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()

Select Case lstBanco.ListIndex
    Case 0, -1 'Depositar
    
        'Negativos y ceros
        If (Val(txtDatos.Text) <= 0 And (UCase$(txtDatos.Text) <> "TODO")) Then lblDatos.Caption = Locale_GUI_Frase(220)
    
        If Val(txtDatos.Text) <= CurrentUser.UserGLD Or UCase$(txtDatos.Text) = "TODO" Then
            Call ClientTCP.Send_Data_Command(cmdDepositar, IIf(Val(txtDatos.Text) > 0, Val(txtDatos.Text), CurrentUser.UserGLD))
            Unload Me
        Else
            lblDatos.Caption = Locale_GUI_Frase(221)
        End If
    Case 1 'Retirar
    
        'Negativos y ceros
        If (Val(txtDatos.Text) <= 0 And (UCase$(txtDatos.Text) <> "TODO")) Then lblDatos.Caption = Locale_GUI_Frase(220)
    
        If Val(txtDatos.Text) <= Oro Or UCase$(txtDatos.Text) = "TODO" Then
            Call ClientTCP.Send_Data_Command(cmdRetirar, IIf(Val(txtDatos.Text) > 0, Val(txtDatos.Text), Oro))
            Call Sound.Sound_Play(SND_RETIRARORO)
            Unload Me
        Else
            lblDatos.Caption = Locale_GUI_Frase(221)
        End If
    Case 2 'Bóveda
        Unload Me
    Case 3 'Transferir - Destino - Cantidad
        If EtapaTransferencia = 0 Then
        
            'Negativos y ceros
            If Val(txtDatos.Text) <= 0 Then
                lblDatos.Caption = Locale_GUI_Frase(221)
                txtDatos.Text = vbNullString
                Exit Sub
            End If
            
            If Val(txtDatos.Text) <= Oro Then
                CantTransferencia = Val(txtDatos.Text)
                lblDatos.Caption = Locale_GUI_Frase(222) & " " & CantTransferencia & " " & Locale_GUI_Frase(223)
                EtapaTransferencia = 1
                txtDatos.Text = vbNullString
            Else
                lblDatos.Caption = Locale_GUI_Frase(221)
                txtDatos.Text = vbNullString
            End If
        ElseIf EtapaTransferencia = 1 Then
            If LenB(txtDatos.Text) > 0 Then
                NombreTransferencia = txtDatos.Text
                lblDatos.Caption = "Se transferirán " & CantTransferencia & " monedas de oro a " & NombreTransferencia & ". Si es correcto, presione aceptar."
                EtapaTransferencia = 2
            Else
                lblDatos.Caption = Locale_GUI_Frase(224)
                txtDatos.Text = vbNullString
            End If
        ElseIf EtapaTransferencia = 2 Then
            Call ClientTCP.Send_Data(Gold_Transfer, Long_To_String(CantTransferencia) & NombreTransferencia)
            Unload Me
        End If
End Select

End Sub

Private Sub Form_Load()

Call FormParser.Parse_Form(Me)

lstBanco.AddItem Locale_GUI_Frase(298)
lstBanco.AddItem Locale_GUI_Frase(299)
lstBanco.AddItem Locale_GUI_Frase(300)
lstBanco.AddItem Locale_GUI_Frase(301)

End Sub

Private Sub lstBanco_Click()

Select Case lstBanco.ListIndex
    Case 0 'Depositar
        lblDatos.Caption = Locale_GUI_Frase(225)
    Case 1 'Retirar
        lblDatos.Caption = Locale_GUI_Frase(226)
    Case 2 'Bóveda
        Call ClientTCP.Send_Data(Bank_Init)
        Unload Me
    Case 3 'Transferir
        EtapaTransferencia = 0
        lblDatos.Caption = Locale_GUI_Frase(227)
End Select

End Sub
