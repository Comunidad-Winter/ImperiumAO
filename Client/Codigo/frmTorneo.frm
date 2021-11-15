VERSION 5.00
Begin VB.Form frmTorneo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "$425"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6585
   Icon            =   "frmTorneo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frameTor 
      Caption         =   "$99"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5265
      Index           =   1
      Left            =   3510
      TabIndex        =   5
      Top             =   60
      Width           =   3015
      Begin VB.CommandButton cmdTorneo 
         Caption         =   "$104"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   4860
         Width           =   2775
      End
      Begin VB.TextBox txtTor 
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
         Index           =   4
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   3210
         Width           =   2775
      End
      Begin VB.TextBox txtTor 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   17
         Top             =   2610
         Width           =   975
      End
      Begin VB.CheckBox chkRestr 
         Caption         =   "$105"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2100
         Width           =   2715
      End
      Begin VB.TextBox txtTor 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   120
         MaxLength       =   2
         TabIndex        =   14
         Top             =   2610
         Width           =   975
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "$107"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1740
         Width           =   1335
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "$106"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1740
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox txtTor 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   90
         MaxLength       =   2
         TabIndex        =   9
         Top             =   1110
         Width           =   2775
      End
      Begin VB.TextBox txtTor 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   90
         MaxLength       =   6
         TabIndex        =   7
         Top             =   480
         Width           =   2805
      End
      Begin VB.Label lblTor 
         BackStyle       =   0  'Transparent
         Caption         =   "$39"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   120
         TabIndex        =   18
         Top             =   2970
         Width           =   2745
      End
      Begin VB.Label lblTor 
         BackStyle       =   0  'Transparent
         Caption         =   "$112"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   1200
         TabIndex        =   16
         Top             =   2370
         Width           =   975
      End
      Begin VB.Label lblTor 
         BackStyle       =   0  'Transparent
         Caption         =   "$111"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   2370
         Width           =   975
      End
      Begin VB.Label lblTor 
         BackStyle       =   0  'Transparent
         Caption         =   "$108"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   90
         TabIndex        =   10
         Top             =   1500
         Width           =   3585
      End
      Begin VB.Label lblTor 
         BackStyle       =   0  'Transparent
         Caption         =   "$109"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   90
         TabIndex        =   8
         Top             =   870
         Width           =   2745
      End
      Begin VB.Label lblTor 
         BackStyle       =   0  'Transparent
         Caption         =   "$110"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   6
         Top             =   240
         Width           =   2805
      End
   End
   Begin VB.Frame frameTor 
      Caption         =   "$100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5265
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3375
      Begin VB.CommandButton cmdTorneo 
         Caption         =   "$103"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   150
         TabIndex        =   4
         Top             =   4830
         Width           =   3105
      End
      Begin VB.CommandButton cmdTorneo 
         Caption         =   "$102"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   150
         TabIndex        =   3
         Top             =   4470
         Width           =   3105
      End
      Begin VB.CommandButton cmdTorneo 
         Caption         =   "$101"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   150
         TabIndex        =   2
         Top             =   4110
         Width           =   3105
      End
      Begin VB.ListBox lstInsc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3765
         Left            =   150
         TabIndex        =   1
         Top             =   240
         Width           =   3075
      End
   End
End
Attribute VB_Name = "frmTorneo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Parse_Torneo_Info(ByVal Cadena As String)

Dim Identificador As Byte, i As Long, tmpStr As String

Identificador = String_To_Byte(Cadena, 1)

Select Case Identificador
    Case 1 'Puede crear...
        cmdTorneo(3).Caption = Locale_GUI_Frase(104)
        
        cmdTorneo(0).Enabled = False
        cmdTorneo(1).Enabled = False
        cmdTorneo(2).Enabled = False
        cmdTorneo(3).Enabled = True
        
        optTipo(0).Enabled = True
        optTipo(1).Enabled = True
        
        txtTor(0).Enabled = True
        txtTor(1).Enabled = True
        txtTor(4).Enabled = True
        
        chkRestr.value = 0
        chkRestr.Enabled = True
        txtTor(2).Enabled = False
        txtTor(3).Enabled = False
        
        Me.Show vbModeless, frmMain
        
    Case 2 'Puede inscribirse (torneo en inscripción)
        cmdTorneo(0).Enabled = False
        cmdTorneo(1).Enabled = False
        cmdTorneo(2).Enabled = True
        cmdTorneo(3).Enabled = False
        
        optTipo(0).Enabled = False
        optTipo(1).Enabled = False
        
        If String_To_Byte(Cadena, 2) = 1 Then
            optTipo(0).value = False
            optTipo(1).value = True
        Else
            optTipo(1).value = False
            optTipo(0).value = True
        End If
        
        txtTor(0).Enabled = False
        txtTor(0).Text = CStr(String_To_Long(Cadena, 3))
        
        txtTor(1).Enabled = False
        txtTor(1).Text = CStr(String_To_Integer(Cadena, 7))
        
        'CStr(String_To_Long(Cadena, 9)) -> tiempo restante
        
        chkRestr.value = CInt(String_To_Byte(Cadena, 13))
        chkRestr.Enabled = False
        
        txtTor(2).Enabled = False
        txtTor(3).Enabled = False
        
        If chkRestr.value Then
            txtTor(2).Text = CStr(String_To_Byte(Cadena, 14))
            txtTor(3).Text = CStr(String_To_Byte(Cadena, 15))
            tmpStr = mid$(Cadena, 16)
        Else
            txtTor(2).Text = vbNullString
            txtTor(3).Text = vbNullString
            tmpStr = mid$(Cadena, 14)
        End If
                
        txtTor(4).Enabled = False
        txtTor(4).Text = General_Field_Read(1, tmpStr, "¬")
        
        tmpStr = General_Field_Read(2, tmpStr, "¬")
        
        For i = 1 To General_Field_Count(tmpStr, 44)
            lstInsc.AddItem General_Field_Read(i, tmpStr, ",")
        Next i
        
        Me.Show vbModeless, frmMain
        
    Case 3 'Es lider de un torneo...
        cmdTorneo(0).Enabled = True
        cmdTorneo(1).Enabled = True
        cmdTorneo(2).Enabled = False
        cmdTorneo(3).Enabled = True
        cmdTorneo(3).Caption = Locale_GUI_Frase(174)
        
        optTipo(0).Enabled = False
        optTipo(1).Enabled = False
        
        If String_To_Byte(Cadena, 2) = 1 Then
            optTipo(0).value = False
            optTipo(1).value = True
        Else
            optTipo(1).value = False
            optTipo(0).value = True
        End If
        
        txtTor(0).Enabled = False
        txtTor(0).Text = CStr(String_To_Long(Cadena, 3))
        
        txtTor(1).Enabled = False
        txtTor(1).Text = CStr(String_To_Integer(Cadena, 7))
        
        'CStr(String_To_Long(Cadena, 9)) -> tiempo restante
        
        chkRestr.value = CInt(String_To_Byte(Cadena, 13))
        chkRestr.Enabled = False
        
        txtTor(2).Enabled = False
        txtTor(3).Enabled = False
        
        If chkRestr.value Then
            txtTor(2).Text = CStr(String_To_Byte(Cadena, 14))
            txtTor(3).Text = CStr(String_To_Byte(Cadena, 15))
            tmpStr = mid$(Cadena, 16)
        Else
            txtTor(2).Text = vbNullString
            txtTor(3).Text = vbNullString
            tmpStr = mid$(Cadena, 14)
        End If
                
        txtTor(4).Enabled = False
        txtTor(4).Text = General_Field_Read(1, tmpStr, "¬")
        
        tmpStr = General_Field_Read(2, tmpStr, "¬")
        
        For i = 1 To General_Field_Count(tmpStr, 44)
            lstInsc.AddItem General_Field_Read(i, tmpStr, ",")
        Next i
        
        Me.Show vbModeless, frmMain
        
End Select

End Sub

Private Sub chkRestr_Click()

If chkRestr.value Then
    txtTor(2).Enabled = True
    txtTor(3).Enabled = True
Else
    txtTor(2).Enabled = False
    txtTor(3).Enabled = False
End If

End Sub

Private Sub cmdTorneo_Click(Index As Integer)

Select Case Index
    Case 0 'Expulsar
        If lstInsc.ListIndex = -1 Then Exit Sub
        Call ClientTCP.Send_Data(Challenge_Main_Cl, Byte_To_String(1) & lstInsc.List(lstInsc.ListIndex))
        Call lstInsc.RemoveItem(lstInsc.ListIndex)
    Case 1 'Ver info
        If lstInsc.ListIndex = -1 Then Exit Sub
        Call ClientTCP.Send_Data(Challenge_Main_Cl, Byte_To_String(2) & lstInsc.List(lstInsc.ListIndex))
    Case 2 'Inscripción
        Call ClientTCP.Send_Data(Challenge_Main_Cl, Byte_To_String(3) & lstInsc.List(lstInsc.ListIndex))
        Unload Me
    Case 3 'Crear o iniciar
        'Crear ?
        If chkRestr.Enabled Then
            If Not Check_Data Then Exit Sub
            
            If chkRestr.value Then
                Call ClientTCP.Send_Data(Challenge_Main_Cl, Byte_To_String(4) & Long_To_String(CLng(txtTor(0).Text)) & Integer_To_String(CInt(txtTor(1).Text)) & Byte_To_String(CByte(txtTor(2).Text)) & Byte_To_String(CByte(txtTor(3).Text)) & Byte_To_String(IIf(optTipo(0).value, 1, 0)) & txtTor(4).Text)
            Else
                Call ClientTCP.Send_Data(Challenge_Main_Cl, Byte_To_String(5) & Long_To_String(CLng(txtTor(0).Text)) & Integer_To_String(CInt(txtTor(1).Text)) & Byte_To_String(IIf(optTipo(0).value, 1, 0)) & txtTor(4).Text)
            End If
        'Iniciar ?
        Else
            Call ClientTCP.Send_Data(Challenge_Main_Cl, Byte_To_String(6))
        End If
        
        Unload Me
End Select

End Sub

Private Sub Form_Load()
Call FormParser.Parse_Form(Me)
End Sub

Private Sub txtTor_KeyPress(Index As Integer, KeyAscii As Integer)

If Index <> 4 Then
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack Or _
            KeyAscii = vbKeyDelete Or (KeyAscii >= 37 And KeyAscii <= 40)) Then
        KeyAscii = 0
    End If
End If

End Sub

Private Function Check_Data() As Boolean

Check_Data = True

If Val(txtTor(0).Text) > 200000 Then
    MensajeAdvertencia Locale_GUI_Frase(249)
    txtTor(2).SetFocus
    Check_Data = False
End If

If Val(txtTor(1).Text) < 2 Then
    MensajeAdvertencia Locale_GUI_Frase(250)
    txtTor(1).SetFocus
    Check_Data = False
End If

If chkRestr.value Then
    If Val(txtTor(2).Text) < 1 Then
        MensajeAdvertencia Locale_GUI_Frase(253)
        txtTor(2).SetFocus
        Check_Data = False
    End If
    
    If Val(txtTor(3).Text) > 50 Then
        MensajeAdvertencia Locale_GUI_Frase(254)
        txtTor(3).SetFocus
        Check_Data = False
    End If
End If

If LenB(txtTor(4).Text) = 0 Then
    MensajeAdvertencia Locale_GUI_Frase(255)
    txtTor(4).SetFocus
    Check_Data = False
End If

End Function
