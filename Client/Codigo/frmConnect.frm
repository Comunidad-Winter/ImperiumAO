VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmConnect 
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
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin SHDocVwCtl.WebBrowser noticias 
      Height          =   2775
      Left            =   2250
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4380
      Width           =   7530
      ExtentX         =   13282
      ExtentY         =   4895
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.ListBox lst_servers 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1200
      ItemData        =   "frmConnect.frx":000C
      Left            =   6675
      List            =   "frmConnect.frx":0013
      TabIndex        =   2
      Top             =   2385
      Width           =   3075
   End
   Begin VB.TextBox PwdTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2250
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3270
      Width           =   2355
   End
   Begin VB.TextBox NameTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2250
      MaxLength       =   20
      TabIndex        =   0
      Top             =   2400
      Width           =   4215
   End
   Begin VB.Image imgAccion 
      Height          =   300
      Index           =   6
      Left            =   11640
      Top             =   60
      Width           =   300
   End
   Begin VB.Image imgAccion 
      Height          =   300
      Index           =   5
      Left            =   11325
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
      Top             =   7575
      Width           =   1755
   End
   Begin VB.Image imgAccion 
      Height          =   615
      Index           =   3
      Left            =   4155
      MousePointer    =   99  'Custom
      Top             =   7575
      Width           =   1755
   End
   Begin VB.Image imgAccion 
      Height          =   615
      Index           =   2
      Left            =   6090
      MousePointer    =   99  'Custom
      Top             =   7575
      Width           =   1755
   End
   Begin VB.Image imgAccion 
      Height          =   615
      Index           =   0
      Left            =   2235
      MousePointer    =   99  'Custom
      Top             =   7575
      Width           =   1755
   End
   Begin VB.Image imgAccion 
      Height          =   615
      Index           =   1
      Left            =   4770
      MousePointer    =   99  'Custom
      Top             =   3000
      Width           =   1755
   End
End
Attribute VB_Name = "frmConnect"
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
'Pablo Ignacio Márquez (morgolock@speedy.com.ar)
'   - First Relase
'Augusto José Rando (barrin@imperiumao.com.ar)
'   - Complete recoding
'*****************************************************************

Option Explicit

Private Sub Form_Activate()

On Error Resume Next

frmConnect.NameTxt.SetFocus

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then Call EndGame(True)

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If (Button = vbLeftButton) And (RunWindowed = 1) Then Call Auto_Drag(Me.hwnd)
End Sub

Private Sub Form_Load()

Dim j
For Each j In imgAccion()
j.Tag = "0"
Next

Me.Caption = Form_Caption
Me.Picture = General_Load_Picture_From_Resource_Ex("conectar")
Call noticias.Navigate(URL_NEWS)
Call ServerList_Load

Call FormParser.Parse_Form(Me)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim i As Integer

For i = 0 To imgAccion.UBound
    If imgAccion(i).Tag = "0" Then
        imgAccion(i).Picture = Nothing
        imgAccion(i).Tag = "1"
    End If
Next i

End Sub

Private Sub imgAccion_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo ErrorHandler

Call Sound.Sound_Play(SND_CLICK)
Call imgAccionRestaurar

Select Case Index
    
    Case 0
                
        ShellExecute Me.hwnd, "open", Chr$(34) & "http://www.imperiumao.com.ar/" & GameLocale & "/crearcuenta.php" & Chr$(34), vbNullString, vbNullString, 1
        
    Case 1
                    
        If frmMain.MainWinsock.State <> sckClosed Then _
            frmMain.MainWinsock.Close
        
        CurrentUser.AccountName = NameTxt.Text
        CurrentUser.UserPassword = MD5String(PwdTxt.Text)
        
        If CheckUserData Then
            EstadoLogin = Normal
            Call FormParser.Parse_Form(Me, E_WAIT)
            frmMain.MainWinsock.Connect CurServerIp, CurServerPort
        End If
        
    Case 2
        Call noticias.Navigate("http://recuperar.imperiumao.com.ar/")
    Case 3
        Call noticias.Navigate("http://borrar.imperiumao.com.ar/")
    Case 4
        frmOpciones.Init
    Case 5
        Me.WindowState = vbMinimized
    Case 6
        Call EndGame(True)
    
End Select

Exit Sub

ErrorHandler:
    Call MsgBox(Locale_GUI_Frase(345) & " (" & Err.Description & " - " & Err.Number & ")", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error al conectar")

End Sub

Private Sub imgAccion_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Select Case Index
    Case 0 'Crear personaje
        imgAccion(0).Picture = General_Load_Picture_From_Resource_Ex("botcreardown")
        imgAccion(0).Tag = "0"
    Case 1 'Conectar
        imgAccion(1).Picture = General_Load_Picture_From_Resource_Ex("botconectardown")
        imgAccion(1).Tag = "0"
    Case 2 'Recuperar
        imgAccion(2).Picture = General_Load_Picture_From_Resource_Ex("botrecuperardown")
        imgAccion(2).Tag = "0"
    Case 3 'Borrar
        imgAccion(3).Picture = General_Load_Picture_From_Resource_Ex("botborrardown")
        imgAccion(3).Tag = "0"
    Case 4 'Opciones
        imgAccion(4).Picture = General_Load_Picture_From_Resource_Ex("botopcionesdown")
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
            imgAccion(0).Picture = General_Load_Picture_From_Resource_Ex("botcrearover")
            imgAccion(0).Tag = "0"
            Call Sound.Sound_Play(SND_OVER)
        End If
        
    Case 1 'Conectar
        If imgAccion(1).Tag = "1" Then
            imgAccion(1).Picture = General_Load_Picture_From_Resource_Ex("botconectarover")
            imgAccion(1).Tag = "0"
            Call Sound.Sound_Play(SND_OVER)
        End If
    Case 2 'Recuperar
        If imgAccion(2).Tag = "1" Then
            imgAccion(2).Picture = General_Load_Picture_From_Resource_Ex("botrecuperarover")
            imgAccion(2).Tag = "0"
            Call Sound.Sound_Play(SND_OVER)
        End If
    Case 3 'Borrar
        If imgAccion(3).Tag = "1" Then
            imgAccion(3).Picture = General_Load_Picture_From_Resource_Ex("botborrarover")
            imgAccion(3).Tag = "0"
            Call Sound.Sound_Play(SND_OVER)
        End If
    Case 4 'Opciones
        If imgAccion(4).Tag = "1" Then
            imgAccion(4).Picture = General_Load_Picture_From_Resource_Ex("botopcionesover")
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
        imgAccion(i).Picture = Nothing
        imgAccion(i).Tag = "1"
    End If
Next i

End Sub

Private Sub lst_servers_Click()

On Error Resume Next

If lst_servers.ListIndex = -1 Then Exit Sub

CurServer = lst_servers.ListIndex + 1

End Sub

Private Sub lst_servers_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call Form_MouseMove(Button, Shift, x, y)
End Sub

Private Sub NameTxt_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call Form_MouseMove(Button, Shift, x, y)
End Sub

Private Sub noticias_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)

If URL <> URL_NEWS And InStr(1, URL, "http://pagead2.googlesyndication.com/pagead/ads") = 0 And _
    InStr(1, URL, "http://smartad.mercadolibre.com.ar/") = 0 And _
    InStr(1, URL, "http://www.game-advertising-online.com/") = 0 Then
    ShellExecute GetDesktopWindow, "open", Chr$(34) & CStr(URL) & Chr$(34), &O0, &O0, SW_Normal
    Cancel = True
End If

End Sub

Private Sub PwdTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call imgAccion_MouseDown(1, 0, 0, 0, 0)
        Call imgAccion_MouseUp(1, 0, 0, 0, 0)
    End If
End Sub

Private Sub NameTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call imgAccion_MouseDown(1, 0, 0, 0, 0)
        Call imgAccion_MouseUp(1, 0, 0, 0, 0)
    End If
End Sub

Private Sub PwdTxt_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call Form_MouseMove(Button, Shift, x, y)
End Sub

Public Sub ServerList_Load()

On Error GoTo ErrorHandler

Dim i As Long

If ServersLstLoaded = False Then Exit Sub

lst_servers.Clear

For i = 1 To UBound(ServersLst)
    If LenB(ServersLst(i).Desc) > 0 Then lst_servers.AddItem ServersLst(i).Desc
Next i

Exit Sub

ErrorHandler:


End Sub

Private Sub WebLink_Click()
ShellExecute Me.hwnd, "open", Chr$(34) & "http://www.imperiumao.com.ar/" & GameLocale & "/" & Chr$(34), vbNullString, vbNullString, 1
End Sub
