VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
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
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4350
      Width           =   7530
      ExtentX         =   13282
      ExtentY         =   4895
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
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
   Begin VB.TextBox IPTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Text            =   "localhost"
      Top             =   8640
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.TextBox PortTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   3900
      TabIndex        =   3
      Text            =   "7666"
      Top             =   8640
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.ListBox lst_servers 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   1200
      ItemData        =   "frmConnect.frx":0ECA
      Left            =   6675
      List            =   "frmConnect.frx":0ED1
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
      MaxLength       =   25
      TabIndex        =   0
      Top             =   2400
      Width           =   4215
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
'frmConnect - ImperiumAO - v1.3.0
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

Public Sub CargarLst()

Dim i As Integer

lst_servers.Clear

If ServersRecibidos Then
    For i = 1 To UBound(ServersLst)
        lst_servers.AddItem ServersLst(i).Desc
    Next i
End If

End Sub

Private Sub Command1_Click()
CurServer = 0
IPdelServidor = IPTxt
PuertoDelServidor = PortTxt
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then
    Call EndGame(True)
End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If (Button = vbLeftButton) And (RunWindowed = 1) Then Call Auto_Drag(Me.hwnd)
End Sub

Private Sub Form_Load()

If ServersRecibidos Then
    If CurServer <> 0 Then
        IPTxt = ServersLst(CurServer).IP
        PortTxt = ServersLst(CurServer).Puerto
    Else
        IPTxt = IPdelServidor
        PortTxt = PuertoDelServidor
    End If
    
    Call CargarLst
Else
    lst_servers.Clear
End If

Dim j
For Each j In imgAccion()
j.Tag = "0"
Next

Me.Picture = General_Load_Picture_From_Resource("conectar.bmp")
Call noticias.Navigate("http://noticias.imperiumao.com.ar")

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim i As Integer

For i = 0 To 4
    If imgAccion(i).Tag = "0" Then
        imgAccion(i).Picture = Nothing
        imgAccion(i).Tag = "1"
    End If
Next i

End Sub

Private Sub imgAccion_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

CurServer = 0
IPdelServidor = IPTxt
PuertoDelServidor = PortTxt

Call Sound.Sound_Play(SND_CLICK)
Call imgAccionRestaurar

Select Case Index
    
    Case 0
        
        frmMain.mainWinsock.Close

        If Musica <> CONST_DESHABILITADA Then
            If Musica <> CONST_DESHABILITADA Then
                Sound.NextMusic = MUS_CrearPersonaje
                Sound.Fading = 200
            End If
        End If
                                                         
        frmCrearPersonaje.Show
        Me.Visible = False
        
    Case 1
            
        If frmConnect.MousePointer = 11 Then
            Exit Sub
        End If
        
        'update user info
        CurrentUser.UserName = NameTxt.Text
        
        Dim aux As String
        aux = PwdTxt.Text
        CurrentUser.UserPassword = MD5String(aux)
        
        If CheckUserData Then
            Me.MousePointer = 11
            EstadoLogin = NORMAL
                
            If frmMain.mainWinsock.State <> sckClosed Then _
                frmMain.mainWinsock.Close
            
            frmMain.mainWinsock.Connect CurServerIp, CurServerPort
        End If
        
    Case 2
        Call noticias.Navigate("http://www.imperiumao.com.ar/recuperar2.php")
    Case 3
        Call noticias.Navigate("http://www.imperiumao.com.ar/borrar2.php")
    Case 4
        frmOpciones.Init

End Select

End Sub

Private Sub imgAccion_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Select Case Index
    Case 0 'Crear personaje
        imgAccion(0).Picture = General_Load_Picture_From_Resource("botcreardown.bmp")
        imgAccion(0).Tag = "0"
    Case 1 'Conectar
        imgAccion(1).Picture = General_Load_Picture_From_Resource("botconectardown.bmp")
        imgAccion(1).Tag = "0"
    Case 2 'Recuperar
        imgAccion(2).Picture = General_Load_Picture_From_Resource("botrecuperardown.bmp")
        imgAccion(2).Tag = "0"
    Case 3 'Borrar
        imgAccion(3).Picture = General_Load_Picture_From_Resource("botborrardown.bmp")
        imgAccion(3).Tag = "0"
    Case 4 'Opciones
        imgAccion(4).Picture = General_Load_Picture_From_Resource("botopcionesdown.bmp")
        imgAccion(4).Tag = "0"
End Select

'Call imgAccionRestaurar(Index)

End Sub

Private Sub imgAccion_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Select Case Index
    Case 0 'Crear personaje
        If imgAccion(0).Tag = "1" Then
            imgAccion(0).Picture = General_Load_Picture_From_Resource("botcrearover.bmp")
            imgAccion(0).Tag = "0"
            Call Sound.Sound_Play(SND_OVER)
        End If
        
    Case 1 'Conectar
        If imgAccion(1).Tag = "1" Then
            imgAccion(1).Picture = General_Load_Picture_From_Resource("botconectarover.bmp")
            imgAccion(1).Tag = "0"
            Call Sound.Sound_Play(SND_OVER)
        End If
    Case 2 'Recuperar
        If imgAccion(2).Tag = "1" Then
            imgAccion(2).Picture = General_Load_Picture_From_Resource("botrecuperarover.bmp")
            imgAccion(2).Tag = "0"
            Call Sound.Sound_Play(SND_OVER)
        End If
    Case 3 'Borrar
        If imgAccion(3).Tag = "1" Then
            imgAccion(3).Picture = General_Load_Picture_From_Resource("botborrarover.bmp")
            imgAccion(3).Tag = "0"
            Call Sound.Sound_Play(SND_OVER)
        End If
    Case 4 'Opciones
        If imgAccion(4).Tag = "1" Then
            imgAccion(4).Picture = General_Load_Picture_From_Resource("botopcionesover.bmp")
            imgAccion(4).Tag = "0"
            Call Sound.Sound_Play(SND_OVER)
        End If
End Select

Call imgAccionRestaurar(Index)

End Sub

Private Sub imgAccionRestaurar(Optional ByVal NoIndex As Integer = 1000)

Dim i As Integer

For i = 0 To 4
    If i <> NoIndex Then
        imgAccion(i).Picture = Nothing
        imgAccion(i).Tag = "1"
    End If
Next i

End Sub

Private Sub lst_servers_Click()

If ServersRecibidos Then
    CurServer = lst_servers.ListIndex + 1
    IPTxt = ServersLst(CurServer).IP
    PortTxt = ServersLst(CurServer).Puerto
End If

End Sub

Private Sub lst_servers_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call Form_MouseMove(Button, Shift, x, y)
End Sub

Private Sub NameTxt_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call Form_MouseMove(Button, Shift, x, y)
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
