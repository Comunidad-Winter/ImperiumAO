VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "ImperiumAO 1.3"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   315
   ClientWidth     =   12000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   0  'User
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin SHDocVwCtl.WebBrowser publi 
      Height          =   645
      Left            =   9000
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   4695
      Visible         =   0   'False
      Width           =   2385
      ExtentX         =   4207
      ExtentY         =   1138
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
   Begin MSWinsockLib.Winsock MainWinsock 
      Left            =   5700
      Top             =   210
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox MiniMap 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1410
      Left            =   10200
      ScaleHeight     =   94
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   29
      Top             =   7380
      Width           =   1455
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   10
      Left            =   6090
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   26
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   9
      Left            =   5505
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   25
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   8
      Left            =   4920
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   24
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   7
      Left            =   4335
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   23
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   6
      Left            =   3750
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   22
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox MainViewPic 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6225
      Left            =   210
      ScaleHeight     =   6225
      ScaleWidth      =   8175
      TabIndex        =   20
      Top             =   2070
      Width           =   8175
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   5
      Left            =   3165
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   12
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   4
      Left            =   2580
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   11
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      ForeColor       =   &H00000000&
      Height          =   2400
      Left            =   9015
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   10
      Top             =   2220
      Width           =   2415
   End
   Begin VB.ListBox hlst 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2565
      Left            =   8850
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2100
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   2
      Left            =   1410
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   3
      Left            =   1995
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   1
      Left            =   825
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   0
      Left            =   240
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   8430
      Width           =   480
   End
   Begin VB.Timer sldTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   16200
      Top             =   16200
   End
   Begin VB.TextBox SendTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   210
      MaxLength       =   500
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1755
      Visible         =   0   'False
      Width           =   7470
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1455
      Left            =   240
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   180
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   2566
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmMain.frx":000C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgHora 
      Height          =   480
      Left            =   6675
      Top             =   8430
      Width           =   1695
   End
   Begin VB.Image imgMiniCerra 
      Enabled         =   0   'False
      Height          =   315
      Left            =   11340
      Top             =   150
      Width           =   510
   End
   Begin VB.Image cmdMenu 
      Height          =   450
      Index           =   4
      Left            =   9270
      MousePointer    =   99  'Custom
      Top             =   4350
      Width           =   1890
   End
   Begin VB.Image cmdMenu 
      Height          =   450
      Index           =   0
      Left            =   9270
      MousePointer    =   99  'Custom
      Top             =   2010
      Width           =   1890
   End
   Begin VB.Image cmdMenu 
      Height          =   450
      Index           =   1
      Left            =   9270
      MousePointer    =   99  'Custom
      Top             =   2595
      Width           =   1890
   End
   Begin VB.Image cmdMenu 
      Height          =   450
      Index           =   2
      Left            =   9270
      MousePointer    =   99  'Custom
      Top             =   3180
      Width           =   1890
   End
   Begin VB.Image cmdMenu 
      Height          =   450
      Index           =   3
      Left            =   9270
      MousePointer    =   99  'Custom
      Top             =   3765
      Width           =   1890
   End
   Begin VB.Image cmdHechizos 
      Height          =   390
      Index           =   0
      Left            =   8775
      MousePointer    =   99  'Custom
      Top             =   4935
      Width           =   1845
   End
   Begin VB.Image cmdMenu 
      Height          =   450
      Index           =   5
      Left            =   9270
      MousePointer    =   99  'Custom
      Top             =   4935
      Width           =   1890
   End
   Begin VB.Image nomodorol 
      Height          =   255
      Left            =   9645
      Picture         =   "frmMain.frx":0089
      ToolTipText     =   "Modo Rol"
      Top             =   7620
      Width           =   300
   End
   Begin VB.Label lblAG 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9330
      TabIndex        =   28
      Top             =   8550
      Width           =   345
   End
   Begin VB.Label lblFU 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9330
      TabIndex        =   27
      Top             =   8340
      Width           =   345
   End
   Begin VB.Image cmdDropGold 
      Height          =   300
      Left            =   10260
      MousePointer    =   99  'Custom
      Top             =   5670
      Width           =   300
   End
   Begin VB.Image imgCentros 
      Height          =   510
      Index           =   2
      Left            =   10740
      Top             =   1230
      Width           =   1065
   End
   Begin VB.Label lblInvInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   9000
      TabIndex        =   21
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   135
      Index           =   1
      Left            =   8820
      TabIndex        =   19
      Top             =   870
      Width           =   1815
   End
   Begin VB.Shape ExpShp 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   8820
      Top             =   900
      Width           =   1815
   End
   Begin VB.Image cmdMinimizar 
      Height          =   225
      Left            =   11280
      Top             =   180
      Width           =   225
   End
   Begin VB.Image cmdCerrar 
      Height          =   225
      Left            =   11580
      Top             =   180
      Width           =   255
   End
   Begin VB.Image cmdMensaje 
      Height          =   255
      Left            =   7800
      Top             =   1725
      Width           =   555
   End
   Begin VB.Label lblNick 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NickDelPersonaje"
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
      Left            =   8610
      TabIndex        =   18
      Top             =   180
      Width           =   2625
   End
   Begin VB.Image imgCentros 
      Height          =   510
      Index           =   1
      Left            =   9660
      Top             =   1230
      Width           =   1065
   End
   Begin VB.Image imgCentros 
      Height          =   510
      Index           =   0
      Left            =   8580
      Top             =   1230
      Width           =   1065
   End
   Begin VB.Image cmdHechizos 
      Height          =   420
      Index           =   3
      Left            =   11460
      Top             =   3405
      Width           =   300
   End
   Begin VB.Image cmdHechizos 
      Height          =   420
      Index           =   2
      Left            =   11475
      Top             =   2910
      Width           =   300
   End
   Begin VB.Image cmdHechizos 
      Height          =   390
      Index           =   1
      Left            =   10650
      MousePointer    =   99  'Custom
      Top             =   4935
      Width           =   945
   End
   Begin VB.Label lblSED 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   10320
      TabIndex        =   17
      Top             =   6660
      Width           =   1365
   End
   Begin VB.Label lblHAM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   10320
      TabIndex        =   16
      Top             =   6255
      Width           =   1365
   End
   Begin VB.Label lblST 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   8745
      TabIndex        =   15
      Top             =   6660
      Width           =   1365
   End
   Begin VB.Label lblMP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   8745
      TabIndex        =   14
      Top             =   6255
      Width           =   1365
   End
   Begin VB.Label lblHP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   8745
      TabIndex        =   13
      Top             =   5850
      Width           =   1350
   End
   Begin VB.Label LvlLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10950
      TabIndex        =   9
      Top             =   870
      Width           =   435
   End
   Begin VB.Image InvEqu 
      Height          =   4275
      Left            =   8580
      Top             =   1230
      Width           =   3240
   End
   Begin VB.Shape Hpshp 
      BackColor       =   &H00000080&
      BorderColor     =   &H8000000D&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   8745
      Top             =   5880
      Width           =   1365
   End
   Begin VB.Shape STAShp 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      Height          =   135
      Left            =   8745
      Top             =   6690
      Width           =   1365
   End
   Begin VB.Shape MANShp 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   8745
      Top             =   6285
      Width           =   1365
   End
   Begin VB.Shape COMIDAsp 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   135
      Left            =   10320
      Top             =   6285
      Width           =   1365
   End
   Begin VB.Shape AGUAsp 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00800000&
      Height          =   135
      Left            =   10320
      Top             =   6690
      Width           =   1365
   End
   Begin VB.Label GldLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10620
      TabIndex        =   8
      Top             =   5745
      Width           =   1110
   End
   Begin VB.Image nomodoseguro 
      Height          =   255
      Left            =   9240
      Picture         =   "frmMain.frx":04C7
      ToolTipText     =   "Seguro"
      Top             =   7620
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image nomodocombate 
      Height          =   255
      Left            =   8820
      Picture         =   "frmMain.frx":0905
      ToolTipText     =   "Modo Combate"
      Top             =   7620
      Width           =   300
   End
   Begin VB.Image modorol 
      Height          =   255
      Left            =   9645
      Picture         =   "frmMain.frx":0D43
      ToolTipText     =   "Modo Rol"
      Top             =   7620
      Width           =   300
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Mapa desconocido"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   8640
      TabIndex        =   2
      Top             =   7020
      Width           =   3105
   End
   Begin VB.Image modoseguro 
      Height          =   255
      Left            =   9240
      Picture         =   "frmMain.frx":12D9
      ToolTipText     =   "Seguro"
      Top             =   7620
      Width           =   300
   End
   Begin VB.Image modocombate 
      Height          =   255
      Left            =   8820
      Picture         =   "frmMain.frx":1717
      ToolTipText     =   "Modo Combate"
      Top             =   7620
      Visible         =   0   'False
      Width           =   300
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmMain - ImperiumAO - v1.4.5 R5
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

Private Const EM_GETSEL As Long = &HB0

Public tX As Integer
Public tY As Integer
Public MouseX As Long
Public MouseY As Long

Public PedimosEst As Boolean

'Barrin
Dim UltimoIndex As Integer

Public UltPos As Integer
Public UltPosInterface As Integer
Public UltPosSolapas As Integer
Public LoadedSkin As String

Public CentroActual As Byte

Private m_Jpeg As clsJpeg
Private m_FileName As String

'DX8 Events
Implements DirectXEvent8

Public WithEvents Engine As clsTileEngineX
Attribute Engine.VB_VarHelpID = -1

Private Sub cmdMensaje_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdMensaje.Picture = General_Load_Skin_Picture_From_Resource_Ex("modotextodown")
End Sub

Private Sub cmdMensaje_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdMensaje.Picture = General_Load_Skin_Picture_From_Resource_Ex("modotextoover")
End Sub

Private Sub cmdMensaje_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Call cmdMensaje_MouseMove(Button, Shift, x, y)
frmMensaje.PopupMenuMensaje
cmdMensaje.Picture = General_Load_Skin_Picture_From_Resource_Ex("modotextoover")
End Sub

Private Sub cmdMinimizar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
imgMiniCerra.Picture = General_Load_Skin_Picture_From_Resource_Ex("minimizardown")
End Sub

Private Sub cmdCerrar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Call ClientTCP.Send_Data_Command(cmdSalir, Integer_To_String(CurrentUser.UserMinMAN) & Integer_To_String(CurrentUser.UserMinSTA) & Byte_To_String(1))
CurrentUser.EndingGame = True
Call EndGame(True)
End Sub

Private Sub cmdMinimizar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
imgMiniCerra.Picture = General_Load_Skin_Picture_From_Resource_Ex("minimizarover")
End Sub

Private Sub cmdCerrar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
imgMiniCerra.Picture = General_Load_Skin_Picture_From_Resource_Ex("cerrarover")
End Sub

Private Sub cmdMinimizar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.WindowState = vbMinimized
imgMiniCerra.Picture = Nothing
End Sub

Private Sub cmdCerrar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
imgMiniCerra.Picture = General_Load_Skin_Picture_From_Resource_Ex("cerrardown")
End Sub

Private Sub Engine_ScrollComplete()

Select Case MoveUserChar(frmMain.Engine.Char_Current_Next_Direction_Get)
    Case 1
        Call frmMain.Engine.Char_Current_Next_Direction_Set(0)
    Case 0
        Call frmMain.Engine.Char_Current_Next_Direction_Set(0)
        Call frmMain.Engine.Char_Current_Blocked_Set(True)
    Case -1
        'Unfocused! Waiting for focus...
End Select

End Sub

Private Sub Form_Activate()
    If SendTxt.Visible Then SendTxt.SetFocus
End Sub

Private Sub cmdDropGold_Click()

ItemElegido = FLAGORO

If Not CurrentUser.Comerciando Then
    If CurrentUser.UserGLD > 0 Then
        frmCantidad.Show vbModeless, frmMain
    End If
Else
    Call PrintToConsole(Locale_GUI_Frase(236), 255, 0, 32, False, False, False)
End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If (Button = vbLeftButton) And (RunWindowed = 1) Then Call Auto_Drag(Me.hwnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
StopURLDetect
End Sub

Private Sub hlst_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If CurrentUser.UsingSkill = Magia Then
    Call FormParser.Parse_Form(frmMain)
    CurrentUser.UsingSkill = 0
End If

End Sub

Private Sub hlst_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call Form_MouseMove(Button, Shift, x, y)
End Sub

Private Sub imgHora_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
imgHora.ToolTipText = Locale_GUI_Frase(302) & " " & Meteo_Engine.Get_Time_String
End Sub

Private Sub InvEqu_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call Form_MouseMove(Button, Shift, x, y)
End Sub

Private Sub Label2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Dim map_x As Integer
Dim map_y As Integer

Call frmMain.Engine.Char_Pos_Get(CurrentUser.CurrentChar, map_x, map_y)

If UltPos <> Index Then
    
    If UltPos >= 0 Then
        If Index = 1 Then
            Label2(Index).Caption = CurrentUser.UserPercExp & "%"
        Else
            If VerLugar = 1 Then
                Label2(Index).Caption = frmMain.Engine.Map_Name_Get
            Else
                Label2(Index).Caption = Locale_GUI_Frase(170) & ": " & CurrentUser.MapNum & ", " & map_x & ", " & map_y
            End If
        End If
    End If
    
    If Index = 1 Then
        Label2(Index).Caption = CurrentUser.UserExp & "/" & CurrentUser.UserPasarNivel
    Else
        If VerLugar = 1 Then
            Label2(Index).Caption = Locale_GUI_Frase(170) & ": " & CurrentUser.MapNum & ", " & map_x & ", " & map_y
        Else
            Label2(Index).Caption = frmMain.Engine.Map_Name_Get
        End If
    End If
    
    If CurrentUser.UserPasarNivel = 0 Then
        Label2(1).Caption = Locale_GUI_Frase(173)
    End If
    
    UltPos = Index
End If

End Sub

Private Sub lbMensaje_Click()
PopupMenu frmMensaje.mnuMensaje
End Sub

Private Sub MiniMap_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If GetKeyState(vbKeyShift) < 0 Then
    Call ClientTCP.Send_Data_Command_GM(cmdTeleploc, Integer_To_String(x) & Integer_To_String(y))
    Exit Sub
End If

End Sub

Private Sub MiniMap_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If GetKeyState(vbKeyShift) < 0 And Engine.Input_Mouse_Button_Left_Get Then
    Call ClientTCP.Send_Data_Command_GM(cmdTeleploc, Integer_To_String(x) & Integer_To_String(y))
    Exit Sub
End If

End Sub

Private Sub modocombate_Click()
    Call ClientTCP.Send_Data(Combat_Mode)
    CurrentUser.Combate = Not CurrentUser.Combate
    modocombate.Visible = Not modocombate.Visible
    nomodocombate.Visible = Not nomodocombate.Visible
End Sub

Private Sub modoseguro_Click()
    Call ClientTCP.Send_Data(Safe_Mode)
    CurrentUser.Seguro = Not CurrentUser.Seguro
    modoseguro.Visible = Not modoseguro.Visible
    nomodoseguro.Visible = Not nomodoseguro.Visible
End Sub

Private Sub modorol_Click()
    Call ClientTCP.Send_Data(Role_Mode)
    CurrentUser.Rol = Not CurrentUser.Rol
    modorol.Visible = Not modorol.Visible
    nomodorol.Visible = Not nomodorol.Visible
End Sub

Private Sub nomodocombate_Click()
    Call ClientTCP.Send_Data(Combat_Mode)
    CurrentUser.Combate = Not CurrentUser.Combate
    modocombate.Visible = Not modocombate.Visible
    nomodocombate.Visible = Not nomodocombate.Visible
End Sub

Private Sub nomodorol_Click()
    Call ClientTCP.Send_Data(Role_Mode)
    CurrentUser.Rol = Not CurrentUser.Rol
    modorol.Visible = Not modorol.Visible
    nomodorol.Visible = Not nomodorol.Visible
End Sub

Private Sub nomodoseguro_Click()
    Call ClientTCP.Send_Data(Safe_Mode)
    CurrentUser.Seguro = Not CurrentUser.Seguro
    modoseguro.Visible = Not modoseguro.Visible
    nomodoseguro.Visible = Not nomodoseguro.Visible
End Sub

Private Sub picInv_Paint()
frmMain.Engine.Engine_Inventory_Render_Set
End Sub

Private Sub picMacro_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

BotonElegido = Index + 1

If MacroKeys(BotonElegido).TipoAccion = 0 Or Button = vbRightButton Then
    frmBindKey.Show vbModeless, frmMain
Else
    Call Bind_Accion(Index + 1)
End If

End Sub

Private Sub picMacro_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If UltimoIndex <> Index Then
    'If UltimoIndex >= 0 Then DibujarMenuMacros UltimoIndex + 1
    'DibujarMenuMacros Index + 1, 1
    UltimoIndex = Index
End If

End Sub

Private Function LoWord(ByVal DWord As Long) As Long
  If DWord And &H8000& Then
    LoWord = DWord Or &HFFFF0000
  Else
    LoWord = DWord And &HFFFF&
  End If
End Function

Private Function HiWord(ByVal DWord As Long) As Long
  HiWord = (DWord And &HFFFF0000) \ &H10000
End Function

Private Sub RecTxt_SelChange()

On Error Resume Next

Clipboard.SetText (RecTxt.SelText)

End Sub

Private Sub TirarItem()
    
    If (ItemElegido > 0 And ItemElegido <= MAX_INVENTORY_SLOTS) Or (ItemElegido = FLAGORO) Then
        frmCantidad.Show vbModeless, frmMain
    End If

End Sub

Private Sub AgarrarItem()
Call ClientTCP.Send_Data(Get_Item)
End Sub

Private Sub UsarItem()
    If (ItemElegido > 0) And (ItemElegido <= MAX_INVENTORY_SLOTS) Then
        If Not ClientTCP.MeditarCheck() Then Call ClientTCP.Send_Data(Use_Item, Integer_To_String(ItemElegido))
    End If
End Sub

Private Sub EquiparItem()
    If (ItemElegido > 0) And (ItemElegido <= MAX_INVENTORY_SLOTS) Then
        If Not ClientTCP.MeditarCheck() And Not ClientTCP.DeadCheck() Then Call ClientTCP.Send_Data(Equip_Item, Integer_To_String(ItemElegido))
    End If
End Sub

Private Sub Form_Load()

Call StartURLDetect(RecTxt.hwnd, Me.hwnd)

Me.Picture = General_Load_Skin_Picture_From_Resource_Ex("todo")
Me.Caption = Form_Caption
Call Make_Transparent_Richtext(RecTxt.hwnd)
Call CambiaCentro(CentroInventario)

UltPos = -1
UltimoIndex = -1
UltPosInterface = -1
UltPosSolapas = -1

Call FormParser.Parse_Form(Me)

LoadedSkin = NombreSkin

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    MouseX = x
    MouseY = y
    
    Dim map_x As Integer, map_y As Integer
    
    If UltimoIndex >= 0 Then
        'DibujarMenuMacros UltimoIndex + 1
        UltimoIndex = -1
    End If
    
    If UltPos >= 0 Then
        Call frmMain.Engine.Char_Pos_Get(CurrentUser.CurrentChar, map_x, map_y)
        
        If UltPos = 1 Then
            Label2(UltPos).Caption = CurrentUser.UserPercExp & "%"
        Else
            If VerLugar = 1 Then
                Label2(UltPos).Caption = frmMain.Engine.Map_Name_Get
            Else
                Label2(0).Caption = Locale_GUI_Frase(170) & ": " & CurrentUser.MapNum & ", " & map_x & ", " & map_y
            End If
        End If
        
        If CurrentUser.UserPasarNivel = 0 Then
            Label2(1).Caption = Locale_GUI_Frase(173)
        End If
        
        UltPos = -1
        
    End If
    
    Call RestaurarCentroActual
    
End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
       KeyCode = 0
End Sub
Private Sub hlst_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
End Sub
Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
        KeyCode = 0
End Sub

Private Sub MostrarCentroInventario()
    InvEqu.Picture = General_Load_Skin_Picture_From_Resource_Ex("centroinventario")
    picInv.Visible = True
    lblInvInfo.Visible = True
    lblInvInfo = vbNullString
    CurrentUser.LastItem = 0
    Call Banner_Logic
End Sub

Private Sub OcultarCentroInventario()
    picInv.Visible = False
    lblInvInfo.Visible = False
    CurrentUser.LastItem = 0
    Call Banner_Logic
End Sub

Private Sub MostrarCentroHechizos()
    InvEqu.Picture = General_Load_Skin_Picture_From_Resource_Ex("centrohechizos")
    cmdHechizos(0).Visible = True
    cmdHechizos(1).Visible = True
    cmdHechizos(2).Visible = True
    cmdHechizos(3).Visible = True
    hlst.Visible = True
End Sub

Private Sub OcultarCentroHechizos()
    hlst.Visible = False
    cmdHechizos(0).Visible = False
    cmdHechizos(1).Visible = False
    cmdHechizos(2).Visible = False
    cmdHechizos(3).Visible = False
End Sub

Private Sub MostrarCentroMenu()
    cmdMenu(0).Visible = True
    cmdMenu(1).Visible = True
    cmdMenu(2).Visible = True
    cmdMenu(3).Visible = True
    cmdMenu(4).Visible = True
    cmdMenu(5).Visible = True
    InvEqu.Picture = General_Load_Skin_Picture_From_Resource_Ex("centromenu")
End Sub

Private Sub OcultarCentroMenu()
    cmdMenu(0).Visible = False
    cmdMenu(1).Visible = False
    cmdMenu(2).Visible = False
    cmdMenu(3).Visible = False
    cmdMenu(4).Visible = False
    cmdMenu(5).Visible = False
End Sub

Public Sub CambiaCentro(NuevoCentro As Byte)

CentroActual = NuevoCentro

If NuevoCentro = CentroMenu Then
    Call MostrarCentroMenu
    Call OcultarCentroHechizos
    Call OcultarCentroInventario
ElseIf NuevoCentro = CentroHechizos Then
    Call MostrarCentroHechizos
    Call OcultarCentroMenu
    Call OcultarCentroInventario
Else
    Call MostrarCentroInventario
    Call OcultarCentroHechizos
    Call OcultarCentroMenu
End If

End Sub

Private Sub picInv_DblClick()
    If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub
    If IntervaloPermiteUsar Then UsarItem
End Sub

Private Sub picInv_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim Mx As Integer
    Dim My As Integer
    Dim aux As Integer
    Dim tStr As String
    
    Mx = x \ 32 + 1
    My = y \ 32
    aux = (Mx + My * 5)
    
    If aux > 0 And aux <= MAX_INVENTORY_SLOTS And aux <> CurrentUser.LastItem Then
        CurrentUser.LastItem = aux
        
        Select Case UserInventory(aux).ObjType
            Case 2
                tStr = UserInventory(aux).name & vbCr & Locale_GUI_Frase(175) & ": " & UserInventory(aux).MinHIT & "/" & UserInventory(aux).MaxHIT
            Case 3
                tStr = UserInventory(aux).name & vbCr & Locale_GUI_Frase(176) & ": " & UserInventory(aux).Def
            Case Else
                tStr = UserInventory(aux).name
        End Select
        
        If Len(UserInventory(aux).ExtraStr) > 0 Then
            tStr = tStr & vbCr & UserInventory(aux).ExtraStr
        End If
        
        If publi.Visible Then
            picInv.ToolTipText = Replace$(tStr, vbCr, " - ")
            lblInvInfo.Caption = vbNullString
        Else
            lblInvInfo.FontSize = IIf(Len(tStr) > 62, 7, 8)
            lblInvInfo.Caption = tStr
            picInv.ToolTipText = vbNullString
        End If
        
    End If
    
End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ItemClick(CInt(x), CInt(y))
End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If picInv.Visible Then
        picInv.SetFocus
    ElseIf hlst.Visible Then
        hlst.SetFocus
    End If
End Sub

Private Sub SendTxt_Change()
stxtBuffer = SendTxt.Text
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Or (KeyAscii = 126) Or (KeyAscii = 176) Then _
        KeyAscii = 0
End Sub

Private Sub CompletarEnvioMensajes()

Select Case CurrentUser.SendingType
    Case 1
        SendTxt.Text = vbNullString
    Case 2
        SendTxt.Text = "-"
    Case 3
        SendTxt.Text = ("\" & CurrentUser.sndPrivateTo & " ")
    Case 4
        SendTxt.Text = "/CMSG "
    Case 5
        SendTxt.Text = "/GMSG "
    Case 6
        SendTxt.Text = "/GRMG "
    Case 7
        SendTxt.Text = ";"
    Case 8
        SendTxt.Text = "/FMMG "
End Select

stxtBuffer = SendTxt.Text
SendTxt.SelStart = Len(SendTxt.Text)

End Sub

Private Sub Enviar_SendTxt()
    
    Dim str1 As String
    Dim str2 As String
    
    If Len(stxtBuffer) > 255 Then stxtBuffer = mid$(stxtBuffer, 1, 255)
    
    'Send text
    If left$(stxtBuffer, 1) = "/" Then
        Call ClientTCP.Parse_Command_Str(stxtBuffer)

    'Shout
    ElseIf left$(stxtBuffer, 1) = "-" Then
        If Right$(stxtBuffer, Len(stxtBuffer) - 1) <> vbNullString Then Call ClientTCP.Send_Data(Shout_Chat, Right$(stxtBuffer, Len(stxtBuffer) - 1))
        CurrentUser.SendingType = 2
        
    'Global
    ElseIf left$(stxtBuffer, 1) = ";" Then
        If LenB(Right$(stxtBuffer, Len(stxtBuffer) - 1)) > 0 And InStr(stxtBuffer, ">") = 0 Then Call ClientTCP.Send_Data(Global_Chat, Right$(stxtBuffer, Len(stxtBuffer) - 1))
        CurrentUser.SendingType = 7

    'Privado
    ElseIf left$(stxtBuffer, 1) = "\" Then
        str1 = Right$(stxtBuffer, Len(stxtBuffer) - 1)
        str2 = General_Field_Read(1, str1, " ")
        If LenB(str1) > 0 And InStr(str1, ">") = 0 Then Call ClientTCP.Send_Data(Private_Chat, str1)
        CurrentUser.sndPrivateTo = str2
        CurrentUser.SendingType = 3
                
    'Say
    Else
        If LenB(stxtBuffer) > 0 Then Call ClientTCP.Send_Data(Normal_Chat, stxtBuffer)
        CurrentUser.SendingType = 1
    End If

    stxtBuffer = vbNullString
    SendTxt.Text = vbNullString
    SendTxt.Visible = False
    
End Sub

'[Barrin]
Private Sub Bind_Accion(ByVal FNUM As Integer)

If MacroKeys(FNUM).TipoAccion = 0 Then Exit Sub

Select Case MacroKeys(FNUM).TipoAccion

Case 1 'Envia comando
    Call ClientTCP.Parse_Command_Str("/" & MacroKeys(FNUM).SendString)
Case 2 'Lanza hechizo
    If hlst.List(MacroKeys(FNUM).hlist - 1) <> Locale_GUI_Frase(269) And CurrentUser.Descansando = False Then
        If ClientTCP.DeadCheck Then Exit Sub
        Call ClientTCP.Send_Data(Cast_Spell, Byte_To_String(MacroKeys(FNUM).hlist) & Integer_To_String(CurrentUser.UserMinSTA))
    End If
Case 3 'Equipa
    If ClientTCP.DeadCheck Then Exit Sub
    Call EquiparItemMacro(MacroKeys(FNUM).invslot)
Case 4 'Usa
    If IntervaloPermiteUsar Then Call UsarItemMacro(MacroKeys(FNUM).invslot)
End Select

End Sub

Private Sub EquiparItemMacro(SelectedItemSlot As Integer)
    If (SelectedItemSlot > 0) And (SelectedItemSlot <= MAX_INVENTORY_SLOTS) Then
        If Not ClientTCP.MeditarCheck() And Not ClientTCP.DeadCheck() Then Call ClientTCP.Send_Data(Equip_Item, Integer_To_String(SelectedItemSlot))
    End If
End Sub

Private Sub UsarItemMacro(SelectedItemSlot As Integer)
    If (SelectedItemSlot > 0) And (SelectedItemSlot <= MAX_INVENTORY_SLOTS) Then
        If Not ClientTCP.MeditarCheck() Then Call ClientTCP.Send_Data(Use_Item, Integer_To_String(SelectedItemSlot))
    End If
End Sub
'[/Barrin]

Private Sub mainWinsock_Connect()

Call ClientTCP.Send_Data(Auth_Start)

End Sub

Private Sub mainWinsock_Close()

On Error Resume Next

Call Winsock_Error_Close_Event

End Sub

Private Sub mainWinsock_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

On Error Resume Next

If MainWinsock.State Then MainWinsock.Close
Call Winsock_Error_Close_Event(Description, Number)

End Sub

Private Sub mainWinsock_DataArrival(ByVal BytesTotal As Long)

On Error Resume Next

Dim Completed_Comm As Boolean
Dim Rec_Str As String
Dim Comm_Size As Integer

MainWinsock.GetData Rec_Str

If CurrentUser.RDBuffer <> vbNullString Then
    CurrentUser.RDBuffer = CurrentUser.RDBuffer & Rec_Str
Else
    CurrentUser.RDBuffer = Rec_Str
End If

Comm_Size = String_To_Integer(CurrentUser.RDBuffer, 1)
Completed_Comm = Comm_Size <= (Len(CurrentUser.RDBuffer) - 2)

Do While Completed_Comm
    Call ClientTCP.Handle_Data(mid$(CurrentUser.RDBuffer, 3, Comm_Size))
    CurrentUser.RDBuffer = mid$(CurrentUser.RDBuffer, Comm_Size + 3)
    If Len(CurrentUser.RDBuffer) > 1 Then
        Comm_Size = String_To_Integer(CurrentUser.RDBuffer, 1)
        Completed_Comm = Comm_Size <= (Len(CurrentUser.RDBuffer) - 2)
    Else
        Completed_Comm = False
    End If
Loop

End Sub

'Private Function HechizoInvalido(ByVal HechizoName As String) As Boolean
'
'HechizoName = UCase$(HechizoName)
'
'If HechizoName = "REMOVER PARALISIS" Or HechizoName = "DESENCANTAR" Or HechizoName = "SANAR" Then
'    HechizoInvalido = True
'    Exit Function
'End If
'
'End Function

'###########################################################
'                        GUI
'###########################################################

Private Sub cmdHechizos_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If CentroActual <> CentroHechizos Then Exit Sub

Select Case Index
    Case 0 'Lanzar
        cmdHechizos(0).Picture = General_Load_Skin_Picture_From_Resource_Ex("[hechizos]lanzar-down")
    Case 1 'Info
        cmdHechizos(1).Picture = General_Load_Skin_Picture_From_Resource_Ex("[hechizos]info-down")
    Case 2 'Subir
        cmdHechizos(2).Picture = General_Load_Skin_Picture_From_Resource_Ex("[hechizos]flechaarriba-down")
    Case 3 'Bajar
        cmdHechizos(3).Picture = General_Load_Skin_Picture_From_Resource_Ex("[hechizos]flechaabajo-down")
End Select

End Sub

Private Sub cmdHechizos_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If hlst.Visible = False Then Exit Sub
If UltPosInterface = Index Then Exit Sub

If UltPosInterface <> -1 Then Call RestaurarCentroActual
UltPosInterface = Index

Select Case Index
    Case 0 'lanzar
        cmdHechizos(0).Picture = General_Load_Skin_Picture_From_Resource_Ex("[hechizos]lanzar-over")
    Case 1 'info
        cmdHechizos(1).Picture = General_Load_Skin_Picture_From_Resource_Ex("[hechizos]info-over")
    Case 2 'Subir
        cmdHechizos(2).Picture = General_Load_Skin_Picture_From_Resource_Ex("[hechizos]flechaarriba-over")
    Case 3 'Bajar
        cmdHechizos(3).Picture = General_Load_Skin_Picture_From_Resource_Ex("[hechizos]flechaabajo-over")
End Select

End Sub

Private Sub cmdHechizos_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If CentroActual <> CentroHechizos Then Exit Sub
Call Form_MouseMove(Button, Shift, x, y)

If hlst.ListIndex = -1 Then Exit Sub

Call Sound.Sound_Play(SND_CLICK)

Select Case Index
    Case 0 'lanzar
        If hlst.List(hlst.ListIndex) <> Locale_GUI_Frase(269) And CurrentUser.Descansando = False Then
            If ClientTCP.DeadCheck Then Exit Sub
            Call ClientTCP.Send_Data(Cast_Spell, Byte_To_String(hlst.ListIndex + 1) & Integer_To_String(CurrentUser.UserMinSTA))
        End If
    Case 1 'info
        Call ClientTCP.Send_Data(Spell_Info_Request, hlst.ListIndex + 1)
    Case 2 'subir
        If hlst.ListIndex = 0 Then Exit Sub
        Call ClientTCP.Send_Data(Spell_Move, 1 & "," & hlst.ListIndex + 1)
        hlst.ListIndex = hlst.ListIndex - 1
    Case 3 'bajar
        If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub
        Call ClientTCP.Send_Data(Spell_Move, 2 & "," & hlst.ListIndex + 1)
        hlst.ListIndex = hlst.ListIndex + 1
End Select

End Sub

Private Sub CentroHechizosRestaurar(Index As Integer)

cmdHechizos(Index).Picture = Nothing

End Sub

Private Sub SolapasRestaurar(Index As Integer)

imgCentros(Index).Picture = Nothing
imgMiniCerra.Picture = Nothing
cmdMensaje.Picture = Nothing

End Sub

Private Sub imgCentros_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Select Case Index
    Case 0 'Inv
        'imgCentros(0).Picture = General_Load_Skin_Picture_From_Resource_Ex("[Solapas]Inventario-Down")
    Case 1 'Hechizos
        'imgCentros(1).Picture = General_Load_Skin_Picture_From_Resource_Ex("[Solapas]Hechizos-Down")
    Case 2 'Menu
        'imgCentros(2).Picture = General_Load_Skin_Picture_From_Resource_Ex("[Solapas]Menu-Down")
End Select

End Sub

Private Sub imgCentros_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If UltPosSolapas = Index Then Exit Sub

If UltPosSolapas <> -1 Then Call RestaurarCentroActual
UltPosSolapas = Index

Select Case Index
    Case 0 'Inv
        imgCentros(0).Picture = General_Load_Skin_Picture_From_Resource_Ex("[solapas]inventario-over")
    Case 1 'Hechizos
        imgCentros(1).Picture = General_Load_Skin_Picture_From_Resource_Ex("[solapas]hechizos-over")
    Case 2 'Menu
        imgCentros(2).Picture = General_Load_Skin_Picture_From_Resource_Ex("[solapas]menu-over")
End Select

End Sub

Private Sub cmdMenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Select Case Index
    Case 0 'Grupo
        cmdMenu(0).Picture = General_Load_Skin_Picture_From_Resource_Ex("[menu]grupo-down")
    Case 1 'Estadisticas
        cmdMenu(1).Picture = General_Load_Skin_Picture_From_Resource_Ex("[menu]estadisticas-down")
    Case 2 'Guild
        cmdMenu(2).Picture = General_Load_Skin_Picture_From_Resource_Ex("[menu]clanes-down")
    Case 3 'Quest
        cmdMenu(3).Picture = General_Load_Skin_Picture_From_Resource_Ex("[menu]quests-down")
    Case 4 'Torneos
        cmdMenu(4).Picture = General_Load_Skin_Picture_From_Resource_Ex("[menu]torneos-down")
    Case 5 'Opciones
        cmdMenu(5).Picture = General_Load_Skin_Picture_From_Resource_Ex("[menu]opciones-down")
End Select

End Sub

Private Sub cmdMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If UltPosInterface = Index Then Exit Sub

If UltPosInterface <> -1 Then Call RestaurarCentroActual
UltPosInterface = Index

Select Case Index

    Case 0 'Grupo
        cmdMenu(0).Picture = General_Load_Skin_Picture_From_Resource_Ex("[menu]grupo-over")
    Case 1 'Estadisticas
        cmdMenu(1).Picture = General_Load_Skin_Picture_From_Resource_Ex("[menu]estadisticas-over")
    Case 2 'Guild
        cmdMenu(2).Picture = General_Load_Skin_Picture_From_Resource_Ex("[menu]clanes-over")
    Case 3 'Quest
        cmdMenu(3).Picture = General_Load_Skin_Picture_From_Resource_Ex("[menu]quests-over")
    Case 4 'Torneos
        cmdMenu(4).Picture = General_Load_Skin_Picture_From_Resource_Ex("[menu]torneos-over")
    Case 5 'Opciones
        cmdMenu(5).Picture = General_Load_Skin_Picture_From_Resource_Ex("[menu]opciones-over")
End Select

End Sub

Private Sub cmdMenu_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If CentroActual <> CentroMenu Then Exit Sub
Call Form_MouseMove(Button, Shift, x, y)

Call Sound.Sound_Play(SND_CLICK)

Select Case Index
    Case 0 'Grupo
        Call ClientTCP.Send_Data(Group_Member_List)
    Case 1 'Estadisticas
        Call ClientTCP.Reset_Skill_Data
        Call ClientTCP.Send_Data(Stats_Att_Request)
        Call ClientTCP.Send_Data(Stats_Skills_Request)
        Call ClientTCP.Send_Data(Stats_Familiar_Request)
        Call ClientTCP.Send_Data(Stats_General_Request)
        PedimosEst = True
    Case 2 'Guild
        If Not (frmGuildLeader.Visible Or frmGuildAdm.Visible) Then _
            Call ClientTCP.Send_Data(Guild_Info_Request)
    Case 3 'Quest
        Call ClientTCP.Send_Data(Quest_Data_Cl)
    Case 4 'Torneos
        Call ClientTCP.Send_Data(Challenge_Main_Cl, Byte_To_String(0))
    Case 5 'Opciones
        Call frmOpciones.Init
End Select

End Sub

Private Sub imgCentros_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Call Form_MouseMove(Button, Shift, x, y)
Call Sound.Sound_Play(SND_CLICK)

Select Case Index
    Case 0
        Call CambiaCentro(CentroInventario)
    Case 1
        Call CambiaCentro(CentroHechizos)
    Case 2
        Call CambiaCentro(CentroMenu)
End Select

End Sub

Private Sub CentroMenuRestaurar(Index As Integer)

cmdMenu(Index).Picture = Nothing

End Sub

Private Sub RestaurarCentroActual()

Select Case CentroActual
    Case CentroHechizos
        If UltPosInterface <> -1 Then Call CentroHechizosRestaurar(UltPosInterface)
    Case CentroInventario
    Case CentroMenu
        If UltPosInterface <> -1 Then Call CentroMenuRestaurar(UltPosInterface)
End Select

If UltPosSolapas <> -1 Then Call SolapasRestaurar(UltPosSolapas)

UltPosInterface = -1
UltPosSolapas = -1

imgMiniCerra.Picture = Nothing
cmdMensaje.Picture = Nothing
lblInvInfo.Caption = vbNullString
CurrentUser.LastItem = 0

End Sub

Public Sub Client_Screenshot(ByVal hDC As Long, ByVal Width As Long, ByVal Height As Long)

On Error GoTo ErrorHandler

Dim i As Long
Dim Index As Long
i = 1

Set m_Jpeg = New clsJpeg

'80 Quality
m_Jpeg.Quality = 100

'Sample the cImage by hDC
m_Jpeg.SampleHDC hDC, Width, Height

m_FileName = App.Path & "\Fotos\ImpAO_Foto"

If dir(App.Path & "\Fotos", vbDirectory) = vbNullString Then
    MkDir (App.Path & "\Fotos")
End If

Do While dir(m_FileName & Trim(str(i)) & ".jpg") <> vbNullString
    i = i + 1
    DoEvents
Loop

Index = i

m_Jpeg.Comment = "Character: " & CurrentUser.UserName & " - " & format(Date, "dd/mm/yyyy") & " - " & format(Time, "hh:mm AM/PM")

'Save the JPG file
m_Jpeg.SaveFile m_FileName & Trim(str(Index)) & ".jpg"

Call PrintToConsole(Locale_GUI_Frase(360) & " " & m_FileName & Trim(str(Index)) & ".jpg", 65, 190, 156, False, True, False)

Set m_Jpeg = Nothing

Exit Sub

ErrorHandler:
    Call PrintToConsole(Locale_GUI_Frase(361), 65, 190, 156, False, True, False)

End Sub

Private Sub DirectXEvent8_DXCallback(ByVal eventid As Long)

On Error GoTo Error_Handler

Dim map_x As Integer
Dim map_y As Integer

If CurrentUser.Pausa Or Not _
CurrentUser.Logged Or CurrentUser.Comerciando Then Exit Sub

If eventid = frmMain.Engine.Input_Keyboard_DXEvent Then
    
    frmMain.Engine.Input_Keyboard_Start
    
    If GetActiveWindow <> hwnd Then
        If Not CurrentUser.AutoNavigation Then Check_Main_Keys
        frmMain.Engine.Input_Keyboard_End
        CurrentUser.bGameLostFocus = True
        Exit Sub
    ElseIf CurrentUser.bGameLostFocus Then
        CurrentUser.bGameLostFocus = False
        Call Main_Logic
    End If
    
    If Check_Main_Keys() Then
        'cri cri
    ElseIf (frmMain.Engine.Input_Keyboard_Last_KeyDOWN(DIK_RETURN) And frmMain.Engine.Input_Keyboard_KeyUP(DIK_RETURN)) Or _
        (frmMain.Engine.Input_Keyboard_Last_KeyDOWN(DIK_NUMPADENTER) And frmMain.Engine.Input_Keyboard_KeyUP(DIK_NUMPADENTER)) Then
        
        If Not SendTxt.Visible Then
            If Not frmCantidad.Visible Then
                Call CompletarEnvioMensajes
                SendTxt.Visible = True
                SendTxt.SetFocus
            End If
        Else
            Call Enviar_SendTxt
        End If
    ElseIf SendTxt.Visible Then
        SendTxt.SetFocus
        'Cri cri
    ElseIf frmMain.Engine.Input_Keyboard_KeyDOWN(DIK_NUMLOCK) Then
        CurrentUser.AutoNavigation = Not CurrentUser.AutoNavigation
        Call Check_Main_Keys
    ElseIf Accionar() Then
        'Cri cri
    ElseIf frmMain.Engine.Input_Keyboard_KeyDOWN(DIK_F1) Then
        BotonElegido = 1
        If MacroKeys(BotonElegido).TipoAccion = 0 Then
            frmBindKey.Show vbModeless, frmMain
        Else
            Call Bind_Accion(BotonElegido)
        End If
    ElseIf frmMain.Engine.Input_Keyboard_KeyDOWN(DIK_F2) Then
        BotonElegido = 2
        If MacroKeys(BotonElegido).TipoAccion = 0 Then
            frmBindKey.Show vbModeless, frmMain
        Else
            Call Bind_Accion(BotonElegido)
        End If
    ElseIf frmMain.Engine.Input_Keyboard_KeyDOWN(DIK_F3) Then
        BotonElegido = 3
        If MacroKeys(BotonElegido).TipoAccion = 0 Then
            frmBindKey.Show vbModeless, frmMain
        Else
            Call Bind_Accion(BotonElegido)
        End If
    ElseIf frmMain.Engine.Input_Keyboard_KeyDOWN(DIK_F4) Then
        BotonElegido = 4
        If MacroKeys(BotonElegido).TipoAccion = 0 Then
            frmBindKey.Show vbModeless, frmMain
        Else
            Call Bind_Accion(BotonElegido)
        End If
    ElseIf frmMain.Engine.Input_Keyboard_KeyDOWN(DIK_F5) Then
        BotonElegido = 5
        If MacroKeys(BotonElegido).TipoAccion = 0 Then
            frmBindKey.Show vbModeless, frmMain
        Else
            Call Bind_Accion(BotonElegido)
        End If
    ElseIf frmMain.Engine.Input_Keyboard_KeyDOWN(DIK_F6) Then
        BotonElegido = 6
        If MacroKeys(BotonElegido).TipoAccion = 0 Then
            frmBindKey.Show vbModeless, frmMain
        Else
            Call Bind_Accion(BotonElegido)
        End If
    ElseIf frmMain.Engine.Input_Keyboard_KeyDOWN(DIK_F7) Then
        BotonElegido = 7
        If MacroKeys(BotonElegido).TipoAccion = 0 Then
            frmBindKey.Show vbModeless, frmMain
        Else
            Call Bind_Accion(BotonElegido)
        End If
    ElseIf frmMain.Engine.Input_Keyboard_KeyDOWN(DIK_F8) Then
        BotonElegido = 8
        If MacroKeys(BotonElegido).TipoAccion = 0 Then
            frmBindKey.Show vbModeless, frmMain
        Else
            Call Bind_Accion(BotonElegido)
        End If
    ElseIf frmMain.Engine.Input_Keyboard_KeyDOWN(DIK_F9) Then
        BotonElegido = 9
        If MacroKeys(BotonElegido).TipoAccion = 0 Then
            frmBindKey.Show vbModeless, frmMain
        Else
            Call Bind_Accion(BotonElegido)
        End If
    ElseIf frmMain.Engine.Input_Keyboard_KeyDOWN(DIK_F10) Then
        BotonElegido = 10
        If MacroKeys(BotonElegido).TipoAccion = 0 Then
            frmBindKey.Show vbModeless, frmMain
        Else
            Call Bind_Accion(BotonElegido)
        End If
    ElseIf frmMain.Engine.Input_Keyboard_KeyDOWN(DIK_F11) Then
        BotonElegido = 11
        If MacroKeys(BotonElegido).TipoAccion = 0 Then
            frmBindKey.Show vbModeless, frmMain
        Else
            Call Bind_Accion(BotonElegido)
        End If
    ElseIf frmMain.Engine.Input_Keyboard_KeyDOWN(DIK_ESCAPE) And CurrentUser.Saliendo Then
        Call ClientTCP.Send_Data(Cancel_Exit)
    End If
    
    frmMain.Engine.Input_Keyboard_End
    
ElseIf eventid = frmMain.Engine.Input_Mouse_DXEvent Then
    
    If GetActiveWindow <> hwnd Then
        CurrentUser.bGameLostFocus = True
        Exit Sub
    ElseIf CurrentUser.bGameLostFocus Then
        CurrentUser.bGameLostFocus = False
        Call Main_Logic
    End If
    
    Select Case frmMain.Engine.Input_Mouse_Start
        Case -1
            frmMain.Engine.Input_Mouse_Map_Get map_x, map_y
            If frmMain.Engine.Pos_In_Current_Area(map_x, map_y) And frmMain.Engine.Map_In_Legal_Bounds(map_x, map_y) Then Call MouseLeftClick(map_x, map_y)
        
        Case -2
            frmMain.Engine.Input_Mouse_Map_Get map_x, map_y
            If frmMain.Engine.Pos_In_Current_Area(map_x, map_y) And frmMain.Engine.Map_In_Legal_Bounds(map_x, map_y) Then Call MouseRightClick(map_x, map_y)
    
    End Select
End If

Exit Sub

Error_Handler:
    Debug.Print "DXCallback: " & Err.Description & " - " & Err.Number

End Sub

Public Function Check_Main_Keys() As Boolean

Dim CurDir As Byte
CurDir = frmMain.Engine.Char_Current_Direction_Get

'Start moving up
If frmMain.Engine.Input_Keyboard_KeyDOWN(BindKeys(14).VirtualKey) And CurDir <> NORTH Then
    frmMain.Engine.Char_Current_Direction_Set (NORTH)
    Check_Main_Keys = True
'Start moving right
ElseIf frmMain.Engine.Input_Keyboard_KeyDOWN(BindKeys(17).VirtualKey) And CurDir <> EAST Then
    frmMain.Engine.Char_Current_Direction_Set (EAST)
    Check_Main_Keys = True
'Start moving down
ElseIf frmMain.Engine.Input_Keyboard_KeyDOWN(BindKeys(15).VirtualKey) And CurDir <> SOUTH Then
    frmMain.Engine.Char_Current_Direction_Set (SOUTH)
    Check_Main_Keys = True
'Start moving left
ElseIf frmMain.Engine.Input_Keyboard_KeyDOWN(BindKeys(16).VirtualKey) And CurDir <> WEST Then
    frmMain.Engine.Char_Current_Direction_Set (WEST)
    Check_Main_Keys = True
ElseIf CurDir > 0 And CurrentUser.AutoNavigation = False Then
    'STOP moving up
    If frmMain.Engine.Input_Keyboard_KeyUP(BindKeys(14).VirtualKey) And CurDir = NORTH Then
        frmMain.Engine.Char_Current_Direction_Set (0)
        Check_Main_Keys = True
        Call Check_Main_Keys
    'STOP moving right
    ElseIf frmMain.Engine.Input_Keyboard_KeyUP(BindKeys(17).VirtualKey) And CurDir = EAST Then
        frmMain.Engine.Char_Current_Direction_Set (0)
        Check_Main_Keys = True
        Call Check_Main_Keys
    'STOP moving down
    ElseIf frmMain.Engine.Input_Keyboard_KeyUP(BindKeys(15).VirtualKey) And CurDir = SOUTH Then
        frmMain.Engine.Char_Current_Direction_Set (0)
        Check_Main_Keys = True
        Call Check_Main_Keys
    'STOP moving left
    ElseIf frmMain.Engine.Input_Keyboard_KeyUP(BindKeys(16).VirtualKey) And CurDir = WEST Then
        frmMain.Engine.Char_Current_Direction_Set (0)
        Check_Main_Keys = True
        Call Check_Main_Keys
    End If
End If
        
End Function

Private Function Check_Background_Keys() As Boolean

Dim CurDir As Byte
CurDir = frmMain.Engine.Char_Current_Direction_Get
        
If CurDir > 0 Then
    'STOP moving up
    If frmMain.Engine.Input_Keyboard_KeyUP(BindKeys(14).VirtualKey) And CurDir = NORTH Then
        frmMain.Engine.Char_Current_Direction_Set (0)
        Check_Background_Keys = True
    'STOP moving right
    ElseIf frmMain.Engine.Input_Keyboard_KeyUP(BindKeys(17).VirtualKey) And CurDir = EAST Then
        frmMain.Engine.Char_Current_Direction_Set (0)
        Check_Background_Keys = True
    'STOP moving down
    ElseIf frmMain.Engine.Input_Keyboard_KeyUP(BindKeys(15).VirtualKey) And CurDir = SOUTH Then
        frmMain.Engine.Char_Current_Direction_Set (0)
        Check_Background_Keys = True
    'STOP moving left
    ElseIf frmMain.Engine.Input_Keyboard_KeyUP(BindKeys(16).VirtualKey) And CurDir = WEST Then
        frmMain.Engine.Char_Current_Direction_Set (0)
        Check_Background_Keys = True
    End If
End If
        
End Function

Private Sub Winsock_Error_Close_Event(Optional ByVal Description As String, Optional ByVal Number As Long)

On Error Resume Next

Dim bFlag As Boolean

CurrentUser.RDBuffer = vbNullString

If CurrentUser.EndingGame Then
    Call EndGame(True)
    Exit Sub
End If

If frmMensaje.Visible And frmConnect.Visible = False Then
    frmMensaje.Visible = False
    bFlag = True
End If

If CurrentUser.Logged Then
    
    Call ResetCurrentUser
    
    frmConnect.Visible = True
    Me.Visible = False
    frmCharList.Visible = False
    frmIniciando.Visible = False
    frmCrearPersonaje.Visible = False
    
    Call FormParser.Parse_Form(frmConnect)
    
    If sMusica <> CONST_DESHABILITADA Then
        Sound.NextMusic = MUS_VolverInicio
        Sound.Fading = 200
    End If
    
Else
    frmConnect.Visible = True
    Me.Visible = False
    frmCrearPersonaje.Visible = False
    frmCharList.Visible = False
    frmIniciando.Visible = False
    Call FormParser.Parse_Form(frmConnect)
End If

If bFlag Then
    If frmConnect.Visible Then
        frmMensaje.Show vbModal, frmConnect
    ElseIf frmCharList.Visible Then
        frmMensaje.Show vbModal, frmCharList
    End If
Else
    If Not frmMensaje.Visible Then _
        If LenB(Description) > 0 Then _
            Call MsgBox(Locale_GUI_Frase(345) & " (" & Description & " - " & Number & ")", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error al conectar")
End If

End Sub

Private Sub publi_GotFocus()

'If frmMain.Engine.Input_Mouse_Button_Right_Get Then
'    Pubilicidad_Deshabilitada = True
'    Call Banner_Logic
'End If

End Sub

Private Sub publi_DownloadComplete()

On Error Resume Next

If CurrentUser.Logged = False Then Exit Sub

Publicidad_Cargada = True
'publi.Document.body.style.BorderStyle = "None"

Call Banner_Logic

End Sub
