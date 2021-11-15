VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BorderStyle     =   0  'None
   Caption         =   "ImperiumAO 1.3"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCrearPersonaje.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1710
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   49
      Top             =   4530
      Width           =   375
   End
   Begin VB.PictureBox picFamiliar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   10380
      ScaleHeight     =   1185
      ScaleWidth      =   840
      TabIndex        =   48
      Top             =   1575
      Width           =   870
   End
   Begin VB.TextBox txtFamiliar 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9330
      MaxLength       =   30
      TabIndex        =   4
      Top             =   990
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.ComboBox lstFamiliar 
      BackColor       =   &H00000000&
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
      Height          =   285
      ItemData        =   "frmCrearPersonaje.frx":000C
      Left            =   8520
      List            =   "frmCrearPersonaje.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1860
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.ComboBox lstProfesion 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":0010
      Left            =   870
      List            =   "frmCrearPersonaje.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2490
      Width           =   2055
   End
   Begin VB.ComboBox lstGenero 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":0014
      Left            =   870
      List            =   "frmCrearPersonaje.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3150
      Width           =   2055
   End
   Begin VB.ComboBox lstRaza 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":0018
      Left            =   870
      List            =   "frmCrearPersonaje.frx":001A
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3810
      Width           =   2055
   End
   Begin VB.ComboBox lstHogar 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":001C
      Left            =   8550
      List            =   "frmCrearPersonaje.frx":0035
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3585
      Width           =   2745
   End
   Begin VB.TextBox txtNombre 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2100
      MaxLength       =   30
      TabIndex        =   0
      Top             =   1050
      Width           =   5865
   End
   Begin VB.Image imgCabeza 
      Height          =   600
      Index           =   1
      Left            =   2130
      Tag             =   "0"
      Top             =   4425
      Width           =   390
   End
   Begin VB.Image imgCabeza 
      Height          =   600
      Index           =   0
      Left            =   1260
      Tag             =   "0"
      Top             =   4425
      Width           =   390
   End
   Begin VB.Label lblFamiInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Descropcion del familiar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   8550
      TabIndex        =   47
      Top             =   2295
      Width           =   1635
   End
   Begin VB.Image imgNoDisp 
      Height          =   2145
      Left            =   8415
      Top             =   780
      Width           =   3045
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   525
      Left            =   2610
      TabIndex        =   46
      Top             =   8220
      Width           =   6795
   End
   Begin VB.Image ImgAtributoMenos 
      Height          =   135
      Index           =   4
      Left            =   2700
      Top             =   7230
      Width           =   195
   End
   Begin VB.Image ImgAtributoMenos 
      Height          =   135
      Index           =   3
      Left            =   2700
      Top             =   6900
      Width           =   195
   End
   Begin VB.Image ImgAtributoMenos 
      Height          =   135
      Index           =   2
      Left            =   2700
      Top             =   6540
      Width           =   195
   End
   Begin VB.Image ImgAtributoMenos 
      Height          =   135
      Index           =   1
      Left            =   2700
      Top             =   6180
      Width           =   195
   End
   Begin VB.Image ImgAtributoMas 
      Height          =   135
      Index           =   4
      Left            =   2700
      Top             =   7080
      Width           =   195
   End
   Begin VB.Image ImgAtributoMas 
      Height          =   135
      Index           =   3
      Left            =   2700
      Top             =   6750
      Width           =   195
   End
   Begin VB.Image ImgAtributoMas 
      Height          =   135
      Index           =   2
      Left            =   2700
      Top             =   6390
      Width           =   195
   End
   Begin VB.Image ImgAtributoMas 
      Height          =   135
      Index           =   1
      Left            =   2700
      Top             =   6030
      Width           =   195
   End
   Begin VB.Label lbAtt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
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
      Height          =   195
      Index           =   4
      Left            =   2445
      TabIndex        =   45
      Top             =   7140
      Width           =   240
   End
   Begin VB.Label lbAtt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
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
      Height          =   195
      Index           =   3
      Left            =   2445
      TabIndex        =   44
      Top             =   6780
      Width           =   240
   End
   Begin VB.Label lbAtt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
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
      Height          =   195
      Index           =   2
      Left            =   2445
      TabIndex        =   43
      Top             =   6420
      Width           =   240
   End
   Begin VB.Label lbAtt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
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
      Height          =   195
      Index           =   1
      Left            =   2445
      TabIndex        =   42
      Top             =   6060
      Width           =   240
   End
   Begin VB.Label lbBonificador 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Height          =   195
      Index           =   4
      Left            =   2145
      TabIndex        =   41
      Top             =   7140
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lbBonificador 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Height          =   195
      Index           =   3
      Left            =   2145
      TabIndex        =   40
      Top             =   6780
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lbBonificador 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Height          =   195
      Index           =   2
      Left            =   2145
      TabIndex        =   39
      Top             =   6420
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lbBonificador 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Height          =   195
      Index           =   1
      Left            =   2145
      TabIndex        =   38
      Top             =   6060
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image ImgAtributoMenos 
      Height          =   135
      Index           =   0
      Left            =   2700
      Top             =   5820
      Width           =   195
   End
   Begin VB.Image ImgAtributoMas 
      Height          =   135
      Index           =   0
      Left            =   2700
      Top             =   5670
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   53
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":0074
      Top             =   6930
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   52
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":01C6
      Top             =   6810
      Width           =   195
   End
   Begin VB.Label Skill 
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
      Index           =   26
      Left            =   7365
      TabIndex        =   37
      Top             =   6840
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   51
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":0318
      Top             =   6540
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   50
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":046A
      Top             =   6420
      Width           =   195
   End
   Begin VB.Label Skill 
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
      Index           =   25
      Left            =   7365
      TabIndex        =   36
      Top             =   6450
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   49
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":05BC
      Top             =   6180
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   48
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":070E
      Top             =   6060
      Width           =   195
   End
   Begin VB.Label Skill 
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
      Index           =   24
      Left            =   7365
      TabIndex        =   35
      Top             =   6090
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   47
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":0860
      Top             =   5790
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   46
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":09B2
      Top             =   5670
      Width           =   195
   End
   Begin VB.Label Skill 
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
      Index           =   23
      Left            =   7365
      TabIndex        =   34
      Top             =   5700
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   45
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":0B04
      Top             =   5430
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   44
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":0C56
      Top             =   5280
      Width           =   195
   End
   Begin VB.Label Skill 
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
      Index           =   22
      Left            =   7365
      TabIndex        =   33
      Top             =   5340
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   43
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":0DA8
      Top             =   5040
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   42
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":0EFA
      Top             =   4920
      Width           =   195
   End
   Begin VB.Label Skill 
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
      Index           =   21
      Left            =   7365
      TabIndex        =   32
      Top             =   4950
      Width           =   240
   End
   Begin VB.Label lbBonificador 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Height          =   195
      Index           =   0
      Left            =   2145
      TabIndex        =   31
      Top             =   5700
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lbAtributos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   225
      Left            =   2520
      TabIndex        =   30
      Top             =   7500
      Width           =   255
   End
   Begin VB.Label Skill 
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
      Index           =   18
      Left            =   7365
      TabIndex        =   29
      Top             =   3840
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   19
      Left            =   7365
      TabIndex        =   28
      Top             =   4200
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   20
      Left            =   7365
      TabIndex        =   27
      Top             =   4590
      Width           =   240
   End
   Begin VB.Image imgAccion 
      Height          =   615
      Index           =   0
      Left            =   9585
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   8175
      Width           =   1755
   End
   Begin VB.Image Image1 
      Height          =   3570
      Left            =   8490
      Stretch         =   -1  'True
      Top             =   4230
      Width           =   2835
   End
   Begin VB.Label Puntos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "10"
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
      Height          =   225
      Left            =   6795
      TabIndex        =   26
      Top             =   7260
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   3
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":104C
      Top             =   2790
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   5
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":119E
      Top             =   3180
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   7
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":12F0
      Top             =   3540
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   9
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":1442
      Top             =   3930
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   11
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":1594
      Top             =   4290
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   13
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":16E6
      Top             =   4680
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   15
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":1838
      Top             =   5040
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   17
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":198A
      Top             =   5430
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   19
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":1ADC
      Top             =   5790
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   21
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":1C2E
      Top             =   6180
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   23
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":1D80
      Top             =   6540
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   25
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":1ED2
      Top             =   6930
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   27
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":2024
      Top             =   7290
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   1
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":2176
      Top             =   2430
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   0
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":22C8
      Top             =   2280
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   2
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":241A
      Top             =   2670
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   4
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":256C
      Top             =   3060
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   6
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":26BE
      Top             =   3420
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   8
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":2810
      Top             =   3780
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   10
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":2962
      Top             =   4170
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   12
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":2AB4
      Top             =   4560
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   14
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":2C06
      Top             =   4920
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   16
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":2D58
      Top             =   5280
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   18
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":2EAA
      Top             =   5670
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   20
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":2FFC
      Top             =   6060
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   22
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":314E
      Top             =   6420
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   24
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":32A0
      Top             =   6810
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   26
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":33F2
      Top             =   7170
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   28
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":3544
      Top             =   2280
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   29
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":3696
      Top             =   2430
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   30
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":37E8
      Top             =   2670
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   31
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":393A
      Top             =   2820
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   32
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":3A8C
      Top             =   3060
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   33
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":3BDE
      Top             =   3180
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   34
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":3D30
      Top             =   3420
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   35
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":3E82
      Top             =   3570
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   36
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":3FD4
      Top             =   3780
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   37
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":4126
      Top             =   3930
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   38
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":4278
      Top             =   4170
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   39
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":43CA
      Top             =   4320
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   40
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":451C
      Top             =   4560
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   41
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":466E
      Top             =   4680
      Width           =   195
   End
   Begin VB.Image imgAccion 
      Height          =   615
      Index           =   1
      Left            =   660
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   8175
      Width           =   1755
   End
   Begin VB.Label Skill 
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
      Index           =   17
      Left            =   7365
      TabIndex        =   25
      Top             =   3450
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   16
      Left            =   7365
      TabIndex        =   24
      Top             =   3090
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   15
      Left            =   7365
      TabIndex        =   23
      Top             =   2700
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   14
      Left            =   7365
      TabIndex        =   22
      Top             =   2340
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   13
      Left            =   5310
      TabIndex        =   21
      Top             =   7200
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   12
      Left            =   5310
      TabIndex        =   20
      Top             =   6840
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   11
      Left            =   5310
      TabIndex        =   19
      Top             =   6450
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   10
      Left            =   5310
      TabIndex        =   18
      Top             =   6090
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   9
      Left            =   5310
      TabIndex        =   17
      Top             =   5700
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   8
      Left            =   5310
      TabIndex        =   16
      Top             =   5340
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   7
      Left            =   5310
      TabIndex        =   15
      Top             =   4950
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   6
      Left            =   5310
      TabIndex        =   14
      Top             =   4590
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   5
      Left            =   5310
      TabIndex        =   13
      Top             =   4200
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   4
      Left            =   5310
      TabIndex        =   12
      Top             =   3840
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   3
      Left            =   5310
      TabIndex        =   11
      Top             =   3450
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   2
      Left            =   5310
      TabIndex        =   10
      Top             =   3090
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   0
      Left            =   5310
      TabIndex        =   9
      Top             =   2340
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Index           =   1
      Left            =   5310
      TabIndex        =   8
      Top             =   2700
      Width           =   240
   End
   Begin VB.Label lbAtt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
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
      Height          =   195
      Index           =   0
      Left            =   2445
      TabIndex        =   7
      Top             =   5700
      Width           =   240
   End
End
Attribute VB_Name = "frmCrearPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmCrearPersonaje - ImperiumAO - v1.4.5 R5
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

Private SkillPoints As Byte
Private Atributos As Byte
Private Const ATT_INICIALES = 40
Private intHeadInd As Integer

Private Function CheckData() As Boolean

If CurrentUser.UserName = vbNullString Then
    lblInfo.Caption = Locale_GUI_Frase(177)
    Exit Function
End If

If Len(CurrentUser.UserName) > 30 Then
    lblInfo.Caption = Locale_GUI_Frase(178)
    Exit Function
End If

If CurrentUser.UserRaza = 0 Then
    lblInfo.Caption = Locale_GUI_Frase(179)
    Exit Function
End If

If CurrentUser.UserSexo = 0 Then
    lblInfo.Caption = Locale_GUI_Frase(180)
    Exit Function
End If

If CurrentUser.UserClase = 0 Then
    lblInfo.Caption = Locale_GUI_Frase(181)
    Exit Function
End If

If CurrentUser.UserHogar = 0 Then
    lblInfo.Caption = Locale_GUI_Frase(182)
    Exit Function
End If

If SkillPoints > 0 Then
    lblInfo.Caption = Locale_GUI_Frase(183)
    Exit Function
End If

If Atributos > 0 Then
    lblInfo.Caption = Locale_GUI_Frase(184)
    Exit Function
End If

If frmCrearPersonaje.lstFamiliar.Visible = True Then

    If CurrentUser.UserPet.Tipo = vbNullString Then
        lblInfo.Caption = Locale_GUI_Frase(185)
        Exit Function
    ElseIf CurrentUser.UserPet.nombre = vbNullString Then
        lblInfo.Caption = Locale_GUI_Frase(186)
        Exit Function
    ElseIf Len(CurrentUser.UserPet.nombre) > 30 Then
        lblInfo.Caption = Locale_GUI_Frase(187)
        Exit Function
    End If

End If

Dim i As Integer
For i = 1 To NUMATRIBUTOS
    CurrentUser.UserAtributos(i) = Val(lbAtt(i - 1).Caption)
    If CurrentUser.UserAtributos(i) = 0 Then
        lblInfo.Caption = Locale_GUI_Frase(188)
        Exit Function
    End If
Next i

CheckData = True

End Function

Private Sub imgAccion_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If FormParser.GetDefaultCursor(Me) = E_WAIT Then Exit Sub

Call Sound.Sound_Play(SND_CLICK)
Call imgAccionRestaurar
    
Select Case Index
    Case 0
        
        Dim i As Integer
        
        For i = 0 To 26
            CurrentUser.UserSkills(SkillRealToIndex(i + 1)) = Val(Skill(i).Caption)
        Next i
        
        CurrentUser.UserName = Trim$(txtNombre.Text)
        CurrentUser.UserRaza = lstRaza.ListIndex
        CurrentUser.UserSexo = lstGenero.ListIndex
        CurrentUser.UserClase = lstProfesion.ListIndex
        CurrentUser.UserPet.Tipo = lstFamiliar.List(lstFamiliar.ListIndex)
        CurrentUser.UserPet.nombre = frmCrearPersonaje.txtFamiliar.Text
        CurrentUser.UserHogar = lstHogar.ListIndex
        Atributos = Val(lbAtributos.Caption)
        
        If CheckData() Then
            Call ClientTCP.CreateNewChar(ValidarLoginMSG(CInt(bRK)), intHeadInd)
            Call FormParser.Parse_Form(Me, E_WAIT)
        End If
        
    Case 1
        If sMusica <> CONST_DESHABILITADA Then
            Sound.NextMusic = MUS_VolverInicio
            Sound.Fading = 200
        End If
        
        Call FormParser.Parse_Form(frmCharList)
        frmCharList.Show
        
        Unload Me

End Select

End Sub

Private Sub imgAccion_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Select Case Index
    Case 0 'Crear
        imgAccion(0).Picture = General_Load_Picture_From_Resource_Ex("creardown")
        imgAccion(0).Tag = "0"
    Case 1 'Volver
        imgAccion(1).Picture = General_Load_Picture_From_Resource_Ex("volverdown")
        imgAccion(1).Tag = "0"
End Select

End Sub

Private Sub imgAccion_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Select Case Index
    Case 0 'Crear personaje
        If imgAccion(0).Tag = "1" Then
            imgAccion(0).Picture = General_Load_Picture_From_Resource_Ex("crearover")
            imgAccion(0).Tag = "0"
        End If
        
    Case 1 'Volver
        If imgAccion(1).Tag = "1" Then
            imgAccion(1).Picture = General_Load_Picture_From_Resource_Ex("volverover")
            imgAccion(1).Tag = "0"
        End If
End Select

Call imgAccionRestaurar(Index)

End Sub

Private Sub imgCabezaRestaurar(Optional ByVal NoIndex As Integer = 1000, Optional ByVal Over As Boolean = False)

Dim i As Integer

For i = 0 To 1
    If i <> NoIndex Then
        imgCabeza(i).Picture = Nothing
        imgCabeza(i).Tag = "1"
    ElseIf Over Then
        If i = 0 Then
            imgCabeza(0).Picture = General_Load_Picture_From_Resource_Ex("izqover")
            imgCabeza(0).Tag = "0"
        Else
            imgCabeza(1).Picture = General_Load_Picture_From_Resource_Ex("derover")
            imgCabeza(1).Tag = "0"
        End If
    End If
Next i

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim i As Integer

For i = 0 To 1
    If imgAccion(i).Tag = "0" Then
        imgAccion(i).Picture = Nothing
        imgAccion(i).Tag = "1"
    End If
Next i

For i = 0 To 1
    If imgCabeza(i).Tag = "0" Then
        imgCabeza(i).Picture = Nothing
        imgCabeza(i).Tag = "1"
    End If
Next i

End Sub

Private Sub Command1_Click(Index As Integer)

Dim indice
If Index Mod 2 = 0 Then
    If SkillPoints > 0 Then
        indice = Index \ 2
        Skill(indice).Caption = Val(Skill(indice).Caption) + 1
        SkillPoints = SkillPoints - 1
    End If
Else
    If SkillPoints < 10 Then
        
        indice = Index \ 2
        If Val(Skill(indice).Caption) > 0 Then
            Skill(indice).Caption = Val(Skill(indice).Caption) - 1
            SkillPoints = SkillPoints + 1
        End If
    End If
End If

Puntos.Caption = SkillPoints
End Sub

Private Sub Form_Load()

SkillPoints = 10
Puntos.Caption = SkillPoints

Me.Caption = Form_Caption
Me.Picture = General_Load_Picture_From_Resource_Ex("cp-interface")

Dim i As Integer

lstProfesion.Clear
lstProfesion.AddItem vbNullString
For i = LBound(ListaClases) To UBound(ListaClases)
    lstProfesion.AddItem ListaClases(i)
Next i

lstGenero.Clear
lstGenero.AddItem vbNullString
lstGenero.AddItem Locale_GUI_Frase(229)
lstGenero.AddItem Locale_GUI_Frase(230)

lstRaza.Clear
lstRaza.AddItem vbNullString
For i = LBound(ListaRazas) To UBound(ListaRazas)
    lstRaza.AddItem ListaRazas(i)
Next i

lstProfesion.ListIndex = 0
lstGenero.ListIndex = 0
lstRaza.ListIndex = 0
lstHogar.ListIndex = 0

Image1.Picture = General_Load_Picture_From_Resource_Ex(LCase$(lstProfesion.Text) & vbNullString)
Call ResetAtributos
Call FormParser.Parse_Form(Me)

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If (Button = vbLeftButton) And (RunWindowed = 1) Then Call Auto_Drag(Me.hwnd)
End Sub

Private Sub ImgAtributoMas_Click(Index As Integer)

If Val(lbAtt(Index).Caption) >= 18 Or Val(lbAtributos.Caption) <= 0 Then Exit Sub
    
lbAtt(Index).Caption = Val(lbAtt(Index).Caption) + 1
lbAtributos.Caption = lbAtributos.Caption - 1

End Sub

Private Sub ImgAtributoMenos_Click(Index As Integer)

If Val(lbAtt(Index).Caption) <= 6 Then Exit Sub

lbAtt(Index).Caption = Val(lbAtt(Index).Caption) - 1
lbAtributos.Caption = lbAtributos.Caption + 1

End Sub

Private Sub imgCabeza_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If lstRaza.ListIndex <= 0 Then Exit Sub
If lstGenero.ListIndex <= 0 Then Exit Sub

Call imgCabezaRestaurar(Index, True)

Select Case Index
    Case 0 'Izq
        intHeadInd = intHeadInd - 1
    Case 1 'Der
        intHeadInd = intHeadInd + 1
End Select

If lstGenero.ListIndex = Masculino Then
    If intHeadInd > Head_Range(lstRaza.ListIndex).mEnd Then intHeadInd = Head_Range(lstRaza.ListIndex).mStart
    If intHeadInd < Head_Range(lstRaza.ListIndex).mStart Then intHeadInd = Head_Range(lstRaza.ListIndex).mEnd
Else
    If intHeadInd > Head_Range(lstRaza.ListIndex).fEnd Then intHeadInd = Head_Range(lstRaza.ListIndex).fStart
    If intHeadInd < Head_Range(lstRaza.ListIndex).fStart Then intHeadInd = Head_Range(lstRaza.ListIndex).fEnd
End If

Call frmMain.Engine.Grh_Render_Head_To_Hdc(intHeadInd, picHead.hDC, 4, 4)
picHead.Refresh

End Sub

Private Sub imgCabeza_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If lstRaza.ListIndex <= 0 Then Exit Sub
If lstGenero.ListIndex <= 0 Then Exit Sub

Call Sound.Sound_Play(SND_CLICK)

Select Case Index
    Case 0 'Izq
        imgCabeza(0).Picture = General_Load_Picture_From_Resource_Ex("izqdown")
        imgCabeza(0).Tag = "0"
    Case 1 'Der
        imgCabeza(1).Picture = General_Load_Picture_From_Resource_Ex("derdown")
        imgCabeza(1).Tag = "0"
End Select

End Sub

Private Sub imgCabeza_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Select Case Index
    Case 0 'Izq
        If imgCabeza(0).Tag = "1" Then
            imgCabeza(0).Picture = General_Load_Picture_From_Resource_Ex("izqover")
            imgCabeza(0).Tag = "0"
        End If
        
    Case 1 'Der
        If imgCabeza(1).Tag = "1" Then
            imgCabeza(1).Picture = General_Load_Picture_From_Resource_Ex("derover")
            imgCabeza(1).Tag = "0"
        End If
End Select

Call imgCabezaRestaurar(Index)

End Sub

Private Sub imgAccionRestaurar(Optional ByVal NoIndex As Integer = 1000)

Dim i As Integer

For i = 0 To 1
    If i <> NoIndex Then
        imgAccion(i).Picture = Nothing
        imgAccion(i).Tag = "1"
    End If
Next i

End Sub

Private Sub lbAtributos_Click()
Call Sound.Sound_Play(SND_DICE)
Call ResetAtributos
End Sub

Private Sub lblInfo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call imgAccionRestaurar
Call imgCabezaRestaurar
End Sub

Private Sub lstFamiliar_Click()

If lstFamiliar.ListIndex > 0 Then
    lblFamiInfo.Caption = ListaFamiliares(lstFamiliar.ListIndex).Desc
    picFamiliar.Picture = General_Load_Picture_From_Resource_Ex(ListaFamiliares(lstFamiliar.ListIndex).Imagen)
Else
    lblFamiInfo.Caption = Locale_GUI_Frase(189)
    picFamiliar.Picture = Nothing
End If

End Sub

Private Sub lstHogar_Click()
    If lstHogar.Text = "Nix" Or lstHogar.Text = "Ullathorpe" Or lstHogar.Text = "Banderbill" Then
        lblInfo.Caption = Locale_GUI_Frase(190)
    ElseIf lstHogar.Text = "Lindos" Or lstHogar.Text = "Illiandor" Or lstHogar.Text = "Suramei" Then
        lblInfo.Caption = Locale_GUI_Frase(191)
    End If
End Sub

Private Sub lstProfesion_Click()

On Error Resume Next

Image1.Picture = General_Load_Picture_From_Resource_Ex(LCase$(lstProfesion.Text) & vbNullString)

If lstProfesion.Text = "Mago" Then
    frmCrearPersonaje.txtFamiliar.Visible = True
    frmCrearPersonaje.lstFamiliar.Visible = True
    imgNoDisp.Picture = Nothing
    lblFamiInfo.Visible = True
    picFamiliar.Visible = True
    Call CambioFamiliar(5)
ElseIf lstProfesion.Text = "Cazador" Or lstProfesion.Text = "Druida" Then
    frmCrearPersonaje.txtFamiliar.Visible = True
    frmCrearPersonaje.lstFamiliar.Visible = True
    imgNoDisp.Picture = Nothing
    lblFamiInfo.Visible = True
    picFamiliar.Visible = True
    Call CambioFamiliar(4)
Else
    frmCrearPersonaje.txtFamiliar.Visible = False
    frmCrearPersonaje.lstFamiliar.Visible = False
    imgNoDisp.Picture = General_Load_Picture_From_Resource_Ex("mascotanodisp")
    picFamiliar.Visible = False
    lblFamiInfo.Visible = False
End If

End Sub

Private Sub txtNombre_Change()
txtNombre.Text = LTrim(txtNombre.Text)
End Sub

Private Sub txtNombre_GotFocus()
lblInfo.Caption = Locale_GUI_Frase(192)
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
End Sub

Private Sub CambioFamiliar(ByVal NumFamiliares As Integer)

If NumFamiliares = 5 Then

    ReDim ListaFamiliares(1 To NumFamiliares) As tListaFamiliares
    ListaFamiliares(1).name = Locale_GUI_Frase(271)
    ListaFamiliares(1).Desc = Locale_GUI_Frase(272)
    ListaFamiliares(1).Imagen = "elefuego"
    
    ListaFamiliares(2).name = Locale_GUI_Frase(273)
    ListaFamiliares(2).Desc = Locale_GUI_Frase(274)
    ListaFamiliares(2).Imagen = "eleagua"
    
    ListaFamiliares(3).name = Locale_GUI_Frase(275)
    ListaFamiliares(3).Desc = Locale_GUI_Frase(276)
    ListaFamiliares(3).Imagen = "eletierra"
    
    ListaFamiliares(4).name = Locale_GUI_Frase(277)
    ListaFamiliares(4).Desc = Locale_GUI_Frase(278)
    ListaFamiliares(4).Imagen = "ely"
    
    ListaFamiliares(5).name = Locale_GUI_Frase(279)
    ListaFamiliares(5).Desc = Locale_GUI_Frase(280)
    ListaFamiliares(5).Imagen = "fatuo"
    
Else

    ReDim ListaFamiliares(1 To NumFamiliares) As tListaFamiliares
    ListaFamiliares(1).name = Locale_GUI_Frase(281)
    ListaFamiliares(1).Desc = Locale_GUI_Frase(282)
    ListaFamiliares(1).Imagen = "tigre"
    
    ListaFamiliares(2).name = Locale_GUI_Frase(283)
    ListaFamiliares(2).Desc = Locale_GUI_Frase(284)
    ListaFamiliares(2).Imagen = "lobo"
    
    ListaFamiliares(3).name = Locale_GUI_Frase(285)
    ListaFamiliares(3).Desc = Locale_GUI_Frase(286)
    ListaFamiliares(3).Imagen = "oso"
    
    ListaFamiliares(4).name = Locale_GUI_Frase(287)
    ListaFamiliares(4).Desc = Locale_GUI_Frase(288)
    ListaFamiliares(4).Imagen = "ent"

End If

Dim i As Integer
lstFamiliar.Clear
lstFamiliar.AddItem vbNullString
For i = 1 To UBound(ListaFamiliares)
    lstFamiliar.AddItem ListaFamiliares(i).name
Next i

lstFamiliar.ListIndex = 0

End Sub

Private Sub txtfamiliar_GotFocus()
lblInfo.Caption = Locale_GUI_Frase(193)
End Sub

'Private Sub lbAtt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'If Button = vbLeftButton Then
'
'    If CurrentUser.UserAtributos(Index + 1) >= 18 Or Val(lbAtributos.Caption) <= 0 Then
'        Beep
'        Exit Sub
'    End If
'
'    lbAtt(Index).Caption = Val(lbAtt(Index).Caption) + 1
'    CurrentUser.UserAtributos(Index + 1) = Val(lbAtt(Index).Caption)
'    lbAtributos.Caption = lbAtributos.Caption - 1
'Else
'
'    If CurrentUser.UserAtributos(Index + 1) <= 6 Then
'        Beep
'        Exit Sub
'    End If
'
'    lbAtt(Index).Caption = Val(lbAtt(Index).Caption) - 1
'    CurrentUser.UserAtributos(Index + 1) = Val(lbAtt(Index).Caption)
'    lbAtributos.Caption = lbAtributos.Caption + 1
'End If
'
'End Sub

Private Sub lstRaza_Click()

Dim tmpInt As Integer

If lstRaza.List(lstRaza.ListIndex) = vbNullString Then Exit Sub

Dim i As Integer

For i = 1 To NUMATRIBUTOS
    tmpInt = BonificadorRaza(i, lstRaza.ListIndex)
    
    lbBonificador(i - 1).Caption = IIf(tmpInt > 0, "+" & CStr(tmpInt), CStr(tmpInt))
    If Val(lbBonificador(i - 1)) = 0 Then
        lbBonificador(i - 1).Visible = False
    Else
        lbBonificador(i - 1).Visible = True
    End If
Next i

If LenB(lstGenero.List(lstGenero.ListIndex)) > 0 Then
    
    If lstGenero.ListIndex = Masculino Then
        intHeadInd = CInt(General_Random_Number(Head_Range(lstRaza.ListIndex).mStart, Head_Range(lstRaza.ListIndex).mEnd))
    Else
        intHeadInd = CInt(General_Random_Number(Head_Range(lstRaza.ListIndex).fStart, Head_Range(lstRaza.ListIndex).fEnd))
    End If
    
    Call frmMain.Engine.Grh_Render_Head_To_Hdc(intHeadInd, picHead.hDC, 4, 4)
    picHead.Refresh
    
End If

End Sub

Private Sub lstGenero_Click()

If LenB(lstGenero.List(lstGenero.ListIndex)) = 0 Then Exit Sub

If LenB(lstRaza.List(lstRaza.ListIndex)) > 0 Then
    
    If lstGenero.ListIndex = Masculino Then
        intHeadInd = CInt(General_Random_Number(Head_Range(lstRaza.ListIndex).mStart, Head_Range(lstRaza.ListIndex).mEnd))
    Else
        intHeadInd = CInt(General_Random_Number(Head_Range(lstRaza.ListIndex).fStart, Head_Range(lstRaza.ListIndex).fEnd))
    End If
    
    Call frmMain.Engine.Grh_Render_Head_To_Hdc(intHeadInd, picHead.hDC, 4, 4)
    picHead.Refresh
    
End If

End Sub

Public Function BonificadorRaza(ByVal Atributo As Integer, ByVal Raza As Byte) As Integer

Select Case Atributo
    Case Fuerza
        If Raza = HUMANO Then BonificadorRaza = 1
        If Raza = DROW Then BonificadorRaza = 2
        If Raza = ENANO Then BonificadorRaza = 3
        If Raza = ELFO Then BonificadorRaza = 0
        If Raza = ORCO Then BonificadorRaza = 5
        If Raza = GNOMO Then BonificadorRaza = -5
    Case Agilidad
        If Raza = HUMANO Then BonificadorRaza = 1
        If Raza = DROW Then BonificadorRaza = 0
        If Raza = ENANO Then BonificadorRaza = -1
        If Raza = ELFO Then BonificadorRaza = 2
        If Raza = ORCO Then BonificadorRaza = -2
        If Raza = GNOMO Then BonificadorRaza = 3
    Case Inteligencia
        If Raza = HUMANO Then BonificadorRaza = 1
        If Raza = DROW Then BonificadorRaza = 2
        If Raza = ENANO Then BonificadorRaza = -7
        If Raza = ELFO Then BonificadorRaza = 3
        If Raza = ORCO Then BonificadorRaza = -6
        If Raza = GNOMO Then BonificadorRaza = 4
    Case Carisma
        If Raza = HUMANO Then BonificadorRaza = 0
        If Raza = DROW Then BonificadorRaza = -1
        If Raza = ENANO Then BonificadorRaza = -1
        If Raza = ELFO Then BonificadorRaza = 2
        If Raza = ORCO Then BonificadorRaza = -4
        If Raza = GNOMO Then BonificadorRaza = 0
    Case Constitucion
        If Raza = HUMANO Then BonificadorRaza = 2
        If Raza = DROW Then BonificadorRaza = 1
        If Raza = ENANO Then BonificadorRaza = 4
        If Raza = ELFO Then BonificadorRaza = 0
        If Raza = ORCO Then BonificadorRaza = 4
        If Raza = GNOMO Then BonificadorRaza = -1
End Select

End Function

Private Sub ResetAtributos()

Atributos = ATT_INICIALES
lbAtributos.Caption = Atributos

Dim i As Integer

For i = 1 To NUMATRIBUTOS
    lbAtt(i - 1).Caption = "6"
    CurrentUser.UserAtributos(i) = 6
Next i

End Sub
