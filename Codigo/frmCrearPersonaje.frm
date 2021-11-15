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
      MaxLength       =   20
      TabIndex        =   26
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
      ItemData        =   "frmCrearPersonaje.frx":0ECA
      Left            =   8520
      List            =   "frmCrearPersonaje.frx":0ECC
      Style           =   2  'Dropdown List
      TabIndex        =   25
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
      ItemData        =   "frmCrearPersonaje.frx":0ECE
      Left            =   870
      List            =   "frmCrearPersonaje.frx":0ED0
      Style           =   2  'Dropdown List
      TabIndex        =   23
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
      ItemData        =   "frmCrearPersonaje.frx":0ED2
      Left            =   870
      List            =   "frmCrearPersonaje.frx":0EDF
      Style           =   2  'Dropdown List
      TabIndex        =   22
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
      ItemData        =   "frmCrearPersonaje.frx":0EFA
      Left            =   870
      List            =   "frmCrearPersonaje.frx":0F13
      Style           =   2  'Dropdown List
      TabIndex        =   21
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
      ItemData        =   "frmCrearPersonaje.frx":0F46
      Left            =   8550
      List            =   "frmCrearPersonaje.frx":0F5F
      Style           =   2  'Dropdown List
      TabIndex        =   2
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
      MaxLength       =   25
      TabIndex        =   0
      Top             =   1050
      Width           =   5865
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
      MouseIcon       =   "frmCrearPersonaje.frx":0F9A
      Top             =   6930
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   52
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":10EC
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
      MouseIcon       =   "frmCrearPersonaje.frx":123E
      Top             =   6540
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   50
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":1390
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
      MouseIcon       =   "frmCrearPersonaje.frx":14E2
      Top             =   6180
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   48
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":1634
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
      MouseIcon       =   "frmCrearPersonaje.frx":1786
      Top             =   5790
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   46
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":18D8
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
      MouseIcon       =   "frmCrearPersonaje.frx":1A2A
      Top             =   5430
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   44
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":1B7C
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
      MouseIcon       =   "frmCrearPersonaje.frx":1CCE
      Top             =   5040
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   42
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":1E20
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
      TabIndex        =   24
      Top             =   7260
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   3
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":1F72
      Top             =   2790
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   5
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":20C4
      Top             =   3180
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   7
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":2216
      Top             =   3540
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   9
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":2368
      Top             =   3930
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   11
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":24BA
      Top             =   4290
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   13
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":260C
      Top             =   4680
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   15
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":275E
      Top             =   5040
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   17
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":28B0
      Top             =   5430
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   19
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":2A02
      Top             =   5790
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   21
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":2B54
      Top             =   6180
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   23
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":2CA6
      Top             =   6540
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   25
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":2DF8
      Top             =   6930
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   27
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":2F4A
      Top             =   7290
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   1
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":309C
      Top             =   2430
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   0
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":31EE
      Top             =   2280
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   2
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":3340
      Top             =   2670
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   4
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":3492
      Top             =   3060
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   6
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":35E4
      Top             =   3420
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   8
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":3736
      Top             =   3780
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   10
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":3888
      Top             =   4170
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   12
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":39DA
      Top             =   4560
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   14
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":3B2C
      Top             =   4920
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   16
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":3C7E
      Top             =   5280
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   18
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":3DD0
      Top             =   5670
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   20
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":3F22
      Top             =   6060
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   22
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":4074
      Top             =   6420
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   24
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":41C6
      Top             =   6810
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   26
      Left            =   5580
      MouseIcon       =   "frmCrearPersonaje.frx":4318
      Top             =   7170
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   28
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":446A
      Top             =   2280
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   29
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":45BC
      Top             =   2430
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   30
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":470E
      Top             =   2670
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   31
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":4860
      Top             =   2820
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   32
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":49B2
      Top             =   3060
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   33
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":4B04
      Top             =   3180
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   34
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":4C56
      Top             =   3420
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   35
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":4DA8
      Top             =   3570
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   36
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":4EFA
      Top             =   3780
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   37
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":504C
      Top             =   3930
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   38
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":519E
      Top             =   4170
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   39
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":52F0
      Top             =   4320
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   40
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":5442
      Top             =   4560
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   41
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":5594
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
      TabIndex        =   20
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
      TabIndex        =   19
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
      TabIndex        =   18
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
      TabIndex        =   17
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
      TabIndex        =   16
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
      TabIndex        =   15
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
      TabIndex        =   14
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
      TabIndex        =   13
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
      TabIndex        =   12
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
      TabIndex        =   11
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
      TabIndex        =   10
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   7
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
      TabIndex        =   6
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
      TabIndex        =   5
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
      TabIndex        =   4
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
      TabIndex        =   3
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
      TabIndex        =   1
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
'frmCrearPersonaje - ImperiumAO - v1.3.0
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

Private Function CheckData() As Boolean

If CurrentUser.UserName = "" Then
    lblInfo.Caption = "Seleccione el nombre del personaje."
    Exit Function
End If

If CurrentUser.UserRaza = 0 Then
    lblInfo.Caption = "Seleccione la raza del personaje."
    Exit Function
End If

If CurrentUser.UserSexo = 0 Then
    lblInfo.Caption = "Seleccione el género del personaje."
    Exit Function
End If

If CurrentUser.UserClase = 0 Then
    lblInfo.Caption = "Seleccione la clase del personaje."
    Exit Function
End If

If CurrentUser.UserHogar = 0 Then
    lblInfo.Caption = "Seleccione el hogar del personaje."
    Exit Function
End If

If SkillPoints > 0 Then
    lblInfo.Caption = "Asigne los skillpoints del personaje."
    Exit Function
End If

If Atributos > 0 Then
    lblInfo.Caption = "Asigne los atributos del personaje."
    Exit Function
End If

If frmCrearPersonaje.lstFamiliar.Visible = True Then

    If CurrentUser.UserPet.Tipo = "" Then
        lblInfo.Caption = "Seleccione su familiar o mascota."
        Exit Function
    ElseIf CurrentUser.UserPet.nombre = "" Then
        lblInfo.Caption = "Asigne un nombre a su familiar o mascota."
        Exit Function
    ElseIf Len(CurrentUser.UserPet.nombre) > 30 Then
        lblInfo.Caption = ("El nombre de tu familiar o mascota debe tener menos de 30 letras.")
        Exit Function
    End If

End If

Dim i As Integer
For i = 1 To NUMATRIBUTOS
    CurrentUser.UserAtributos(i) = Val(lbAtt(i - 1).Caption)
    If CurrentUser.UserAtributos(i) = 0 Then
        lblInfo.Caption = "Los atributos del personaje son invalidos."
        Exit Function
    End If
Next i

CheckData = True

End Function

Private Sub imgAccion_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Call Sound.Sound_Play(SND_CLICK)
Call imgAccionRestaurar
    
Select Case Index
    Case 0
        
        Dim i As Integer
        
        For i = 0 To 26
            CurrentUser.UserSkills(SkillRealToIndex(i + 1)) = Val(Skill(i).Caption)
        Next i
        
        CurrentUser.UserName = txtNombre.Text
        
        If Right(CurrentUser.UserName, 1) = " " Then
            CurrentUser.UserName = RTrim(CurrentUser.UserName)
            lblInfo.Caption = "¡Nombre invalido, se han removido los espacios al final del nombre!"
        End If
        
        CurrentUser.UserRaza = lstRaza.ListIndex
        CurrentUser.UserSexo = lstGenero.ListIndex
        CurrentUser.UserClase = lstProfesion.ListIndex
        CurrentUser.UserPet.Tipo = lstFamiliar.List(lstFamiliar.ListIndex)
        CurrentUser.UserPet.nombre = frmCrearPersonaje.txtFamiliar.Text
        CurrentUser.UserHogar = lstHogar.ListIndex
        Atributos = Val(lbAtributos.Caption)
        
        frmPasswd.lblStatus.Caption = ""
        If CheckData() Then frmPasswd.Show vbModeless, frmCrearPersonaje
        
    Case 1
        If Musica <> CONST_DESHABILITADA Then
            Sound.NextMusic = MUS_VolverInicio
            Sound.Fading = 200
        End If
        
        frmConnect.MousePointer = 1
        frmPasswd.MousePointer = 1
        
        If frmMain.mainWinsock.State Then
            frmMain.mainWinsock.Close
        Else
            Call ResetCurrentUser
        End If
        
        frmConnect.Show
        Unload Me

End Select

End Sub

Private Sub imgAccion_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Select Case Index
    Case 0 'Crear
        imgAccion(0).Picture = General_Load_Picture_From_Resource("creardown.bmp")
        imgAccion(0).Tag = "0"
    Case 1 'Volver
        imgAccion(1).Picture = General_Load_Picture_From_Resource("volverdown.bmp")
        imgAccion(1).Tag = "0"
End Select

End Sub

Private Sub imgAccion_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Select Case Index
    Case 0 'Crear personaje
        If imgAccion(0).Tag = "1" Then
            imgAccion(0).Picture = General_Load_Picture_From_Resource("crearover.bmp")
            imgAccion(0).Tag = "0"
            Call Sound.Sound_Play(SND_OVER)
        End If
        
    Case 1 'Volver
        If imgAccion(1).Tag = "1" Then
            imgAccion(1).Picture = General_Load_Picture_From_Resource("volverover.bmp")
            imgAccion(1).Tag = "0"
            Call Sound.Sound_Play(SND_OVER)
        End If
End Select

Call imgAccionRestaurar(Index)

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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim i As Integer

For i = 0 To 1
    If imgAccion(i).Tag = "0" Then
        imgAccion(i).Picture = Nothing
        imgAccion(i).Tag = "1"
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

Me.Picture = General_Load_Picture_From_Resource("cp-interface.bmp")

Dim i As Integer
lstProfesion.Clear
lstProfesion.AddItem ""
For i = LBound(ListaClases) To UBound(ListaClases)
    lstProfesion.AddItem ListaClases(i)
Next i

lstProfesion.ListIndex = 0
lstGenero.ListIndex = 0
lstRaza.ListIndex = 0
lstHogar.ListIndex = 0

Image1.Picture = General_Load_Picture_From_Resource(LCase$(lstProfesion.Text) & ".bmp")
Call ResetAtributos

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

Private Sub lbAtributos_Click()
Call Sound.Sound_Play(SND_DICE)
Call ResetAtributos
End Sub

Private Sub lblInfo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call imgAccionRestaurar
End Sub

Private Sub lstFamiliar_Click()

If lstFamiliar.ListIndex > 0 Then
    lblFamiInfo.Caption = ListaFamiliares(lstFamiliar.ListIndex).Desc
    picFamiliar.Picture = General_Load_Picture_From_Resource(ListaFamiliares(lstFamiliar.ListIndex).Imagen)
Else
    lblFamiInfo.Caption = "Selecciona tu familiar o mascota para saber más de él"
    picFamiliar.Picture = Nothing
End If

End Sub

Private Sub lstHogar_Click()
    If lstHogar.Text = "Rinkel" Or lstHogar.Text = "Lindos" Then
        lblInfo.Caption = "El seleccionar como ciudad natal Rinkel o Lindos te hará neutral. Esto implica varias cosas a tener en cuenta, te recomendamos pensarlo cuidadosamente."
    End If
End Sub

Private Sub lstProfesion_Click()
On Error Resume Next
Image1.Picture = General_Load_Picture_From_Resource(LCase$(lstProfesion.Text) & ".bmp")

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
    imgNoDisp.Picture = General_Load_Picture_From_Resource("mascotanodisp.bmp")
    picFamiliar.Visible = False
    lblFamiInfo.Visible = False
End If

End Sub

Private Sub txtNombre_Change()
txtNombre.Text = LTrim(txtNombre.Text)
End Sub

Private Sub txtNombre_GotFocus()
lblInfo.Caption = "Sea cuidadoso al seleccionar el nombre de su personaje, ImperiumAO es un juego de rol, un mundo magico y fantastico, si selecciona un nombre obsceno o con connotación politica los administradores borrarán su personaje y no habrá ninguna posibilidad de recuperarlo."
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
End Sub

Private Sub CambioFamiliar(ByVal NumFamiliares As Integer)

If NumFamiliares = 5 Then

    ReDim ListaFamiliares(1 To NumFamiliares) As tListaFamiliares
    ListaFamiliares(1).Name = "Elemental De Fuego"
    ListaFamiliares(1).Desc = "Hecho de puro fuego, lanzará tormentas sobre tus contrincantes."
    ListaFamiliares(1).Imagen = "elefuego.bmp"
    
    ListaFamiliares(2).Name = "Elemental De Agua"
    ListaFamiliares(2).Desc = "Con su cuerpo acuoso paralizará a tus enemigos."
    ListaFamiliares(2).Imagen = "eleagua.bmp"
    
    ListaFamiliares(3).Name = "Elemental De Tierra"
    ListaFamiliares(3).Desc = "Sus fuertes brazos inmovilizarán cualquier criatura viviente."
    ListaFamiliares(3).Imagen = "eletierra.bmp"
    
    ListaFamiliares(4).Name = "Ely"
    ListaFamiliares(4).Desc = "Te protegerá constantemente con sus conjuros defensivos."
    ListaFamiliares(4).Imagen = "ely.bmp"
    
    ListaFamiliares(5).Name = "Fuego Fatuo"
    ListaFamiliares(5).Desc = "Débil pero con gran poder mágico, siempre estará a tu lado."
    ListaFamiliares(5).Imagen = "fatuo.bmp"
    
Else

    ReDim ListaFamiliares(1 To NumFamiliares) As tListaFamiliares
    ListaFamiliares(1).Name = "Tigre"
    ListaFamiliares(1).Desc = "Poseen grandes y filosas garras para atacar a tus oponentes."
    ListaFamiliares(1).Imagen = "tigre.bmp"
    
    ListaFamiliares(2).Name = "Lobo"
    ListaFamiliares(2).Desc = "Astutos y arrogantes, su mordedura causa estragos en sus víctimas."
    ListaFamiliares(2).Imagen = "lobo.bmp"
    
    ListaFamiliares(3).Name = "Oso Pardo"
    ListaFamiliares(3).Desc = "Se caracterizan por ser territoriales y muy resistentes."
    ListaFamiliares(3).Imagen = "oso.bmp"
    
    ListaFamiliares(4).Name = "Ent"
    ListaFamiliares(4).Desc = "¡Esta robusta criatura te defenderá cual muro de piedra!"
    ListaFamiliares(4).Imagen = "ent.bmp"

End If

Dim i As Integer
lstFamiliar.Clear
lstFamiliar.AddItem ""
For i = 1 To UBound(ListaFamiliares)
    lstFamiliar.AddItem ListaFamiliares(i).Name
Next i

lstFamiliar.ListIndex = 0

End Sub

Private Sub txtfamiliar_GotFocus()
lblInfo.Caption = "Mucho cuidado al colocarle nombre a su familiar, no puede ponerle el mismo o parecido nombre de su personaje, recuerde que es su companía. En caso de que el familiar o mascota tenga nombre inapropiado, podrá ser retirado."
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

If lstRaza.List(lstRaza.ListIndex) = "" Then Exit Sub

Dim i As Integer

For i = 1 To NUMATRIBUTOS
    lbBonificador(i - 1).Caption = IIf(BonificadorRaza(i, lstRaza.List(lstRaza.ListIndex)) > 0, "+" & BonificadorRaza(i, lstRaza.List(lstRaza.ListIndex)), BonificadorRaza(i, lstRaza.List(lstRaza.ListIndex)))
    If Val(lbBonificador(i - 1)) = 0 Then
        lbBonificador(i - 1).Visible = False
    Else
        lbBonificador(i - 1).Visible = True
    End If
Next i

End Sub

Public Function BonificadorRaza(Atributo As Integer, Raza As String) As Integer

Dim TmpStr As String
TmpStr = UCase$(Raza)

Select Case Atributo
'Ryghar: Nuevo balance
    Case Fuerza
        If TmpStr = "HUMANO" Then BonificadorRaza = 1
        If TmpStr = "ELFO DROW" Then BonificadorRaza = 2
        If TmpStr = "ENANO" Then BonificadorRaza = 3
        If TmpStr = "ELFO" Then BonificadorRaza = -2
        If TmpStr = "ORCO" Then BonificadorRaza = 5
        If TmpStr = "GNOMO" Then BonificadorRaza = -5
        Exit Function
    Case Agilidad
        If TmpStr = "HUMANO" Then BonificadorRaza = 1
        If TmpStr = "ELFO DROW" Then BonificadorRaza = -2
        If TmpStr = "ENANO" Then BonificadorRaza = -2
        If TmpStr = "ELFO" Then BonificadorRaza = 2
        If TmpStr = "ORCO" Then BonificadorRaza = -1
        If TmpStr = "GNOMO" Then BonificadorRaza = 3
        Exit Function
    Case Inteligencia
        If TmpStr = "HUMANO" Then BonificadorRaza = 1
        If TmpStr = "ELFO DROW" Then BonificadorRaza = 3
        If TmpStr = "ENANO" Then BonificadorRaza = -5
        If TmpStr = "ELFO" Then BonificadorRaza = 2
        If TmpStr = "ORCO" Then BonificadorRaza = -5
        If TmpStr = "GNOMO" Then BonificadorRaza = 4
        Exit Function
    Case Carisma
        If TmpStr = "HUMANO" Then BonificadorRaza = 0
        If TmpStr = "ELFO DROW" Then BonificadorRaza = -1
        If TmpStr = "ENANO" Then BonificadorRaza = -1
        If TmpStr = "ELFO" Then BonificadorRaza = 2
        If TmpStr = "ORCO" Then BonificadorRaza = -4
        If TmpStr = "GNOMO" Then BonificadorRaza = 0
        Exit Function
    Case Constitucion
        If TmpStr = "HUMANO" Then BonificadorRaza = 2
        If TmpStr = "ELFO DROW" Then BonificadorRaza = 0
        If TmpStr = "ENANO" Then BonificadorRaza = 4
        If TmpStr = "ELFO" Then BonificadorRaza = 1
        If TmpStr = "ORCO" Then BonificadorRaza = 3
        If TmpStr = "GNOMO" Then BonificadorRaza = -1
        Exit Function
'/Ryghar: Nuevo balance
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
