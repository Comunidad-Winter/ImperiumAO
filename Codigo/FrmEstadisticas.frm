VERSION 5.00
Begin VB.Form frmEstadisticas 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Estadisticas del personaje"
   ClientHeight    =   6795
   ClientLeft      =   0
   ClientTop       =   -90
   ClientWidth     =   6510
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   453
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   434
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Fami 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   135
      Index           =   3
      Left            =   5610
      TabIndex        =   5
      Top             =   5880
      Width           =   645
   End
   Begin VB.Shape fHPShp 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Left            =   5610
      Top             =   5910
      Width           =   645
   End
   Begin VB.Label Fami 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   2
      Left            =   4335
      TabIndex        =   6
      Top             =   5430
      Width           =   1260
   End
   Begin VB.Shape fExpShp 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Left            =   4335
      Top             =   5460
      Width           =   1275
   End
   Begin VB.Label est 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Raza"
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
      Height          =   135
      Index           =   6
      Left            =   930
      TabIndex        =   52
      Top             =   3300
      Width           =   975
   End
   Begin VB.Label est 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Género"
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
      Height          =   135
      Index           =   5
      Left            =   930
      TabIndex        =   51
      Top             =   3090
      Width           =   975
   End
   Begin VB.Label Puntos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   195
      Left            =   5850
      TabIndex        =   50
      Top             =   3750
      Width           =   285
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   41
      Left            =   5970
      MouseIcon       =   "FrmEstadisticas.frx":0000
      Top             =   2100
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   40
      Left            =   5970
      MouseIcon       =   "FrmEstadisticas.frx":0152
      Top             =   2010
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   39
      Left            =   5970
      MouseIcon       =   "FrmEstadisticas.frx":02A4
      Top             =   1890
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   38
      Left            =   5970
      MouseIcon       =   "FrmEstadisticas.frx":03F6
      Top             =   1770
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   37
      Left            =   5970
      MouseIcon       =   "FrmEstadisticas.frx":0548
      Top             =   1650
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   36
      Left            =   5970
      MouseIcon       =   "FrmEstadisticas.frx":069A
      Top             =   1560
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   35
      Left            =   5970
      MouseIcon       =   "FrmEstadisticas.frx":07EC
      Top             =   1440
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   34
      Left            =   5970
      MouseIcon       =   "FrmEstadisticas.frx":093E
      Top             =   1350
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   33
      Left            =   5970
      MouseIcon       =   "FrmEstadisticas.frx":0A90
      Top             =   1200
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   32
      Left            =   5970
      MouseIcon       =   "FrmEstadisticas.frx":0BE2
      Top             =   1110
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   31
      Left            =   5970
      MouseIcon       =   "FrmEstadisticas.frx":0D34
      Top             =   990
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   30
      Left            =   5970
      MouseIcon       =   "FrmEstadisticas.frx":0E86
      Top             =   900
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   29
      Left            =   5970
      MouseIcon       =   "FrmEstadisticas.frx":0FD8
      Top             =   750
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   28
      Left            =   5970
      MouseIcon       =   "FrmEstadisticas.frx":112A
      Top             =   660
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   42
      Left            =   5970
      MouseIcon       =   "FrmEstadisticas.frx":127C
      Top             =   2250
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   43
      Left            =   5970
      MouseIcon       =   "FrmEstadisticas.frx":13CE
      Top             =   2370
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   44
      Left            =   5970
      MouseIcon       =   "FrmEstadisticas.frx":1520
      Top             =   2460
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   45
      Left            =   5970
      MouseIcon       =   "FrmEstadisticas.frx":1672
      Top             =   2580
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   46
      Left            =   5970
      MouseIcon       =   "FrmEstadisticas.frx":17C4
      Top             =   2700
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   47
      Left            =   5970
      MouseIcon       =   "FrmEstadisticas.frx":1916
      Top             =   2790
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   48
      Left            =   5970
      MouseIcon       =   "FrmEstadisticas.frx":1A68
      Top             =   2910
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   49
      Left            =   5970
      MouseIcon       =   "FrmEstadisticas.frx":1BBA
      Top             =   3000
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   50
      Left            =   5970
      MouseIcon       =   "FrmEstadisticas.frx":1D0C
      Top             =   3150
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   51
      Left            =   5970
      MouseIcon       =   "FrmEstadisticas.frx":1E5E
      Top             =   3240
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   52
      Left            =   5970
      MouseIcon       =   "FrmEstadisticas.frx":1FB0
      Top             =   3360
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   53
      Left            =   5970
      MouseIcon       =   "FrmEstadisticas.frx":2102
      Top             =   3450
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   26
      Left            =   4320
      MouseIcon       =   "FrmEstadisticas.frx":2254
      Top             =   3600
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   24
      Left            =   4320
      MouseIcon       =   "FrmEstadisticas.frx":23A6
      Top             =   3360
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   22
      Left            =   4320
      MouseIcon       =   "FrmEstadisticas.frx":24F8
      Top             =   3120
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   20
      Left            =   4320
      MouseIcon       =   "FrmEstadisticas.frx":264A
      Top             =   2910
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   18
      Left            =   4320
      MouseIcon       =   "FrmEstadisticas.frx":279C
      Top             =   2700
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   16
      Left            =   4320
      MouseIcon       =   "FrmEstadisticas.frx":28EE
      Top             =   2460
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   14
      Left            =   4320
      MouseIcon       =   "FrmEstadisticas.frx":2A40
      Top             =   2250
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   12
      Left            =   4320
      MouseIcon       =   "FrmEstadisticas.frx":2B92
      Top             =   2010
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   10
      Left            =   4320
      MouseIcon       =   "FrmEstadisticas.frx":2CE4
      Top             =   1800
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   8
      Left            =   4320
      MouseIcon       =   "FrmEstadisticas.frx":2E36
      Top             =   1560
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   6
      Left            =   4320
      MouseIcon       =   "FrmEstadisticas.frx":2F88
      Top             =   1320
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   4
      Left            =   4320
      MouseIcon       =   "FrmEstadisticas.frx":30DA
      Top             =   1110
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   2
      Left            =   4320
      MouseIcon       =   "FrmEstadisticas.frx":322C
      Top             =   870
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   0
      Left            =   4320
      MouseIcon       =   "FrmEstadisticas.frx":337E
      Top             =   660
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   1
      Left            =   4320
      MouseIcon       =   "FrmEstadisticas.frx":34D0
      Top             =   750
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   27
      Left            =   4320
      MouseIcon       =   "FrmEstadisticas.frx":3622
      Top             =   3690
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   25
      Left            =   4320
      MouseIcon       =   "FrmEstadisticas.frx":3774
      Top             =   3450
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   23
      Left            =   4320
      MouseIcon       =   "FrmEstadisticas.frx":38C6
      Top             =   3210
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   21
      Left            =   4320
      MouseIcon       =   "FrmEstadisticas.frx":3A18
      Top             =   3000
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   19
      Left            =   4320
      MouseIcon       =   "FrmEstadisticas.frx":3B6A
      Top             =   2790
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   17
      Left            =   4320
      MouseIcon       =   "FrmEstadisticas.frx":3CBC
      Top             =   2550
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   15
      Left            =   4320
      MouseIcon       =   "FrmEstadisticas.frx":3E0E
      Top             =   2340
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   13
      Left            =   4320
      MouseIcon       =   "FrmEstadisticas.frx":3F60
      Top             =   2130
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   11
      Left            =   4320
      MouseIcon       =   "FrmEstadisticas.frx":40B2
      Top             =   1890
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   9
      Left            =   4320
      MouseIcon       =   "FrmEstadisticas.frx":4204
      Top             =   1680
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   7
      Left            =   4320
      MouseIcon       =   "FrmEstadisticas.frx":4356
      Top             =   1440
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   5
      Left            =   4320
      MouseIcon       =   "FrmEstadisticas.frx":44A8
      Top             =   1200
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   3
      Left            =   4320
      MouseIcon       =   "FrmEstadisticas.frx":45FA
      Top             =   960
      Width           =   195
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
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
      Height          =   165
      Index           =   1
      Left            =   765
      TabIndex        =   49
      Top             =   4380
      Width           =   1020
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
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
      Height          =   165
      Index           =   2
      Left            =   765
      TabIndex        =   48
      Top             =   4605
      Width           =   1020
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
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
      Height          =   165
      Index           =   3
      Left            =   765
      TabIndex        =   47
      Top             =   5535
      Width           =   1020
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
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
      Height          =   165
      Index           =   4
      Left            =   765
      TabIndex        =   46
      Top             =   4830
      Width           =   1020
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
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
      Height          =   165
      Index           =   5
      Left            =   765
      TabIndex        =   45
      Top             =   5310
      Width           =   1020
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
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
      Height          =   165
      Index           =   6
      Left            =   765
      TabIndex        =   44
      Top             =   5760
      Width           =   1020
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
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
      Height          =   165
      Index           =   8
      Left            =   765
      TabIndex        =   43
      Top             =   5070
      Width           =   1020
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   195
      Index           =   1
      Left            =   4050
      TabIndex        =   42
      Top             =   900
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   4050
      TabIndex        =   41
      Top             =   690
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   4050
      TabIndex        =   40
      Top             =   1110
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   4050
      TabIndex        =   39
      Top             =   1350
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   4050
      TabIndex        =   38
      Top             =   1560
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   4050
      TabIndex        =   37
      Top             =   1800
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   4050
      TabIndex        =   36
      Top             =   2040
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   4050
      TabIndex        =   35
      Top             =   2250
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   4050
      TabIndex        =   34
      Top             =   2490
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   4050
      TabIndex        =   33
      Top             =   2700
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   4050
      TabIndex        =   32
      Top             =   2940
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   4050
      TabIndex        =   31
      Top             =   3150
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   4050
      TabIndex        =   30
      Top             =   3390
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   4050
      TabIndex        =   29
      Top             =   3600
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   5700
      TabIndex        =   28
      Top             =   690
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   15
      Left            =   5700
      TabIndex        =   27
      Top             =   930
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   16
      Left            =   5700
      TabIndex        =   26
      Top             =   1140
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   17
      Left            =   5700
      TabIndex        =   25
      Top             =   1350
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   20
      Left            =   5700
      TabIndex        =   24
      Top             =   2040
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   19
      Left            =   5700
      TabIndex        =   23
      Top             =   1800
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   18
      Left            =   5700
      TabIndex        =   22
      Top             =   1590
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   21
      Left            =   5700
      TabIndex        =   21
      Top             =   2250
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   22
      Left            =   5700
      TabIndex        =   20
      Top             =   2490
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   23
      Left            =   5700
      TabIndex        =   19
      Top             =   2700
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   24
      Left            =   5700
      TabIndex        =   18
      Top             =   2940
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   25
      Left            =   5700
      TabIndex        =   17
      Top             =   3150
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   26
      Left            =   5700
      TabIndex        =   16
      Top             =   3390
      Width           =   255
   End
   Begin VB.Image imgEstado 
      Height          =   315
      Left            =   615
      Top             =   6315
      Width           =   885
   End
   Begin VB.Image imgFami 
      Height          =   1680
      Left            =   4155
      Top             =   4980
      Width           =   2265
   End
   Begin VB.Image cmdGuardar 
      Height          =   480
      Left            =   3780
      Tag             =   "1"
      Top             =   3900
      Width           =   1050
   End
   Begin VB.Image cmdClose 
      Height          =   450
      Left            =   6120
      Tag             =   "1"
      Top             =   0
      Width           =   390
   End
   Begin VB.Label est 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Veces muerto"
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
      Height          =   195
      Index           =   4
      Left            =   2250
      TabIndex        =   15
      Top             =   5280
      Width           =   1665
   End
   Begin VB.Label Fami 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Acá van las habilidades especiales del familiar"
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
      Height          =   345
      Index           =   5
      Left            =   4230
      TabIndex        =   14
      Top             =   6270
      Width           =   2160
   End
   Begin VB.Label est 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Criminales"
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
      Height          =   195
      Index           =   3
      Left            =   2250
      TabIndex        =   13
      Top             =   6420
      Width           =   1665
   End
   Begin VB.Label est 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ciudadanos"
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
      Height          =   195
      Index           =   2
      Left            =   2250
      TabIndex        =   12
      Top             =   6030
      Width           =   1665
   End
   Begin VB.Label est 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Criaturas matadas"
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
      Height          =   195
      Index           =   1
      Left            =   2250
      TabIndex        =   11
      Top             =   5640
      Width           =   1665
   End
   Begin VB.Label est 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Clase"
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
      Height          =   135
      Index           =   0
      Left            =   930
      TabIndex        =   10
      Top             =   2850
      Width           =   975
   End
   Begin VB.Label Fami 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
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
      Height          =   165
      Index           =   4
      Left            =   4830
      TabIndex        =   9
      Top             =   5790
      Width           =   435
   End
   Begin VB.Label Fami 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   4170
      TabIndex        =   8
      Top             =   4950
      Width           =   2220
   End
   Begin VB.Label Fami 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   1
      Left            =   5910
      TabIndex        =   7
      Top             =   5370
      Width           =   225
   End
   Begin VB.Label Atri 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   165
      Index           =   5
      Left            =   1590
      TabIndex        =   4
      Top             =   1800
      Width           =   105
   End
   Begin VB.Label Atri 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   165
      Index           =   4
      Left            =   1590
      TabIndex        =   3
      Top             =   1530
      Width           =   105
   End
   Begin VB.Label Atri 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   165
      Index           =   3
      Left            =   1590
      TabIndex        =   2
      Top             =   1260
      Width           =   105
   End
   Begin VB.Label Atri 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   165
      Index           =   2
      Left            =   1590
      TabIndex        =   1
      Top             =   975
      Width           =   105
   End
   Begin VB.Label Atri 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   165
      Index           =   1
      Left            =   1590
      TabIndex        =   0
      Top             =   720
      Width           =   105
   End
End
Attribute VB_Name = "frmEstadisticas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmEstadisticas - ImperiumAO - v1.3.0
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

Private SkillsOrig(1 To NUMSKILLS) As Integer
Private LibresOrig As Integer
Private RealizoCambios As Boolean

Private Sub cmdClose_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call Sound.Sound_Play(SND_CLICK)
cmdClose.Picture = General_Load_Picture_From_Resource("cerrar-est-down.bmp")
End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If cmdClose.Tag = "0" Then
    cmdClose.Tag = "1"
    cmdClose.Picture = General_Load_Picture_From_Resource("cerrar-est-over.bmp")
End If

End Sub

Private Sub cmdClose_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim Resultado As VbMsgBoxResult

If RealizoCambios Then
    Resultado = MsgBox("Realizo cambios en sus skillpoints ¿desea guardar antes de salir?", vbQuestion + vbYesNoCancel, "Guardar cambios")
    If Resultado = vbYes Then Call cmdGuardar_MouseUp(Button, Shift, x, y)
End If

If Resultado <> vbCancel Then Unload Me

End Sub

Private Sub cmdGuardar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call Sound.Sound_Play(SND_CLICK)
cmdGuardar.Picture = General_Load_Picture_From_Resource("guardar-down.bmp")
End Sub

Private Sub cmdGuardar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If cmdGuardar.Tag = "0" Then
    cmdGuardar.Tag = "1"
    cmdGuardar.Picture = General_Load_Picture_From_Resource("guardar-over.bmp")
End If

End Sub

Private Sub cmdGuardar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim i As Integer, Cambio As Integer
Dim cad As String

Call Form_MouseMove(Button, Shift, x, y)

If RealizoCambios Then
    For i = 1 To NUMSKILLS
        Cambio = (CurrentUser.UserSkills(i) - SkillsOrig(i))
        If Cambio < 0 Then Exit Sub
        cad = cad & Cambio & ","
    Next
    SendData "SKSE" & cad
    RealizoCambios = False
End If

End Sub

Private Sub command1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Dim indice
Dim skillreal As Integer

If Index Mod 2 = 0 Then
    indice = Index \ 2
    If (CurrentUser.SkillPoints > 0) And (Val(Skill(indice).Caption) < 100) Then
        Skill(indice).Caption = Val(Skill(indice).Caption) + 1
        CurrentUser.UserSkills(SkillRealToIndex(indice + 1)) = Val(Skill(indice).Caption)
        CurrentUser.SkillPoints = CurrentUser.SkillPoints - 1
    End If
Else
    indice = Index \ 2
    If (Val(Skill(indice).Caption) > 0) And (SkillsOrig(SkillRealToIndex(indice + 1)) <= (Val(Skill(indice).Caption) - 1)) Then
        Skill(indice).Caption = Val(Skill(indice).Caption) - 1
        CurrentUser.UserSkills(SkillRealToIndex(indice + 1)) = Val(Skill(indice).Caption)
        CurrentUser.SkillPoints = CurrentUser.SkillPoints + 1
    End If
End If

Puntos.Caption = CurrentUser.SkillPoints
RealizoCambios = (CurrentUser.SkillPoints <> LibresOrig)
Skill(indice).ForeColor = IIf(CurrentUser.UserSkills(SkillRealToIndex(indice + 1)) = SkillsOrig(SkillRealToIndex(indice + 1)), vbWhite, vbRed)

End Sub

Private Sub Form_Load()
Me.Picture = General_Load_Picture_From_Resource("stats.bmp")
End Sub

Public Sub Iniciar_Labels()

Dim i As Integer

'Iniciamos los labels con los valores de los atributos y los skills
For i = 1 To NUMATRIBUTOS
    Atri(i).Caption = CurrentUser.UserAtributos(i)
Next

For i = 1 To NUMSKILLS
    Skill(RealSkillToIndex(i) - 1).Caption = CurrentUser.UserSkills(i)
    'Skill(i - 1).FontBold = (CurrentUser.UserSkills(i) < 100)
    SkillsOrig(i) = CurrentUser.UserSkills(i)
Next

Puntos.Caption = CurrentUser.SkillPoints
LibresOrig = CurrentUser.SkillPoints

Label4(1).Caption = CurrentUser.UserReputacion.AsesinoRep
Label4(2).Caption = CurrentUser.UserReputacion.BandidoRep
Label4(3).Caption = CurrentUser.UserReputacion.BurguesRep
Label4(4).Caption = CurrentUser.UserReputacion.LadronesRep
Label4(5).Caption = CurrentUser.UserReputacion.NobleRep
Label4(6).Caption = CurrentUser.UserReputacion.PlebeRep
Label4(8).Caption = CurrentUser.UserReputacion.Culpabilidad

If CurrentUser.UserReputacion.Promedio < 0 Then
    imgEstado.Picture = General_Load_Picture_From_Resource("criminal.bmp")
ElseIf CurrentUser.UserReputacion.Promedio > 0 Then
    imgEstado.Picture = Nothing
Else
    imgEstado.Picture = General_Load_Picture_From_Resource("neutral.bmp")
End If

'Ponemos las estadisticas del familiar en pantalla
If CurrentUser.UserPet.TieneFamiliar <> 0 Then
    imgFami.Picture = Nothing
    Fami(0).Visible = True
    Fami(1).Visible = True
    Fami(2).Visible = True
    Fami(3).Visible = True
    Fami(4).Visible = True
    Fami(5).Visible = True
    fHPShp.Visible = True
    fExpShp.Visible = True
    
    Fami(0).Caption = CurrentUser.UserPet.nombre
    Fami(1).Caption = CurrentUser.UserPet.ELV
    
    Call PetExpPerc
    
    If CurrentUser.PetPercExp <> 0 Then
        fExpShp.Width = (((CurrentUser.UserPet.EXP / 100) / (CurrentUser.UserPet.ELU / 100)) * 85)
    Else
        fExpShp.Width = 0
    End If
    
    Fami(2).Caption = CurrentUser.PetPercExp & "%"
    
    If CurrentUser.UserPet.MinHP = 0 Then
        Fami(3).Caption = "Muerto"
        Fami(3).ForeColor = vbWhite
        fHPShp.Width = 0
    Else
        fExpShp.Width = (((CurrentUser.UserPet.MinHP / 100) / (CurrentUser.UserPet.MaxHP / 100)) * 43)
        Fami(3).Caption = CurrentUser.UserPet.MinHP & "/" & CurrentUser.UserPet.MaxHP
    End If
    Fami(4).Caption = CurrentUser.UserPet.MinHIT & "/" & CurrentUser.UserPet.MaxHIT
    Fami(5).Caption = IIf(CurrentUser.UserPet.Abilidad = "", "Ninguna", CurrentUser.UserPet.Abilidad)
Else
    imgFami.Picture = General_Load_Picture_From_Resource("fmnodisp.bmp")
    Fami(0).Visible = False
    Fami(1).Visible = False
    Fami(2).Visible = False
    Fami(3).Visible = False
    Fami(4).Visible = False
    Fami(5).Visible = False
    fHPShp.Visible = False
    fExpShp.Visible = False
End If

'Stats generales
est(0).Caption = CharClaseValueToString(CurrentUser.UserStats.Clase)
est(1).Caption = CurrentUser.UserStats.NPCsMatados
est(2).Caption = CurrentUser.UserStats.CiudasMatados
est(3).Caption = CurrentUser.UserStats.CrimisMatados
est(4).Caption = CurrentUser.UserStats.TimesKilled
est(5).Caption = IIf(CurrentUser.UserStats.Genero = Masculino, "Masculino", "Femenino")
est(6).Caption = RazaToString(CurrentUser.UserStats.Raza)

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If (Button = vbLeftButton) Then
    Call Auto_Drag(Me.hwnd)
Else
    Unload Me
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If cmdClose.Tag = "1" Then
    cmdClose.Tag = "0"
    cmdClose.Picture = Nothing
End If

If cmdGuardar.Tag = "1" Then
    cmdGuardar.Tag = "0"
    cmdGuardar.Picture = Nothing
End If

End Sub

Private Function RazaToString(ByVal Raza As Byte) As String

Select Case Raza
    Case HUMANO
        RazaToString = "Humano"
    Case ENANO
        RazaToString = "Enano"
    Case ELFO
        RazaToString = "Elfo"
    Case DROW
        RazaToString = "Elfo Drow"
    Case GNOMO
        RazaToString = "Gnomo"
    Case ORCO
        RazaToString = "Orco"
End Select

End Function

Private Function CharClaseValueToString(ByVal Clase As Byte) As String

Select Case Clase

Case CLERIGO
    CharClaseValueToString = "Clérigo"
Case MAGO
    CharClaseValueToString = "Mago"
Case GUERRERO
    CharClaseValueToString = "Guerrero"
Case ASESINO
    CharClaseValueToString = "Asesino"
Case LADRON
    CharClaseValueToString = "Ladrón"
Case BARDO
    CharClaseValueToString = "Bardo"
Case DRUIDA
    CharClaseValueToString = "Druida"
Case CAZARECOMPENSAS
    CharClaseValueToString = "Cazarecompensas"
Case PALADIN
    CharClaseValueToString = "Paladín"
Case CAZADOR
    CharClaseValueToString = "Cazador"
Case PESCADOR
    CharClaseValueToString = "Pescador"
Case HERRERO
    CharClaseValueToString = "Herrero"
Case LEÑADOR
    CharClaseValueToString = "Leñador"
Case MINERO
    CharClaseValueToString = "Minero"
Case CARPINTERO
    CharClaseValueToString = "Carpintero"
Case SASTRE
    CharClaseValueToString = "Sastre"
Case PIRATA
    CharClaseValueToString = "Pirata"
Case NIGROMANTE
    CharClaseValueToString = "Nigromante"
Case gm
    CharClaseValueToString = "Game Master"
Case Else
    CharClaseValueToString = ""

End Select

End Function
