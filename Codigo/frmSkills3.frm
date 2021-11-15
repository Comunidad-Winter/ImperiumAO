VERSION 5.00
Begin VB.Form frmSkills3 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6390
   ClientLeft      =   675
   ClientTop       =   45
   ClientWidth     =   10785
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   26
      Left            =   5400
      TabIndex        =   54
      Top             =   4800
      Width           =   2205
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   27
      Left            =   8640
      TabIndex        =   53
      Top             =   4800
      Width           =   765
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   53
      Left            =   8160
      Top             =   4800
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   52
      Left            =   9480
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   25
      Left            =   5400
      TabIndex        =   52
      Top             =   4440
      Width           =   2205
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   26
      Left            =   8640
      TabIndex        =   51
      Top             =   4440
      Width           =   765
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   51
      Left            =   8160
      Top             =   4440
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   50
      Left            =   9480
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   24
      Left            =   5400
      TabIndex        =   50
      Top             =   4080
      Width           =   2205
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   25
      Left            =   8640
      TabIndex        =   49
      Top             =   4080
      Width           =   765
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   49
      Left            =   8160
      Top             =   4080
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   48
      Left            =   9480
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   23
      Left            =   5400
      TabIndex        =   48
      Top             =   3720
      Width           =   2205
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   24
      Left            =   8640
      TabIndex        =   47
      Top             =   3720
      Width           =   765
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   47
      Left            =   8160
      Top             =   3720
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   46
      Left            =   9480
      Top             =   3720
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   45
      Left            =   8160
      Top             =   3360
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   44
      Left            =   9480
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   23
      Left            =   8640
      TabIndex        =   46
      Top             =   3360
      Width           =   765
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   22
      Left            =   5400
      TabIndex        =   45
      Top             =   3360
      Width           =   2205
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   43
      Left            =   8160
      Top             =   3000
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   42
      Left            =   9480
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   22
      Left            =   8640
      TabIndex        =   44
      Top             =   3000
      Width           =   765
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   21
      Left            =   5400
      TabIndex        =   43
      Top             =   3000
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   0
      Left            =   600
      TabIndex        =   42
      Top             =   480
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   1
      Left            =   600
      TabIndex        =   41
      Top             =   840
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   2
      Left            =   600
      TabIndex        =   40
      Top             =   1200
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   3
      Left            =   600
      TabIndex        =   39
      Top             =   1560
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   4
      Left            =   600
      TabIndex        =   38
      Top             =   1920
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   5
      Left            =   600
      TabIndex        =   37
      Top             =   2280
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   6
      Left            =   600
      TabIndex        =   36
      Top             =   2640
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   7
      Left            =   600
      TabIndex        =   35
      Top             =   3000
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   8
      Left            =   600
      TabIndex        =   34
      Top             =   3360
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   9
      Left            =   600
      TabIndex        =   33
      Top             =   3720
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   10
      Left            =   600
      TabIndex        =   32
      Top             =   4080
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   11
      Left            =   600
      TabIndex        =   31
      Top             =   4440
      Width           =   2205
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   30
      Top             =   480
      Width           =   765
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   3840
      TabIndex        =   29
      Top             =   840
      Width           =   765
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   3840
      TabIndex        =   28
      Top             =   1200
      Width           =   765
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   4
      Left            =   3840
      TabIndex        =   27
      Top             =   1560
      Width           =   765
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   5
      Left            =   3840
      TabIndex        =   26
      Top             =   1920
      Width           =   765
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   6
      Left            =   3840
      TabIndex        =   25
      Top             =   2280
      Width           =   765
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   7
      Left            =   3840
      TabIndex        =   24
      Top             =   2640
      Width           =   765
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   8
      Left            =   3840
      TabIndex        =   23
      Top             =   3000
      Width           =   765
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   9
      Left            =   3840
      TabIndex        =   22
      Top             =   3360
      Width           =   765
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   10
      Left            =   3840
      TabIndex        =   21
      Top             =   3720
      Width           =   765
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   11
      Left            =   3840
      TabIndex        =   20
      Top             =   4080
      Width           =   765
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   12
      Left            =   3840
      TabIndex        =   19
      Top             =   4440
      Width           =   765
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   0
      Left            =   4680
      Top             =   480
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   2
      Left            =   4680
      Top             =   840
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   3
      Left            =   3360
      Top             =   840
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   4
      Left            =   4680
      Top             =   1200
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   5
      Left            =   3360
      Top             =   1200
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   6
      Left            =   4680
      Top             =   1560
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   7
      Left            =   3360
      Top             =   1560
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   8
      Left            =   4680
      Top             =   1920
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   9
      Left            =   3360
      Top             =   1920
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   10
      Left            =   4680
      Top             =   2280
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   11
      Left            =   3360
      Top             =   2280
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   12
      Left            =   4680
      Top             =   2640
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   13
      Left            =   3360
      Top             =   2640
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   14
      Left            =   4680
      Top             =   3000
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   15
      Left            =   3360
      Top             =   3000
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   16
      Left            =   4680
      Top             =   3360
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   17
      Left            =   3360
      Top             =   3360
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   18
      Left            =   4680
      Top             =   3720
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   19
      Left            =   3360
      Top             =   3720
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   20
      Left            =   4680
      Top             =   4080
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   21
      Left            =   3360
      Top             =   4080
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   22
      Left            =   4680
      Top             =   4440
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   23
      Left            =   3360
      Top             =   4440
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   24
      Left            =   4680
      Top             =   4800
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   25
      Left            =   3360
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   13
      Left            =   3840
      TabIndex        =   18
      Top             =   4800
      Width           =   765
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   12
      Left            =   600
      TabIndex        =   17
      Top             =   4800
      Width           =   2205
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   26
      Left            =   4680
      Top             =   5160
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   27
      Left            =   3360
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   14
      Left            =   3840
      TabIndex        =   16
      Top             =   5160
      Width           =   765
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   13
      Left            =   600
      TabIndex        =   15
      Top             =   5160
      Width           =   2205
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   28
      Left            =   9480
      Top             =   480
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   29
      Left            =   8160
      Top             =   480
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   15
      Left            =   8640
      TabIndex        =   14
      Top             =   480
      Width           =   765
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   14
      Left            =   5400
      TabIndex        =   13
      Top             =   480
      Width           =   2205
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   30
      Left            =   9480
      Top             =   840
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   31
      Left            =   8160
      Top             =   840
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   16
      Left            =   8640
      TabIndex        =   12
      Top             =   840
      Width           =   765
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   15
      Left            =   5400
      TabIndex        =   11
      Top             =   840
      Width           =   2205
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   32
      Left            =   9480
      Top             =   1200
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   33
      Left            =   8160
      Top             =   1200
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   17
      Left            =   8640
      TabIndex        =   10
      Top             =   1200
      Width           =   765
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   16
      Left            =   5400
      TabIndex        =   9
      Top             =   1200
      Width           =   2205
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   34
      Left            =   9480
      Top             =   1560
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   35
      Left            =   8160
      Top             =   1560
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   18
      Left            =   8640
      TabIndex        =   8
      Top             =   1560
      Width           =   765
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   17
      Left            =   5400
      TabIndex        =   7
      Top             =   1560
      Width           =   2205
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   1
      Left            =   3360
      Top             =   480
      Width           =   345
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   18
      Left            =   5400
      TabIndex        =   6
      Top             =   1920
      Width           =   2205
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   19
      Left            =   8640
      TabIndex        =   5
      Top             =   1920
      Width           =   765
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   36
      Left            =   9480
      Top             =   1920
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   37
      Left            =   8160
      Top             =   1920
      Width           =   345
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   19
      Left            =   5400
      TabIndex        =   4
      Top             =   2280
      Width           =   2205
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   20
      Left            =   8640
      TabIndex        =   3
      Top             =   2280
      Width           =   765
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   38
      Left            =   9480
      Top             =   2280
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   39
      Left            =   8160
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "skill1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   20
      Left            =   5400
      TabIndex        =   2
      Top             =   2640
      Width           =   2205
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   21
      Left            =   8640
      TabIndex        =   1
      Top             =   2640
      Width           =   765
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   40
      Left            =   9480
      Top             =   2640
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   41
      Left            =   8160
      Top             =   2640
      Width           =   345
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   4920
      Top             =   5760
      Width           =   855
   End
   Begin VB.Label puntos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Puntos:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   630
   End
End
Attribute VB_Name = "frmSkills3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Private cargado As Boolean

Private Sub Form_Paint()
If Not cargado Then
    cargado = True
End If
End Sub
Private Sub Command1_Click(Index As Integer)

Call IAO_SE.PlaySound(SND_CLICK)

Dim indice
If Index Mod 2 = 0 Then
    If Alocados > 0 Then
        indice = Index \ 2 + 1
        If indice > NUMSKILLS Then indice = NUMSKILLS
        If UserSkills(indice) < MAXSKILLPOINTS Then
            Text1(indice).Caption = Val(Text1(indice).Caption) + 1
            flags(indice) = flags(indice) + 1
            Alocados = Alocados - 1
        End If
            
    End If
Else
    If Alocados < SkillPoints Then
        
        indice = Index \ 2 + 1
        If Val(Text1(indice).Caption) > 0 And flags(indice) > 0 Then
            Text1(indice).Caption = Val(Text1(indice).Caption) - 1
            flags(indice) = flags(indice) - 1
            Alocados = Alocados + 1
        End If
    End If
End If

Puntos.Caption = "Puntos:" & Alocados
End Sub

Private Sub Form_Load()

Image1.Picture = LoadPicture(App.Path & "\Interface\Botonok.jpg")

'Nombres de los skills
Dim l
Dim i As Integer
i = 1
For Each l In Label2
    l.Caption = SkillsNames(i)
    l.AutoSize = True
    i = i + 1
Next
i = 0

'Flags para saber que skills se modificaron
ReDim flags(1 To NUMSKILLS)


'Cargamos el jpg correspondiente
For i = 0 To NUMSKILLS * 2 - 1
    If i Mod 2 = 0 Then
        Command1(i).Picture = LoadPicture(App.Path & "\Interface\BotonMás.jpg")
    Else
        Command1(i).Picture = LoadPicture(App.Path & "\Interface\BotonMenos.jpg")
    End If
Next

'Alocados = SkillPoints
End Sub

Private Sub Image1_Click()

Dim i As Integer
Dim cad As String
For i = 1 To NUMSKILLS
    cad = cad & flags(i) & ","
    If flags(i) > 0 Then
        UserSkills(i) = UserSkills(i) + flags(i) 'Barrin: actualizar skills dinámicamente
    End If
Next
SendData "SKSE" & cad
If Alocados = 0 Then frmMain.Label1.Visible = False
SkillPoints = Alocados
Unload Me
End Sub
