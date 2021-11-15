VERSION 5.00
Begin VB.Form frmClanes 
   BorderStyle     =   0  'None
   Caption         =   "Clanes"
   ClientHeight    =   6390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   426
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   610
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frameListaDeClanes 
      Caption         =   "Lista de Clanes"
      Height          =   1575
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   7335
      Begin VB.ListBox lstListaDeClanes 
         Height          =   1230
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   6975
      End
   End
   Begin VB.Timer tmrShowInfo 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   120
      Top             =   1710
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   0
      ScaleHeight     =   285
      ScaleWidth      =   9150
      TabIndex        =   1
      Top             =   6105
      Width           =   9150
      Begin VB.Label lblShowInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   45
      End
   End
   Begin VB.PictureBox picBotones 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   6105
      Left            =   7455
      ScaleHeight     =   407
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   113
      TabIndex        =   0
      Top             =   0
      Width           =   1695
      Begin VB.CommandButton cmdOpciones 
         Caption         =   "Votar nuevo lider"
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   129
         Top             =   4200
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdOpciones 
         Caption         =   "Políticas"
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   10
         Top             =   3720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdOpciones 
         Caption         =   "Administrar GM"
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   9
         Top             =   3240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdOpciones 
         Caption         =   "Administrar Clan"
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   8
         Top             =   2760
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdOpciones 
         Caption         =   "Fundar Clan"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Top             =   2280
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdOpciones 
         Caption         =   "Información"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdOpciones 
         Caption         =   "Salir del Clan"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdOpciones 
         Caption         =   "Solicitar Ingreso"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdOpciones 
         Caption         =   "Cerrar"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame frameSolicitarIngreso 
      Caption         =   "Solicitar Ingreso "
      Height          =   4335
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   7335
      Begin VB.CommandButton cmdSolicitarIngresoCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   4320
         TabIndex        =   135
         Top             =   3360
         Width           =   1455
      End
      Begin VB.CommandButton cmdFrameSolicitarIngresoMandarSolicitud 
         Caption         =   "Solicitar ingreso"
         Height          =   375
         Left            =   1080
         TabIndex        =   17
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox txtMensajeAlLider 
         Appearance      =   0  'Flat
         Height          =   975
         Left            =   480
         TabIndex        =   14
         Top             =   1920
         Width           =   5775
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmClanes.frx":0000
         Height          =   735
         Left            =   360
         TabIndex        =   16
         Top             =   600
         Width           =   6615
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Mensaje al lider:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   480
         TabIndex        =   15
         Top             =   1560
         Width           =   1140
      End
   End
   Begin VB.Frame frameAdministracionRecursos 
      Caption         =   "Administración de Clan - Recursos"
      Height          =   4335
      Left            =   120
      TabIndex        =   94
      Top             =   1680
      Visible         =   0   'False
      Width           =   7335
      Begin VB.CommandButton CmdAdministrarRecursosAtras 
         Caption         =   "Atras"
         Height          =   375
         Left            =   3120
         TabIndex        =   136
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton CmdRecursosSacar 
         Caption         =   "Sacar"
         Height          =   255
         Index           =   4
         Left            =   4800
         TabIndex        =   128
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton CmdRecursosSacar 
         Caption         =   "Sacar"
         Height          =   255
         Index           =   3
         Left            =   4800
         TabIndex        =   127
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton CmdRecursosSacar 
         Caption         =   "Sacar"
         Height          =   255
         Index           =   2
         Left            =   4800
         TabIndex        =   126
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton CmdRecursosSacar 
         Caption         =   "Sacar"
         Height          =   255
         Index           =   1
         Left            =   4800
         TabIndex        =   125
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton CmdRecursosSacar 
         Caption         =   "Sacar"
         Height          =   255
         Index           =   0
         Left            =   4800
         TabIndex        =   124
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton CmdRecursosAgregar 
         Caption         =   "Agregar"
         Height          =   255
         Index           =   4
         Left            =   3840
         TabIndex        =   123
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton CmdRecursosAgregar 
         Caption         =   "Agregar"
         Height          =   255
         Index           =   3
         Left            =   3840
         TabIndex        =   122
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton CmdRecursosAgregar 
         Caption         =   "Agregar"
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   121
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton CmdRecursosAgregar 
         Caption         =   "Agregar"
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   120
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton CmdRecursosAgregar 
         Caption         =   "Agregar"
         Height          =   255
         Index           =   0
         Left            =   3840
         TabIndex        =   119
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtRecursosCantidad 
         Height          =   285
         Index           =   4
         Left            =   5760
         TabIndex        =   118
         Text            =   "0"
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtRecursosCantidad 
         Height          =   285
         Index           =   3
         Left            =   5760
         TabIndex        =   117
         Text            =   "0"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtRecursosCantidad 
         Height          =   285
         Index           =   2
         Left            =   5760
         TabIndex        =   116
         Text            =   "0"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtRecursosCantidad 
         Height          =   285
         Index           =   1
         Left            =   5760
         TabIndex        =   115
         Text            =   "0"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtRecursosCantidad 
         Height          =   285
         Index           =   0
         Left            =   5760
         TabIndex        =   114
         Text            =   "0"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtCantidadRecurso 
         Height          =   285
         Index           =   4
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   113
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox txtCantidadRecurso 
         Height          =   285
         Index           =   3
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   112
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox txtCantidadRecurso 
         Height          =   285
         Index           =   2
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   111
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtCantidadRecurso 
         Height          =   285
         Index           =   1
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   110
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtCantidadRecurso 
         Height          =   285
         Index           =   0
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   109
         Top             =   480
         Width           =   1935
      End
      Begin VB.Line Line3 
         X1              =   3720
         X2              =   3720
         Y1              =   2400
         Y2              =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Lingotes de hierro"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   108
         Top             =   1920
         Width           =   1260
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Lingotes de plata"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   107
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Lingotes de oro"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   106
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Madera"
         Height          =   195
         Left            =   360
         TabIndex        =   105
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Oro"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   104
         Top             =   480
         Width           =   255
      End
      Begin VB.Label lblSolicitudesData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   26
         Left            =   3240
         TabIndex        =   103
         Top             =   1680
         Width           =   45
      End
      Begin VB.Label lblSolicitudesData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   25
         Left            =   3240
         TabIndex        =   102
         Top             =   1920
         Width           =   45
      End
      Begin VB.Label lblSolicitudesData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   24
         Left            =   3240
         TabIndex        =   101
         Top             =   2160
         Width           =   45
      End
      Begin VB.Label lblSolicitudesData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   23
         Left            =   3240
         TabIndex        =   100
         Top             =   2400
         Width           =   45
      End
      Begin VB.Label lblSolicitudesData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   22
         Left            =   3240
         TabIndex        =   99
         Top             =   2640
         Width           =   45
      End
      Begin VB.Label lblSolicitudesData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   21
         Left            =   3240
         TabIndex        =   98
         Top             =   2880
         Width           =   45
      End
      Begin VB.Label lblSolicitudesData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   20
         Left            =   3240
         TabIndex        =   97
         Top             =   3120
         Width           =   45
      End
      Begin VB.Label lblSolicitudesData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   19
         Left            =   3240
         TabIndex        =   96
         Top             =   3360
         Width           =   45
      End
      Begin VB.Label lblSolicitudesData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   18
         Left            =   3240
         TabIndex        =   95
         Top             =   3600
         Width           =   45
      End
   End
   Begin VB.Frame frameADM 
      Caption         =   "Administración de Clan - Miembros"
      Height          =   4335
      Left            =   120
      TabIndex        =   54
      Top             =   1680
      Visible         =   0   'False
      Width           =   7335
      Begin VB.CommandButton CmdSolicitudesRechazar 
         Caption         =   "Rechazar Miembro"
         Enabled         =   0   'False
         Height          =   495
         Left            =   1560
         TabIndex        =   68
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CommandButton CmdSolicitudesAceptar 
         Caption         =   "Aceptar Miembro"
         Enabled         =   0   'False
         Height          =   495
         Left            =   240
         TabIndex        =   67
         Top             =   3600
         Width           =   1095
      End
      Begin VB.ListBox lstSolucitudesDeIngreso 
         Height          =   1815
         Left            =   120
         TabIndex        =   58
         Top             =   1680
         Width           =   2895
      End
      Begin VB.CommandButton cmdListaPjsEchar 
         Caption         =   "Echar miembro"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5400
         TabIndex        =   57
         Top             =   480
         Width           =   1695
      End
      Begin VB.ListBox lstListaDeMiembros 
         Height          =   645
         Left            =   240
         TabIndex        =   56
         Top             =   360
         Width           =   4935
      End
      Begin VB.CommandButton cmdADMCancel 
         Caption         =   "Atras"
         Height          =   495
         Left            =   5880
         TabIndex        =   55
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblSolicitudesData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   7
         Left            =   3240
         TabIndex        =   66
         Top             =   3360
         Width           =   45
      End
      Begin VB.Label lblSolicitudesData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   6
         Left            =   3240
         TabIndex        =   65
         Top             =   3120
         Width           =   45
      End
      Begin VB.Label lblSolicitudesData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   5
         Left            =   3240
         TabIndex        =   64
         Top             =   2880
         Width           =   45
      End
      Begin VB.Label lblSolicitudesData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   4
         Left            =   3240
         TabIndex        =   63
         Top             =   2640
         Width           =   45
      End
      Begin VB.Label lblSolicitudesData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   3
         Left            =   3240
         TabIndex        =   62
         Top             =   2400
         Width           =   45
      End
      Begin VB.Label lblSolicitudesData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   2
         Left            =   3240
         TabIndex        =   61
         Top             =   2160
         Width           =   45
      End
      Begin VB.Label lblSolicitudesData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   1
         Left            =   3240
         TabIndex        =   60
         Top             =   1920
         Width           =   45
      End
      Begin VB.Label lblSolicitudesData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   0
         Left            =   3240
         TabIndex        =   59
         Top             =   1680
         Width           =   45
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   5760
         Y1              =   1440
         Y2              =   1440
      End
   End
   Begin VB.Frame frameAdministrarClanLider 
      Caption         =   "Administración de Clan - Lider"
      Height          =   4335
      Left            =   120
      TabIndex        =   50
      Top             =   1680
      Visible         =   0   'False
      Width           =   7335
      Begin VB.CommandButton cmdAdministrarLiderTratados 
         Caption         =   "Adnibistrar Tratados"
         Height          =   495
         Left            =   600
         TabIndex        =   167
         Top             =   3240
         Width           =   4455
      End
      Begin VB.CommandButton cmdAdministrarLiderRecursos 
         Caption         =   "Administrar Recursos del Clan"
         Height          =   495
         Left            =   600
         TabIndex        =   93
         ToolTipText     =   "Para Agregar o sacar recursos al clan"
         Top             =   2280
         Width           =   4455
      End
      Begin VB.CommandButton cmdAdministrarLiderCancelar 
         Caption         =   "Atras"
         Height          =   495
         Left            =   5520
         TabIndex        =   53
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CommandButton cmdAdministrarLiderClan 
         Caption         =   "Administrar Clan"
         Height          =   495
         Left            =   600
         TabIndex        =   52
         ToolTipText     =   "Para cambiar la descripción del clan y para recivir o no ingresos"
         Top             =   1320
         Width           =   4455
      End
      Begin VB.CommandButton cmdAdministrarLiderMiembros 
         Caption         =   "Administrar miembros del Clan"
         Height          =   495
         Left            =   600
         TabIndex        =   51
         ToolTipText     =   "Para dar de alta y de baja a pjs"
         Top             =   360
         Width           =   4455
      End
   End
   Begin VB.Frame FrameInformacionDelClan 
      Caption         =   "Información del Clan "
      Height          =   4335
      Left            =   120
      TabIndex        =   18
      Top             =   1680
      Visible         =   0   'False
      Width           =   7335
      Begin VB.CommandButton CmAceptar 
         Caption         =   "Atras"
         Height          =   255
         Left            =   6120
         TabIndex        =   41
         Top             =   3840
         Width           =   1095
      End
      Begin VB.ListBox lstTitulos 
         Height          =   645
         Left            =   240
         TabIndex        =   19
         Top             =   3600
         Width           =   5775
      End
      Begin VB.Label lblParseInfo 
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   40
         Top             =   240
         Width           =   4920
      End
      Begin VB.Label lblParseInfo 
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   1
         Left            =   960
         TabIndex        =   39
         Top             =   480
         Width           =   4920
      End
      Begin VB.Label lblParseInfo 
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   38
         Top             =   720
         Width           =   4920
      End
      Begin VB.Label lblParseInfo 
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   4
         Left            =   960
         TabIndex        =   37
         Top             =   1200
         Width           =   4920
      End
      Begin VB.Label lblParseInfo 
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   5
         Left            =   1680
         TabIndex        =   36
         Top             =   1440
         Width           =   4920
      End
      Begin VB.Label lblParseInfo 
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   6
         Left            =   1440
         TabIndex        =   35
         Top             =   1680
         Width           =   4920
      End
      Begin VB.Label lblParseInfo 
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   7
         Left            =   1080
         TabIndex        =   34
         Top             =   1920
         Width           =   4920
      End
      Begin VB.Label lblParseInfo 
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   8
         Left            =   960
         TabIndex        =   33
         Top             =   2160
         Width           =   4920
      End
      Begin VB.Label lblParseInfo 
         BackStyle       =   0  'Transparent
         Height          =   675
         Index           =   9
         Left            =   1080
         TabIndex        =   32
         Top             =   2400
         Width           =   4920
      End
      Begin VB.Label lblParseInfo 
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   3
         Left            =   480
         TabIndex        =   31
         Top             =   960
         Width           =   1680
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fundado: "
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fundador: "
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Width           =   765
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lider: "
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   435
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Oro: "
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   345
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GuildPoints: "
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   885
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número de miembros: "
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   25
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Miembros on line: "
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   24
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lider online?"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   23
         Top             =   1920
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alineacion: "
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   22
         Top             =   2160
         Width           =   825
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción: "
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   21
         Top             =   2400
         Width           =   930
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Títulos"
         Height          =   195
         Index           =   10
         Left            =   240
         TabIndex        =   20
         Top             =   3360
         Width           =   495
      End
   End
   Begin VB.Frame frameAdministracionClan 
      Caption         =   "Administración de Clan"
      Height          =   4335
      Left            =   120
      TabIndex        =   69
      Top             =   1680
      Visible         =   0   'False
      Width           =   7335
      Begin VB.TextBox txtAdminClanDescripcion 
         Height          =   1455
         Left            =   240
         MaxLength       =   300
         TabIndex        =   83
         Top             =   1560
         Width           =   6615
      End
      Begin VB.CheckBox chkAdminClanAceptarSolicitudes 
         Caption         =   "Aceptar solicitudes de ingreso al clan"
         Height          =   255
         Left            =   360
         TabIndex        =   81
         Top             =   600
         Width           =   6015
      End
      Begin VB.CommandButton cmdAdminClanAceptar 
         Caption         =   "Aceptar Cambios"
         Enabled         =   0   'False
         Height          =   495
         Left            =   720
         TabIndex        =   71
         Top             =   3600
         Width           =   1935
      End
      Begin VB.CommandButton cmdAdminClanCancelar 
         Caption         =   "Atras"
         Height          =   495
         Left            =   4560
         TabIndex        =   70
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Descripción del clan"
         Height          =   255
         Left            =   240
         TabIndex        =   82
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblSolicitudesData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   17
         Left            =   3240
         TabIndex        =   80
         Top             =   1680
         Width           =   45
      End
      Begin VB.Label lblSolicitudesData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   16
         Left            =   3240
         TabIndex        =   79
         Top             =   1920
         Width           =   45
      End
      Begin VB.Label lblSolicitudesData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   15
         Left            =   3240
         TabIndex        =   78
         Top             =   2160
         Width           =   45
      End
      Begin VB.Label lblSolicitudesData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   14
         Left            =   3240
         TabIndex        =   77
         Top             =   2400
         Width           =   45
      End
      Begin VB.Label lblSolicitudesData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   13
         Left            =   3240
         TabIndex        =   76
         Top             =   2640
         Width           =   45
      End
      Begin VB.Label lblSolicitudesData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   12
         Left            =   3240
         TabIndex        =   75
         Top             =   2880
         Width           =   45
      End
      Begin VB.Label lblSolicitudesData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   11
         Left            =   3240
         TabIndex        =   74
         Top             =   3120
         Width           =   45
      End
      Begin VB.Label lblSolicitudesData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   10
         Left            =   3240
         TabIndex        =   73
         Top             =   3360
         Width           =   45
      End
      Begin VB.Label lblSolicitudesData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   9
         Left            =   3240
         TabIndex        =   72
         Top             =   3600
         Width           =   45
      End
   End
   Begin VB.Frame frameAdminGM 
      Caption         =   "Administración de Clan - GM"
      Height          =   4335
      Left            =   120
      TabIndex        =   84
      Top             =   1680
      Visible         =   0   'False
      Width           =   7335
      Begin VB.CommandButton CmdGmUpdateData 
         Caption         =   "Mandar al server"
         Enabled         =   0   'False
         Height          =   495
         Left            =   5400
         TabIndex        =   92
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CheckBox chkGmAbierto 
         Caption         =   "Clan abierto"
         Height          =   255
         Left            =   5400
         TabIndex        =   91
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox txtGMEditDesc 
         Height          =   855
         Left            =   360
         TabIndex        =   90
         Top             =   3000
         Width           =   4935
      End
      Begin VB.CommandButton CmdAdminGmHabilitarClan 
         Caption         =   "Habilitar Clan"
         Height          =   495
         Left            =   600
         TabIndex        =   89
         Top             =   480
         Width           =   4455
      End
      Begin VB.CommandButton CmdAdminGmEditarClan 
         Caption         =   "Editar Clan"
         Height          =   495
         Left            =   600
         TabIndex        =   88
         Top             =   1680
         Width           =   4455
      End
      Begin VB.CommandButton CmdAdminGmBorrar 
         Caption         =   "Borrar Clan"
         Height          =   495
         Left            =   600
         TabIndex        =   87
         Top             =   2280
         Width           =   855
      End
      Begin VB.CommandButton CmdAdminGmDesHabilitarClan 
         Caption         =   "DesHabilitar Clan"
         Height          =   495
         Left            =   600
         TabIndex        =   86
         Top             =   1080
         Width           =   4455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cancelar"
         Height          =   495
         Left            =   5520
         TabIndex        =   85
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   6720
         Y1              =   2880
         Y2              =   2880
      End
   End
   Begin VB.Frame frameFundarClan 
      Caption         =   "Fundación de clan"
      Height          =   4335
      Left            =   120
      TabIndex        =   42
      Top             =   1680
      Visible         =   0   'False
      Width           =   7335
      Begin VB.CommandButton CmdFundarClanesCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   5040
         TabIndex        =   49
         Top             =   3480
         Width           =   1575
      End
      Begin VB.CommandButton CmdFundarClanesAceptar 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   720
         TabIndex        =   48
         Top             =   3480
         Width           =   1575
      End
      Begin VB.TextBox txtNombreNuevoClan 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   240
         MaxLength       =   25
         TabIndex        =   44
         Top             =   1920
         Width           =   6855
      End
      Begin VB.TextBox txtDescripcionNuevoClan 
         Appearance      =   0  'Flat
         Height          =   645
         Left            =   240
         MaxLength       =   200
         TabIndex        =   43
         Top             =   2520
         Width           =   6855
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre del clan a crear"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   360
         TabIndex        =   47
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmClanes.frx":00DD
         Height          =   1215
         Left            =   360
         TabIndex        =   46
         Top             =   360
         Width           =   6855
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción del clan a crear"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   360
         TabIndex        =   45
         Top             =   2280
         Width           =   1980
      End
   End
   Begin VB.Frame framePoliticasMain 
      Caption         =   "Relaciones entre clanes"
      Height          =   4335
      Left            =   120
      TabIndex        =   138
      Top             =   1680
      Visible         =   0   'False
      Width           =   7335
      Begin VB.CommandButton CmdrelacionesEntreClanesAtras 
         Caption         =   "Atras"
         Height          =   315
         Left            =   2760
         TabIndex        =   176
         Top             =   3960
         Width           =   1695
      End
      Begin VB.ListBox lstListaDeClanesEnGuerra 
         Height          =   1425
         Left            =   3720
         TabIndex        =   151
         Top             =   2400
         Width           =   3255
      End
      Begin VB.ListBox lstListaDeClanesEnPaz 
         Height          =   1425
         Left            =   120
         TabIndex        =   149
         Top             =   2400
         Width           =   3255
      End
      Begin VB.CommandButton cmdPoliticasMainVerGlobal 
         Caption         =   "Ver relación"
         Height          =   255
         Left            =   5280
         TabIndex        =   147
         Top             =   1680
         Width           =   1695
      End
      Begin VB.ComboBox cmbRelacionesEntreClanesGlobal 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   146
         Top             =   1680
         Width           =   3135
      End
      Begin VB.CommandButton cmdPoliticasMainVerRelacion 
         Caption         =   "Ver relación"
         Height          =   255
         Left            =   5280
         TabIndex        =   141
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox cmbRelacionesEntreClanes 
         Height          =   315
         Index           =   1
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   140
         Top             =   240
         Width           =   3135
      End
      Begin VB.ComboBox cmbRelacionesEntreClanes 
         Height          =   315
         Index           =   0
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   139
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Clanes en GUERRA"
         Height          =   195
         Left            =   3840
         TabIndex        =   150
         Top             =   2160
         Width           =   1440
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Clanes en PAZ"
         Height          =   195
         Left            =   240
         TabIndex        =   148
         Top             =   2160
         Width           =   1065
      End
      Begin VB.Label Label12 
         Caption         =   "Ver información global"
         Height          =   255
         Left            =   240
         TabIndex        =   145
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label lblEstadoEntreClanes 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   144
         Top             =   1200
         Width           =   45
      End
      Begin VB.Label lblEstadoEntreClanes 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   143
         Top             =   960
         Width           =   45
      End
      Begin VB.Label lblEstadoEntreClanes 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   142
         Top             =   720
         Width           =   45
      End
      Begin VB.Line Line4 
         X1              =   120
         X2              =   7080
         Y1              =   1560
         Y2              =   1560
      End
   End
   Begin VB.Frame frameSolicitudesPoliticas 
      Caption         =   "Solicitudes de tratados"
      Height          =   4335
      Left            =   120
      TabIndex        =   152
      Top             =   1680
      Visible         =   0   'False
      Width           =   7335
      Begin VB.CommandButton CmdSolicitudTratadosAtras 
         Caption         =   "Atras"
         Height          =   255
         Left            =   5880
         TabIndex        =   175
         Top             =   3960
         Width           =   1215
      End
      Begin VB.CommandButton cmdrechazarTratado 
         Caption         =   "RechazarTratado"
         Height          =   375
         Left            =   1800
         TabIndex        =   170
         Top             =   3840
         Width           =   1455
      End
      Begin VB.CommandButton cmdAceptarTratado 
         Caption         =   "Aceptar Tratado"
         Height          =   375
         Left            =   120
         TabIndex        =   169
         Top             =   3840
         Width           =   1455
      End
      Begin VB.ListBox lstTratadosPendientes 
         Enabled         =   0   'False
         Height          =   1815
         Left            =   120
         TabIndex        =   168
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox txtSolicitarPoliticasDuración 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   4680
         MaxLength       =   4
         TabIndex        =   166
         Text            =   "-1"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   495
         Left            =   120
         TabIndex        =   160
         Top             =   960
         Width           =   2175
         Begin VB.OptionButton optPagoSolicitud 
            Caption         =   "Solicitar oro por el trato"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   162
            Top             =   240
            Width           =   1935
         End
         Begin VB.OptionButton optPagoSolicitud 
            Caption         =   "Ofrecer oro por el trato"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   161
            Top             =   0
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Left            =   1920
            TabIndex        =   164
            Top             =   240
            Width           =   60
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   195
            Left            =   1920
            TabIndex        =   163
            Top             =   0
            Width           =   60
         End
      End
      Begin VB.TextBox txtSolicitarPoliticasPago 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   158
         Text            =   "0"
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmdEnviarSolicitud 
         Caption         =   "Enviar solicitud"
         Height          =   375
         Left            =   5280
         TabIndex        =   157
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton optSolicitarPlitica 
         Caption         =   "Solicitar Guerra"
         Height          =   195
         Index           =   2
         Left            =   3240
         TabIndex        =   156
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton optSolicitarPlitica 
         Caption         =   "Solicitar Paz"
         Height          =   195
         Index           =   1
         Left            =   1920
         TabIndex        =   155
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optSolicitarPlitica 
         Caption         =   "Solicitar Neutralidad"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   154
         Top             =   600
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.ComboBox cmbSolicitarPoliticas 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   153
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label lblsSolicitudTratados 
         AutoSize        =   -1  'True
         Caption         =   "Duración del tratado"
         Height          =   195
         Index           =   3
         Left            =   3360
         TabIndex        =   174
         Top             =   2880
         Width           =   1440
      End
      Begin VB.Label lblsSolicitudTratados 
         AutoSize        =   -1  'True
         Caption         =   "Oferta "
         Height          =   195
         Index           =   2
         Left            =   3360
         TabIndex        =   173
         Top             =   2520
         Width           =   480
      End
      Begin VB.Label lblsSolicitudTratados 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de tratado"
         Height          =   195
         Index           =   1
         Left            =   3360
         TabIndex        =   172
         Top             =   2160
         Width           =   1080
      End
      Begin VB.Label lblsSolicitudTratados 
         AutoSize        =   -1  'True
         Caption         =   "Tratado ofertado por "
         Height          =   195
         Index           =   0
         Left            =   3360
         TabIndex        =   171
         Top             =   1800
         Width           =   1500
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Duración (días) del tratado (-1 indeterminado)"
         Height          =   195
         Left            =   3960
         TabIndex        =   165
         Top             =   960
         Width           =   3195
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "* pago diario"
         Height          =   195
         Left            =   2280
         TabIndex        =   159
         Top             =   1320
         Width           =   885
      End
      Begin VB.Line Line5 
         X1              =   360
         X2              =   6720
         Y1              =   1560
         Y2              =   1560
      End
   End
   Begin VB.Frame frameVotar 
      Caption         =   "Votar Lider de Clan"
      Height          =   4335
      Left            =   120
      TabIndex        =   130
      Top             =   1680
      Visible         =   0   'False
      Width           =   7335
      Begin VB.CommandButton CmdVotarCancelar 
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   2880
         TabIndex        =   134
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton CmdVotarNuevoLider 
         Caption         =   "Votar"
         Height          =   375
         Left            =   5160
         TabIndex        =   133
         Top             =   2160
         Width           =   1575
      End
      Begin VB.ComboBox cmbMiembroParaSerLider 
         Height          =   315
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   131
         Top             =   2160
         Width           =   4335
      End
      Begin VB.Label Label11 
         Caption         =   $"frmClanes.frx":02BE
         Height          =   855
         Left            =   360
         TabIndex        =   137
         Top             =   360
         Width           =   6735
      End
      Begin VB.Label Label10 
         Caption         =   "Candidatos para ser lider"
         Height          =   255
         Left            =   480
         TabIndex        =   132
         Top             =   1800
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmClanes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum eFrames
    Ninguno
    AdministracionMiembros
    AdministracionClan
    AdministracionViaGM
    AdministracionViaLider
    FundarClan
    InformacionDelClan
    SolicitarIngreso
    AdministracionRecursos
    Votacion
    PoliticasMain
    PoliticasSolicitudes
End Enum

Private lista() As String
Private Habilitado() As Boolean
Private NumClanes As Long
Private IniFlags As Long
Private pedidos() As String
Private Miembros() As String
Private SolicitudesPendientesPoliticas() As String

Public Sub ParseADM(sData As String)
Dim parametros() As String
Dim valores() As String
Dim i As Long
ShowInfo "Llegaron datos.. procesando"
parametros = Split(sData, CV2_CHAR_SEP_PARAMETROS)
If UBound(parametros) < 0 Then
    ShowInfo "Información corrupta"
    Exit Sub
End If
valores = Split(parametros(0), CV2_CHAR_SEP_VALORES)
lstListaDeMiembros.Clear
ReDim Miembros(LBound(valores) To UBound(valores))
For i = LBound(valores) To UBound(valores)
    Miembros(i) = valores(i)
    lstListaDeMiembros.AddItem Miembros(i)
Next i
lstSolucitudesDeIngreso.Clear
If i > 1 Then
    cmdListaPjsEchar.Enabled = True
End If

If parametros(1) = "" Then
    lstSolucitudesDeIngreso.Enabled = False
    CmdSolicitudesAceptar.Enabled = False
    CmdSolicitudesRechazar.Enabled = False
Else
    ReDim pedidos(LBound(parametros) To UBound(parametros))
    For i = 1 To UBound(parametros)
        valores = Split(parametros(i), CV2_CHAR_SEP_VALORES)
        lstSolucitudesDeIngreso.AddItem valores(4)
        pedidos(i) = parametros(i)
    Next i
    CmdSolicitudesAceptar.Enabled = True
    CmdSolicitudesRechazar.Enabled = True
End If
ShowFrame AdministracionMiembros
End Sub

Private Sub CmAceptar_Click()
EnableButtons
FrameInformacionDelClan.Caption = "Información del Clan "
lstListaDeClanes.Enabled = True
ShowFrame Ninguno
End Sub

Private Sub cmdAceptarTratado_Click()
SendData "CV2LAT" & lstTratadosPendientes.ItemData(lstTratadosPendientes.ListIndex)
lstTratadosPendientes.RemoveItem lstTratadosPendientes.ListIndex
If lstTratadosPendientes.ListCount = 0 Then
    cmdAceptarTratado.Enabled = False
    cmdrechazarTratado.Enabled = False
    lstTratadosPendientes.Enabled = False
End If
End Sub

Private Sub cmdADMCancel_Click()
ShowFrame AdministracionViaLider
DisableButtons
End Sub

Private Sub cmdAdminClanAceptar_Click()
SendData "CV2ACC" & IIf(chkAdminClanAceptarSolicitudes.value = 1, "+", "-") & Me.txtAdminClanDescripcion
cmdAdminClanCancelar_Click
End Sub

Private Sub cmdAdminClanCancelar_Click()
ShowFrame AdministracionViaLider
txtAdminClanDescripcion = ""
chkAdminClanAceptarSolicitudes.value = 0
End Sub

Private Sub CmdAdminGmBorrar_Click()
Dim NombreClan As String
If lstListaDeClanes.ListIndex < 0 Then
    ShowInfo "selecciona el clan peterete"
Else
    NombreClan = lista(lstListaDeClanes.ListIndex)
    If MsgBox("Vas a borrar el clan " & NombreClan & ". Si no sabes que haces" & vbCrLf & "lo que haces, apreta no. Si lo borras al pedo" & vbCrLf & "Sinuhe te va a cagar a patadas en el orto", vbCritical + vbYesNo, "Confirmación") = vbYes Then _
        SendData "CV2BOR" & NombreClan
End If
Unload Me
End Sub

Private Sub CmdAdminGmDesHabilitarClan_Click()
Dim NombreClan As String
If lstListaDeClanes.ListIndex < 0 Then
    ShowInfo "selecciona el clan peterete"
Else
    If Not Habilitado(lstListaDeClanes.ListIndex) Then
        ShowInfo "Clan no habilitado marmota"
    Else
        SendData "CV2DHC" & lstListaDeClanes.List(lstListaDeClanes.ListIndex)
        
    End If
End If

End Sub

Private Sub CmdAdminGmEditarClan_Click()
If lstListaDeClanes.ListIndex < 0 Then
    ShowInfo "pff.. todavía no sabes que tenes que marcar un clan?!"
Else
    ShowInfo "Esperando datos..."
    SendData "CV2EDI" & lista(lstListaDeClanes.ListIndex)
End If
End Sub

Public Sub ParseGmEdit(sData As String)
ShowInfo "Llegaron los datos =)"
CmdGmUpdateData.Enabled = True
If mid(sData, 1, 1) = "+" Then
    chkGmAbierto.value = 1
Else
    chkGmAbierto.value = 0
End If
If Len(sData) > 2 Then
    txtGMEditDesc.Text = mid(sData, 3)
End If
End Sub

Private Sub CmdAdminGmHabilitarClan_Click()
Dim NombreClan As String
If lstListaDeClanes.ListIndex < 0 Then
    ShowInfo "selecciona el clan peterete"
Else
    If Habilitado(lstListaDeClanes.ListIndex) Then
        ShowInfo "Clan habilitado marmota"
    Else
        SendData "CV2HCL" & mid$(lstListaDeClanes.List(lstListaDeClanes.ListIndex), 2)
    End If
End If
End Sub

Private Sub cmdAdministrarLiderCancelar_Click()
EnableButtons
ShowFrame Ninguno
End Sub
Private Sub BotonesAdmin(habilitados As Boolean)
cmdAdministrarLiderClan.Enabled = habilitados
cmdAdministrarLiderRecursos.Enabled = habilitados
cmdAdministrarLiderMiembros.Enabled = habilitados
cmdAdministrarLiderTratados.Enabled = habilitados
'cmdAdministrarLiderPoliticasPaz.Enabled = habilitados
'cmdAdministrarLiderPoliticasNeutralidad.Enabled = habilitados
'cmdAdministrarLiderPoliticasGuerra.Enabled = habilitados
End Sub
Private Sub cmdAdministrarLiderClan_Click()
BotonesAdmin False
ShowInfo "Esperando datos... espere por favor"
SendData "CV2ALC"
End Sub

Private Sub cmdAdministrarLiderMiembros_Click()
BotonesAdmin False
ShowInfo "Esperando datos... espere por favor", True
SendData "CV2ALM"
End Sub

Private Sub cmdAdministrarLiderRecursos_Click()
BotonesAdmin False
SendData "CV2LAR"
ShowInfo "Esperando datos..."
End Sub

Private Sub CmdAdministrarRecursosAtras_Click()
ShowFrame AdministracionViaLider
DisableButtons
End Sub

Private Sub cmdEnviarSolicitud_Click()
Dim i As Long
Dim TipoTratado As Long
Dim Pago As Long
Dim Duracion As Long
If cmbSolicitarPoliticas.ListIndex < 0 Then
    ShowInfo "Primero selecciona un clan a ofrecer un tratado"
    cmbSolicitarPoliticas.SetFocus
    Exit Sub
End If
If cmbSolicitarPoliticas.ItemData(cmbSolicitarPoliticas.ListIndex) = -1 Then
    ShowInfo "El clan no esta habilitado"
    cmbSolicitarPoliticas.SetFocus
    Exit Sub
End If
For i = 0 To 2
    If optSolicitarPlitica(i).value Then
        Exit For
    End If
Next i
If i > 2 Then
    ShowInfo "Error en el formulario. Seleccione que tipo de tratado ofrece"
    Exit Sub
End If
TipoTratado = i
If optPagoSolicitud(0).value Then
    'yo pago
    Pago = txtSolicitarPoliticasPago
    If Pago < 0 Then
        ShowInfo "No podes poner números negativos como oferta monetaria"
        txtSolicitarPoliticasPago.SetFocus
        Exit Sub
    End If
Else
    'que me paguen!!
    Pago = txtSolicitarPoliticasPago * -1
    If Pago > 0 Then
        ShowInfo "No podes poner números negativos como oferta monetaria"
        txtSolicitarPoliticasPago.SetFocus
        Exit Sub
    End If
End If
Duracion = txtSolicitarPoliticasDuración
If Duracion = 0 Then
    ShowInfo "No se pueden ofrecer tratados de duración de 0 días"
    txtSolicitarPoliticasDuración.SetFocus
    Exit Sub
End If

SendData "CV2LOT" & cmbSolicitarPoliticas.ItemData(cmbSolicitarPoliticas.ListIndex) & CV2_CHAR_SEP_PARAMETROS & _
        Trim(Str(TipoTratado)) & CV2_CHAR_SEP_PARAMETROS & Pago & CV2_CHAR_SEP_PARAMETROS & Duracion
End Sub

Private Sub cmdFrameSolicitarIngresoMandarSolicitud_Click()
SendData "CV2SIC" & lista(lstListaDeClanes.ListIndex) & CV2_CHAR_SEP_PARAMETROS & txtMensajeAlLider
Unload Me
End Sub

Public Sub ParseInfoClan(sData As String)
Dim i As Integer
Dim parametros() As String
Dim valores() As String
ShowInfo "Llegaron datos.. procesando"
parametros = Split(sData, CV2_CHAR_SEP_PARAMETROS)


FrameInformacionDelClan.Caption = FrameInformacionDelClan.Caption & parametros(0)
For i = 1 To 10
    lblParseInfo(i - 1).Caption = parametros(i)
Next i
If parametros(11) = "+" Then
    FrameInformacionDelClan.Caption = FrameInformacionDelClan.Caption & " - Clan abierto"
Else
    FrameInformacionDelClan.Caption = FrameInformacionDelClan.Caption & " - Clan cerrado"
End If
valores = Split(parametros(12), CV2_CHAR_SEP_VALORES)
lstTitulos.Clear
If UBound(valores) = -1 Then
    lstTitulos.AddItem "Sin títulos"
Else
    For i = LBound(valores) To UBound(valores)
        lstTitulos.AddItem valores(i)
    Next i
End If

FrameInformacionDelClan.Visible = True
End Sub

Private Sub CmdFundarClanesAceptar_Click()
If Len(txtNombreNuevoClan.Text) = 0 Then
    ShowInfo "Debes ponerle un nombre al clan"
Else
    ShowInfo ""
    SendData "CV2FNC" & txtNombreNuevoClan & CV2_CHAR_SEP_PARAMETROS & txtDescripcionNuevoClan
    Unload Me
End If
End Sub

Private Sub CmdFundarClanesCancelar_Click()
EnableButtons
ShowFrame Ninguno
lstListaDeClanes.Enabled = True
End Sub


Private Sub CmdGmUpdateData_Click()
SendData "CV2ECV" & lista(lstListaDeClanes.ListIndex) & CV2_CHAR_SEP_PARAMETROS & IIf(chkGmAbierto.value = 1, "+", "-") & txtGMEditDesc.Text
End Sub

Private Sub cmdListaPjsEchar_Click()
If lstListaDeMiembros.ListIndex < 0 Then
    ShowInfo "Selecciona un miembro de la lista antes"
Else
    SendData "CV2LRM" & UCase$(Trim(lstListaDeMiembros.List(lstListaDeMiembros.ListIndex)))
    Unload Me
End If
End Sub


Private Sub cmdPoliticasMainVerGlobal_Click()
If cmbRelacionesEntreClanesGlobal.Text = "Clan no habilitado" Then
    ShowInfo "El clan no esta habilitado"
    Exit Sub
End If

If cmbRelacionesEntreClanesGlobal.ListIndex < 0 Then
    ShowInfo "selecciona un clan antes"
Else
    SendData "CV2GRG" & cmbRelacionesEntreClanesGlobal.ListIndex + 1
    ShowInfo "Datos enviados, esperando respuesta.."
End If
End Sub

Private Sub cmdPoliticasMainVerRelacion_Click()
If cmbRelacionesEntreClanes(0).Text = "Clan no habilitado" Or _
     cmbRelacionesEntreClanes(1).Text = "Clan no habilitado" Then
    ShowInfo "El clan no esta habilitado"
    Exit Sub
End If
If cmbRelacionesEntreClanes(0).ListIndex < 0 Or _
    cmbRelacionesEntreClanes(1).ListIndex < 0 Then
    ShowInfo "Selecciona los 2 clanes"
Else
    SendData "CV2GRI" & Trim(Str(cmbRelacionesEntreClanes(0).ListIndex + 1)) & CV2_CHAR_SEP_PARAMETROS & Trim(Str(cmbRelacionesEntreClanes(1).ListIndex + 1))
    ShowInfo "Datos enviados, esperando respuesta.."
End If
End Sub

Private Sub cmdrechazarTratado_Click()
SendData "CV2LRT" & lstTratadosPendientes.ItemData(lstTratadosPendientes.ListIndex)
lstTratadosPendientes.RemoveItem lstTratadosPendientes.ListIndex
If lstTratadosPendientes.ListCount = 0 Then
    cmdAceptarTratado.Enabled = False
    cmdrechazarTratado.Enabled = False
    lstTratadosPendientes.Enabled = False
End If

End Sub

Private Sub CmdRecursosAgregar_Click(Index As Integer)
SendData "CV2LPR" & Trim(txtRecursosCantidad(Index).Text) & CV2_CHAR_SEP_PARAMETROS & Trim(Str(Index))
End Sub

Private Sub CmdRecursosSacar_Click(Index As Integer)
Dim valor_sacar As Long
valor_sacar = Val(txtRecursosCantidad(Index).Text)
If valor_sacar > Val(txtCantidadRecurso(Index).Text) Then
    ShowInfo "Estas intentando sacar más recursos de los que tiene el clan."
Else
    SendData "CV2LSR" & Trim(Str(valor_sacar)) & CV2_CHAR_SEP_PARAMETROS & Trim(Str(Index))
End If
End Sub

Private Sub CmdrelacionesEntreClanesAtras_Click()
ShowFrame Ninguno
EnableButtons
End Sub

Private Sub cmdSolicitarIngresoCancelar_Click()
ShowFrame Ninguno
EnableButtons
End Sub

Private Sub CmdSolicitudesAceptar_Click()
If lstSolucitudesDeIngreso.ListIndex < 0 Then
    ShowInfo "Selecciona una solicitud de la lista"
Else
    SendData "CV2ASM" & UCase(Trim(lstSolucitudesDeIngreso.List(lstSolucitudesDeIngreso.ListIndex)))
    lstSolucitudesDeIngreso.RemoveItem lstSolucitudesDeIngreso.ListIndex
End If
If lstSolucitudesDeIngreso.ListCount = 0 Then
    CmdSolicitudesRechazar.Enabled = False
    CmdSolicitudesAceptar.Enabled = False
End If
End Sub

Private Sub CmdSolicitudesRechazar_Click()
If lstSolucitudesDeIngreso.ListIndex < 0 Then
    ShowInfo "Selecciona una solicitud de la lista"
Else
    SendData "CV2RSM" & UCase(Trim(lstSolucitudesDeIngreso.List(lstSolucitudesDeIngreso.ListIndex)))
    lstSolucitudesDeIngreso.RemoveItem lstSolucitudesDeIngreso.ListIndex
End If
If lstSolucitudesDeIngreso.ListCount = 0 Then
    CmdSolicitudesRechazar.Enabled = False
    CmdSolicitudesAceptar.Enabled = False
End If
End Sub

Private Sub CmdSolicitudTratadosAtras_Click()
DisableButtons
ShowFrame AdministracionViaLider
End Sub

Private Sub CmdVotarCancelar_Click()
ShowFrame Ninguno
EnableButtons
End Sub

Private Sub CmdVotarNuevoLider_Click()
If cmbMiembroParaSerLider.ListIndex < 0 Then
    ShowInfo "Primero selecciona un candidato"
Else
    ShowInfo "Datos enviados"
    SendData "CV2VOT" & UCase(Trim(cmbMiembroParaSerLider.List(cmbMiembroParaSerLider.ListIndex)))
End If
End Sub

Private Sub Command1_Click()
EnableButtons
ShowFrame Ninguno
End Sub

Public Sub ShowRelacionesEntreClanes(sData As String)
Dim valores() As String
valores = Split(sData, CV2_CHAR_SEP_VALORES)
ShowInfo "Llegaro datos, procesando"

Select Case Val(valores(0))
    Case 0
        lblEstadoEntreClanes(0).Caption = "Los clanes son NEUTRALES"
    Case 1
        lblEstadoEntreClanes(0).Caption = "Los clanes tienen PAZ"
    Case 2
        lblEstadoEntreClanes(0).Caption = "Los clanes estan en GUERRA"
End Select
If Val(valores(1)) > 0 Then
    lblEstadoEntreClanes(1) = "El tratado dura " & valores(1) & " días"
Else
    lblEstadoEntreClanes(1) = "El tratado es por tiempo indefinido"
End If
If Val(valores(2)) > 0 Then
    lblEstadoEntreClanes(2).Caption = "El primer clan le paga al segundo " & Val(valores(2)) & " diaros por el tratado"
Else
    lblEstadoEntreClanes(2).Caption = "El primer clan le cobra al segundo " & Val(valores(2)) * -1 & " diaros por el tratado"
End If
End Sub



Public Sub ShowRelacionesGlobales(sData As String)
Dim parametros() As String
Dim valores() As String
Dim i As Long
parametros = Split(sData, CV2_CHAR_SEP_PARAMETROS)
lstListaDeClanesEnPaz.Clear
If parametros(0) <> "" Then
    valores = Split(parametros(0), CV2_CHAR_SEP_VALORES)
    For i = LBound(valores) To UBound(valores)
        If Habilitado(Val(valores(i)) - 1) Then
            lstListaDeClanesEnPaz.AddItem lista(Val(valores(i)) - 1)
        Else
            lstListaDeClanesEnPaz.AddItem "Clan no habilitado"
        End If
    Next i
Else
    lstListaDeClanesEnPaz.AddItem "Nadie en paz con este clan"
End If
lstListaDeClanesEnGuerra.Clear
If parametros(1) <> "" Then
    valores = Split(parametros(1), CV2_CHAR_SEP_VALORES)
    For i = LBound(valores) To UBound(valores)
        If Habilitado(Val(valores(i)) - 1) Then
            lstListaDeClanesEnGuerra.AddItem lista(Val(valores(i)) - 1)
        Else
            lstListaDeClanesEnGuerra.AddItem "Clan no habilitado"
        End If
    Next i
Else
    lstListaDeClanesEnGuerra.AddItem "Nadie en guerra con este clan"
End If


End Sub

Private Sub cmdAdministrarLiderTratados_Click()
Dim i As Long
cmbSolicitarPoliticas.Clear
If NumClanes >= 1 Then
    For i = 0 To NumClanes
        If Habilitado(i) Then
            cmbSolicitarPoliticas.AddItem lista(i)
            cmbSolicitarPoliticas.ItemData(cmbSolicitarPoliticas.ListCount - 1) = i + 1
        Else
            cmbSolicitarPoliticas.AddItem "Clan no habilitado"
            cmbSolicitarPoliticas.ItemData(cmbSolicitarPoliticas.ListCount - 1) = -1
        End If
    Next i
Else
    ShowInfo "No hay clanes suficientes para tratados"
    Exit Sub
End If
'get tratados
SendData "CV2GTR"
ShowInfo "Esperando datos... espere por favor"
ShowFrame PoliticasSolicitudes
End Sub

Private Sub Form_Load()
VentanaClanesVisible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
VentanaClanesVisible = False
End Sub

Private Sub lstListaDeClanes_Click()
cmdAdminClanAceptar.Enabled = True
End Sub

Private Sub lstSolucitudesDeIngreso_Click()
Dim i As Long
Dim j As Long
Dim valores() As String
i = lstSolucitudesDeIngreso.ListIndex + 1
If i <= UBound(pedidos) Then
    valores = Split(pedidos(i), CV2_CHAR_SEP_VALORES)
    For j = LBound(valores) To UBound(valores)
        lblSolicitudesData(j).Caption = TextoDescripcionSolicitudIngreso(CInt(j), valores(j))
    Next j
End If
End Sub
Private Function TextoDescripcionSolicitudIngreso(i As Integer, sData As String)
Select Case i
    Case 0
        TextoDescripcionSolicitudIngreso = "Clase: " & sData
    Case 1
        Select Case sData
            Case "0"
                TextoDescripcionSolicitudIngreso = "Ejercito: Neutral"
            Case "1"
                TextoDescripcionSolicitudIngreso = "Ejercito: Armada"
            Case "2"
                TextoDescripcionSolicitudIngreso = "Ejercito: Caos"
        End Select
    Case 2
        TextoDescripcionSolicitudIngreso = "Fecha solicitud: " & sData
    Case 3
        TextoDescripcionSolicitudIngreso = "Nivel al momento de solicitud: " & sData
    Case 4
        TextoDescripcionSolicitudIngreso = "Nombre del Pj: " & sData
    Case 5
        TextoDescripcionSolicitudIngreso = "Pedido: " & sData
    Case 6
        TextoDescripcionSolicitudIngreso = "Raza: " & sData
    Case 7
        TextoDescripcionSolicitudIngreso = "Reputación al momento del pedido: " & sData
    Case Else
    Debug.Print "ASODJKASD  " & i
End Select
End Function


Public Sub ShowRecursos(sData As String)
Dim parametros() As String
Dim i As Long
ShowInfo "Llegaron datos, procesando"
parametros = Split(sData, CV2_CHAR_SEP_PARAMETROS)
For i = LBound(parametros) To UBound(parametros)
    txtCantidadRecurso(i) = parametros(i)
Next i
frameAdministracionRecursos.Visible = True
End Sub


Private Sub lstTratadosPendientes_Click()
Dim pedidoIndex As Long
Dim valores() As String
Dim sTmp As String
cmdAceptarTratado.Enabled = True
cmdrechazarTratado.Enabled = True
pedidoIndex = lstTratadosPendientes.ListIndex
valores = Split(SolicitudesPendientesPoliticas(pedidoIndex), CV2_CHAR_SEP_VALORES)

lblsSolicitudTratados(0) = lblsSolicitudTratados(0) & lista(Val(valores(0)) - 1)
If valores(1) = "0" Then
    sTmp = " NEUTRALIDAD"
ElseIf valores(1) = "1" Then
    sTmp = " PAZ"
Else
    sTmp = " GUERRA"
End If
lblsSolicitudTratados(1) = lblsSolicitudTratados(1) & sTmp
If Val(valores(2)) < 0 Then
    lblsSolicitudTratados(2) = lblsSolicitudTratados(2) & " pagarte por el tratado" & valores(2) * -1
Else
    lblsSolicitudTratados(2) = lblsSolicitudTratados(2) & " cobrarte por el tratado" & valores(2)
End If
If Val(valores(3)) > 0 Then
    lblsSolicitudTratados(3) = lblsSolicitudTratados(3) & " la duración del tratado es " & valores(3)
Else
    lblsSolicitudTratados(3) = lblsSolicitudTratados(3) & " la duración del tratado es indeterminada"
End If
lstTratadosPendientes.Enabled = False
End Sub

Private Sub tmrShowInfo_Timer()
lblShowInfo.Caption = ""
tmrShowInfo.Enabled = False
End Sub

Private Sub ShowInfo(sData As String, Optional Permanente As Boolean = False)
lblShowInfo.Caption = sData
tmrShowInfo.Enabled = False
If Not Permanente Then tmrShowInfo.Enabled = True
End Sub

Private Sub EnableButtons()
Dim i As Integer
For i = 1 To cmdOpciones.count - 1
    cmdOpciones(i).Enabled = True
Next i
End Sub


Private Sub DisableButtons()
Dim i As Integer
For i = 1 To cmdOpciones.count - 1
    cmdOpciones(i).Enabled = False
Next i
End Sub

Private Sub cmdOpciones_Click(Index As Integer)
Dim i As Long
DisableButtons
Select Case Index
    Case 0
    'cerrar
        Unload Me
    Case 1
    'ingreso
        If HayClanSeleccionado Then
            frameSolicitarIngreso.Visible = True
            For i = 1 To cmdOpciones.count - 1
                cmdOpciones(i).Enabled = False
            Next i
            lstListaDeClanes.Enabled = False
            frameSolicitarIngreso.Caption = frameSolicitarIngreso.Caption & lista(lstListaDeClanes.ListIndex)
        Else
            If lblShowInfo.Caption <> "No es un clan habilitado el seleccionado!" Then _
                ShowInfo "Primero selecciona el clan de la lista"
        End If
    Case 2
    ' salir clan
        SendData "CV2DEC"
        Unload Me
    Case 3
    'info
        If HayClanSeleccionado Then
            For i = 1 To cmdOpciones.count - 1
                cmdOpciones(i).Enabled = False
            Next i
            lstListaDeClanes.Enabled = False
            SendData "CV2GCI" & lista(lstListaDeClanes.ListIndex)
            ShowInfo "Esperando datos... espere por favor"
        Else
            If lblShowInfo.Caption <> "No es un clan habilitado el seleccionado!" Then _
                ShowInfo "Primero selecciona el clan de la lista"
        End If
            
    Case 4
    'Fundar clan
        For i = 1 To cmdOpciones.count - 1
            cmdOpciones(i).Enabled = False
        Next i
        lstListaDeClanes.Enabled = False
        frameFundarClan.Visible = True
    Case 5
    'Administrar
        For i = 1 To cmdOpciones.count - 1
            cmdOpciones(i).Enabled = False
        Next i
        lstListaDeClanes.Enabled = False
        frameAdministrarClanLider.Visible = True
    Case 6
    'AdminGM
        ShowList True
        frameAdminGM.Visible = True
        'If HayClanSeleccionado Then
        '    For i = 1 To cmdOpciones.Count - 1
        '        cmdOpciones(i).Enabled = False
        '    Next i
        '    lstListaDeClanes.Enabled = False
        '    SendData "CV2AGM" & lista(lstListaDeClanes.ListIndex)
        '    ShowInfo "Esperando datos... espere por favor"
        'Else
        '    ShowInfo "Primero selecciona el clan de la lista"
        'End If
        
    Case 7
    'políticas
    lstListaDeClanesEnPaz.Clear
    lstListaDeClanesEnGuerra.Clear
    cmbRelacionesEntreClanes(0).Clear
    cmbRelacionesEntreClanes(1).Clear
    cmbRelacionesEntreClanesGlobal.Clear
    If NumClanes > 0 Then
        For i = LBound(lista) To UBound(lista)
            If Habilitado(i) Then
                cmbRelacionesEntreClanes(0).AddItem lista(i)
                cmbRelacionesEntreClanes(1).AddItem lista(i)
                cmbRelacionesEntreClanesGlobal.AddItem lista(i)
            Else
                cmbRelacionesEntreClanes(0).AddItem "Clan no habilitado"
                cmbRelacionesEntreClanes(1).AddItem "Clan no habilitado"
                cmbRelacionesEntreClanesGlobal.AddItem "Clan no habilitado"
            End If
        Next i
    End If
    ShowFrame PoliticasMain
    Case 8
    'votar lider
        SendData "CV2GCL"
        ShowInfo "Datos enviados... espere"
End Select
End Sub

Public Sub HandleListaClanes(sData As String)
Dim i As Long
Dim valores() As String
valores = Split(sData, CV2_CHAR_SEP_VALORES)
NumClanes = UBound(valores)
If NumClanes >= 0 Then
    ReDim lista(LBound(valores) To UBound(valores))
    ReDim Habilitado(LBound(valores) To UBound(valores))
    For i = LBound(valores) To UBound(valores)
        lista(i) = mid$(valores(i), 2)
        Habilitado(i) = mid$(valores(i), 1, 1) = "+"
    Next i
End If
End Sub

Public Sub showBotones(flags As Long)
'IniFlags = flags
Dim i As Long
Dim prox As Long
cmdOpciones(1).Visible = FlagOn(flags, CV2_VENTANACLANES_FLAG_BTN_SolicitarIngreso)
cmdOpciones(2).Visible = FlagOn(flags, CV2_VENTANACLANES_FLAG_BTN_SalirClan)
cmdOpciones(3).Visible = FlagOn(flags, CV2_VENTANACLANES_FLAG_BTN_Informacion)
cmdOpciones(4).Visible = FlagOn(flags, CV2_VENTANACLANES_FLAG_BTN_FundarClan)
cmdOpciones(5).Visible = FlagOn(flags, CV2_VENTANACLANES_FLAG_BTN_Administrar)
cmdOpciones(6).Visible = FlagOn(flags, CV2_VENTANACLANES_FLAG_BTN_AdministrarGM)
cmdOpciones(7).Visible = FlagOn(flags, CV2_VENTANACLANES_FLAG_BTN_Politicas)
cmdOpciones(8).Visible = FlagOn(flags, CV2_VENTANACLANES_FLAG_BTN_Votar)
prox = 0
DoEvents
For i = 1 To cmdOpciones.count - 1
    If cmdOpciones(i).Visible Then
        cmdOpciones(i).top = 56 + 32 * prox
        prox = prox + 1
    End If
Next i
End Sub
Private Function FlagOn(flags As Long, flag As Long) As Boolean
FlagOn = (flags And flag) = flag
End Function


Public Sub ShowList(Optional VerTipoGM As Boolean = False)
Dim i As Long
lstListaDeClanes.Clear
If NumClanes >= 0 Then
    For i = 0 To NumClanes
        If Habilitado(i) Or VerTipoGM Then
            lstListaDeClanes.AddItem IIf(Habilitado(i), "", "*") & lista(i)
        Else
            lstListaDeClanes.AddItem "Clan no habilitado"
        End If
    Next i
Else
    lstListaDeClanes.AddItem "No hay clanes Creados"
    For i = 1 To cmdOpciones.count - 1
        cmdOpciones(i).Visible = False
    Next i
End If
End Sub

Private Function HayClanSeleccionado() As Boolean
HayClanSeleccionado = lstListaDeClanes.ListIndex > -1
If HayClanSeleccionado Then
If Not Habilitado(lstListaDeClanes.ListIndex) Then
    ShowInfo "No es un clan habilitado el seleccionado!"
    HayClanSeleccionado = False
End If
End If
End Function

Public Sub ParseAdminClanInfo(sData As String)
ShowInfo "Llegaron datos... procesando"
frameAdministrarClanLider.Visible = False
If mid$(sData, 1, 1) = "+" Then
    chkAdminClanAceptarSolicitudes.value = 1
Else
    chkAdminClanAceptarSolicitudes.value = 0
End If
txtAdminClanDescripcion = mid$(sData, 3)
frameAdministrarClanLider.Visible = False
frameAdministracionClan.Visible = True
End Sub

Private Sub txtAdminClanDescripcion_Change()
cmdAdminClanAceptar.Enabled = True
End Sub

Private Sub txtRecursosCantidad_Change(Index As Integer)
If Index <> 0 Then
    If Val(txtRecursosCantidad(Index).Text) > 10000 Then
        ShowInfo "Estos recursos no se pueden agregar de más de 10000 a la vez"
        txtRecursosCantidad(Index).Text = "10000"
    End If
Else
    If Val(txtRecursosCantidad(0).Text) > 100000 Then
        ShowInfo "Este recurso no se puede agregar de más de 100000 a la vez"
        txtRecursosCantidad(0).Text = "100000"
    End If
End If

End Sub

Public Sub ShowFrame(Cual As eFrames)
If Cual = AdministracionMiembros Then
    frameADM.Visible = True
Else
    frameADM.Visible = False
End If

If Cual = AdministracionClan Then
    frameAdministracionClan.Visible = True
Else
    frameAdministracionClan.Visible = False
End If

If Cual = AdministracionViaGM Then
    frameAdminGM.Visible = True
Else
    frameAdminGM.Visible = False
End If

If Cual = AdministracionViaLider Then
    frameAdministrarClanLider.Visible = True
Else
    frameAdministrarClanLider.Visible = False
End If

If Cual = FundarClan Then
    frameFundarClan.Visible = True
Else
    frameFundarClan.Visible = False
End If

If Cual = InformacionDelClan Then
    FrameInformacionDelClan.Visible = True
Else
    FrameInformacionDelClan.Visible = False
End If

If Cual = SolicitarIngreso Then
    frameSolicitarIngreso.Visible = True
Else
    frameSolicitarIngreso.Visible = False
End If

If Cual = AdministracionRecursos Then
    frameAdministracionRecursos.Visible = True
Else
    frameAdministracionRecursos.Visible = False
End If
If Cual = Votacion Then
    frameVotar.Visible = True
Else
    frameVotar.Visible = False
End If
If Cual = PoliticasMain Then
    framePoliticasMain.Visible = True
Else
    framePoliticasMain.Visible = False
End If
If Cual = PoliticasSolicitudes Then
    frameSolicitudesPoliticas.Visible = True
Else
    frameSolicitudesPoliticas.Visible = False
End If

End Sub

Public Sub ShowVotar(sData As String)
Dim valores() As String
Dim i As Long
valores = Split(sData, CV2_CHAR_SEP_VALORES)
cmbMiembroParaSerLider.Clear
For i = LBound(valores) To UBound(valores)
    cmbMiembroParaSerLider.AddItem valores(i)
Next i
ShowFrame Votacion
End Sub

Public Sub HandleTratadosPendientes(sData As String)
'solo sirve esto si estamos viendo lassollicitudes
If Not frameSolicitudesPoliticas.Visible Then Exit Sub
Dim parametros() As String
Dim i As Long
Dim valores() As String
Dim ClanLocal As Long
parametros = Split(sData, CV2_CHAR_SEP_PARAMETROS)
lstTratadosPendientes.Clear
ReDim SolicitudesPendientesPoliticas(LBound(parametros) To UBound(parametros))
If UBound(parametros) = 0 Then
    If parametros(0) = "" Then
        ShowInfo "no hay pedidos pendentes"
        Exit Sub
    End If
End If
lstTratadosPendientes.Enabled = True
ShowInfo "Llegaron datos... procesando"
For i = LBound(parametros) To UBound(parametros)
    SolicitudesPendientesPoliticas(i) = parametros(i)
    ClanLocal = General_Field_Read(1, SolicitudesPendientesPoliticas(i), Asc(CV2_CHAR_SEP_VALORES)) - 1
    If Habilitado(ClanLocal) Then
        lstTratadosPendientes.AddItem lista(ClanLocal)
        lstTratadosPendientes.ItemData(lstTratadosPendientes.ListCount - 1) = ClanLocal + 1
    Else
        lstTratadosPendientes.AddItem "Clan no habilitado"
        lstTratadosPendientes.ItemData(lstTratadosPendientes.ListCount - 1) = -1
    End If
Next i
End Sub
