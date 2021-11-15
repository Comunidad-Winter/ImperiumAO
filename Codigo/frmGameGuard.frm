VERSION 5.00
Begin VB.Form frmGameGuard 
   BorderStyle     =   0  'None
   Caption         =   "ImperiumAO 1.3"
   ClientHeight    =   2850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8055
   Icon            =   "frmGameGuard.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   190
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   537
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmGameGuard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Me.Picture = General_Load_Picture_From_Resource("gameguard.bmp")
Call General_Form_On_Top_Set(Me, True)
End Sub
