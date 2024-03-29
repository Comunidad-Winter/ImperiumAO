VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum CursorType
    E_NORMAL = 0
    E_ATTACK
    E_ARROW
    E_CAST
    E_WAIT
    E_SHOOT
    E_SHIP
End Enum

Private Const NUM_CURSORS = 6

Private hndlList(0 To NUM_CURSORS) As IPictureDisp

Public Sub Parse_Form(ByRef aFrm As Form, Optional ByVal cType As CursorType = E_NORMAL)

On Error Resume Next

Dim aControl As Control
Dim lngHandle As Long

If CursoresStandar = 1 Then
    Select Case cType
        Case E_NORMAL
            lngHandle = vbDefault
        Case E_ATTACK
            lngHandle = vbCrosshair
        Case E_ARROW
            lngHandle = vbCrosshair
        Case E_CAST
            lngHandle = vbCrosshair
        Case E_WAIT
            lngHandle = vbHourglass
        Case E_SHOOT
            lngHandle = vbCrosshair
        Case E_SHIP
            lngHandle = vbCrosshair
    End Select
    
    aFrm.MouseIcon = Nothing
    aFrm.MousePointer = lngHandle
    
    For Each aControl In aFrm.Controls
        If LenB(aControl.Caption) > 0 Then aControl.Caption = Locale_Parse_GUI(aControl.Caption)
    Next
    
Else
    For Each aControl In aFrm.Controls
        aControl.MouseIcon = hndlList(cType)
        aControl.MousePointer = vbCustom
        
        If LenB(aControl.Caption) > 0 Then aControl.Caption = Locale_Parse_GUI(aControl.Caption)
    Next
    
    aFrm.MouseIcon = hndlList(cType)
    aFrm.MousePointer = vbCustom
End If

If LenB(aFrm.Caption) > 0 Then aFrm.Caption = Locale_Parse_GUI(aFrm.Caption)

End Sub

Public Function GetDefaultCursor(ByRef aFrm As Form) As CursorType

Dim lngHandle As Long

Select Case aFrm.MouseIcon.handle
    Case hndlList(E_NORMAL).handle
        GetDefaultCursor = E_NORMAL
    Case hndlList(E_ATTACK).handle
        GetDefaultCursor = E_ATTACK
    Case hndlList(E_ARROW).handle
        GetDefaultCursor = E_ARROW
    Case hndlList(E_CAST).handle
        GetDefaultCursor = E_CAST
    Case hndlList(E_WAIT).handle
        GetDefaultCursor = E_WAIT
    Case hndlList(E_SHOOT).handle
        GetDefaultCursor = E_SHOOT
    Case hndlList(E_SHIP).handle
        GetDefaultCursor = E_SHIP
End Select

End Function

Public Sub Init()

Set hndlList(E_NORMAL) = LoadResPicture("MAIN", vbResCursor)
Set hndlList(E_ATTACK) = LoadResPicture("ATTACK", vbResCursor)
Set hndlList(E_ARROW) = LoadResPicture("ARROW", vbResCursor)
Set hndlList(E_CAST) = LoadResPicture("CAST", vbResCursor)
Set hndlList(E_WAIT) = LoadResPicture("WAIT", vbResCursor)
Set hndlList(E_SHOOT) = LoadResPicture("SHOOT", vbResCursor)
Set hndlList(E_SHIP) = LoadResPicture("SHIP", vbResCursor)

End Sub
