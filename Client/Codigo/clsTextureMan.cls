VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTextureMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**************************************************************
' CTextureMan.cls - Maneja todas las texturas del cliente,
' usando la menor cantidad de memoria posible
'
' Escrito por Maraxus (Juan Martín Sotuyo Dodero - juansotuyo@hotmail.com)
'**************************************************************
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

Option Explicit

Private Type TEXT_DB_ENTRY
    FileName As Integer
    UltimoAcceso As Long
    texture As Direct3DTexture8
    Dimension As Long
    Size As Long
End Type

Private mGraficos() As TEXT_DB_ENTRY
Private mD3D As D3DX8
Private device As Direct3DDevice8

Private mMaxEntries As Integer
Private mCantidadGraficos As Integer
Private mFreeMemoryBytes As Long

Private GrhPath As String

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub Class_Initialize()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'
'**************************************************************
End Sub

Private Sub Class_Terminate()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'
'**************************************************************
On Error Resume Next
    Dim i As Long
    
    For i = 1 To mCantidadGraficos
        Set mGraficos(i).texture = Nothing
    Next i
    
    Erase mGraficos
    
End Sub

Public Property Let Path(ByVal inString As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'
'**************************************************************
    GrhPath = inString
End Property

Public Function GetDimension(ByVal FileName As Integer, ByRef Index As Integer) As Long
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 2/08/2006
'
'**************************************************************

    Index = ObtenerIndice(FileName)

    If Index > 0 Then
        With mGraficos(Index)
            'Ultimo acceso
            .UltimoAcceso = GetTickCount
            'Devuelvo una texture con el grafico cargado
            GetDimension = .Dimension
        End With
    Else    'Gráfico no cargado
        'Vemos si puedo agregar uno a la lista
        If mMaxEntries = mCantidadGraficos Then
            'Sacamos el que hace más que no usamos para hacer lugar
            BorraMenosUsado
        End If
        
        'Agrego una textura nueva a la lista
        Index = CrearGrafico_Ex(FileName, Index)
        GetDimension = mGraficos(Index).Dimension

    End If
End Function

Public Sub SetTexture(ByVal Index As Integer)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/08/2006
'Modified by Augusto José Rando
'Nunca se llama a esto antes que GetDimension
'
'**************************************************************

    If Index > 0 Then
        With mGraficos(Index)
            'Ultimo acceso
            .UltimoAcceso = GetTickCount
            'Hacemos el SetTexture en el device
            device.SetTexture 0, .texture
        End With
    Else    'Gráfico no cargado
        Debug.Print "¿Qué hago acá? Grh no cargado grr... se llamó antes que GetDimension."
    End If
    
End Sub

Private Function ObtenerIndice(ByVal FileName As Integer) As Integer
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'Busqueda binaria para hallar la texture deseada
'**************************************************************
    Dim max As Integer  'Max index
    Dim min As Integer  'Min index
    Dim mid As Integer  'Middle index
    
    min = 1
    mid = 1
    max = mCantidadGraficos
    Do While min <= max
        mid = (min + max) \ 2
        If FileName < mGraficos(mid).FileName Then
            max = mid - 1
        ElseIf FileName > mGraficos(mid).FileName Then
            min = mid + 1
        Else
            ObtenerIndice = mid
            Exit Function
        End If
    Loop
    
    'Maraxus - usado para binary insertion
    ObtenerIndice = Not mid
    
End Function

Public Function Init(ByRef D3D8 As D3DX8, ByRef d3d_device As Direct3DDevice8, ByVal MaxMemory As Long, ByVal MaxEntries As Long) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'
'**************************************************************
    
    mMaxEntries = MaxEntries
    
    If mMaxEntries < 1 Then 'por lo menos 1 gráfico
        Exit Function
    End If
    
    mCantidadGraficos = 0
    
    'Seteamos el objeto
    Set mD3D = D3D8
    Set device = d3d_device
    mFreeMemoryBytes = MaxMemory
    
    Init = True
End Function

Public Sub BorrarTodo()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'Vacia la lista de graficos
'**************************************************************
    Dim i As Long
    
    For i = 1 To mCantidadGraficos
        Set mGraficos(i).texture = Nothing
    Next i
    
    ReDim mGraficos(0) As TEXT_DB_ENTRY
    mCantidadGraficos = 0
End Sub

Public Function Borrar(ByVal FileName As Integer) As Integer
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'Borra un grafico
'**************************************************************
    Dim indice As Long
    
    'Obtenemos el ìndice
    indice = ObtenerIndice(FileName)
    
    If indice > 0 Then
        'Lo eliminamos
        With mGraficos(indice)
            Set .texture = Nothing
            .Dimension = 0
            mFreeMemoryBytes = mFreeMemoryBytes + .Size
            .Size = 0
        End With
        
        'Movemos para atrás el resto de la lista
        If indice < mCantidadGraficos Then
            Dim i As Long
            For i = indice To mCantidadGraficos - 1
                mGraficos(i) = mGraficos(i + 1)
            Next i
        End If
        
        'Actualizamos el número de gráficos
        mCantidadGraficos = mCantidadGraficos - 1
        
        'Redimencionamos la lista
        ReDim Preserve mGraficos(1 To mCantidadGraficos) As TEXT_DB_ENTRY
    End If
End Function

Private Function CrearGrafico_Ex(ByVal Archivo As Integer, ByVal Index As Integer) As Integer
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 5/04/2005
'Crea una textura desde memoria
'**************************************************************
On Error GoTo errhandler
  
    Dim surface_desc As D3DSURFACE_DESC, bytArr() As Byte
            
    If Index < 0 Then
        Index = Not Index   ' Obtenemos el índice donde debe ser insertado
        ReDim Preserve mGraficos(1 To mCantidadGraficos + 1) As TEXT_DB_ENTRY    ' Incrementamos la lista
        If Index < mCantidadGraficos + 1 Then
            Dim i As Long
            For i = mCantidadGraficos To Index Step -1
                mGraficos(i + 1) = mGraficos(i)
            Next i
        End If
    End If
    
    With mGraficos(Index)
        'Nombre
        .FileName = Archivo
        
        'Ultimo acceso
        .UltimoAcceso = GetTickCount
        
        If Extract_File_Ex(Graphics, GrhPath, LTrim$(CStr(Archivo)) & ".bmp", bytArr) Then
            'Cargamos el gráfico y seteamos la Color Key
            Set .texture = mD3D.CreateTextureFromFileInMemoryEx(device, bytArr(0), _
                UBound(bytArr) + 1, D3DX_DEFAULT, D3DX_DEFAULT, 1, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_NONE, _
                D3DX_FILTER_NONE, &HFF000000, ByVal 0, ByVal 0)
            
            .texture.GetLevelDesc 0, surface_desc
            .Dimension = surface_desc.Width
            .Size = surface_desc.Size
                        
            mFreeMemoryBytes = mFreeMemoryBytes - surface_desc.Size
                        
        End If
        
    End With
    
    'Aumentamos la cantidad de gráficos
    mCantidadGraficos = mCantidadGraficos + 1
    
    'Devolvemos el ìndice en que lo cargamos
    CrearGrafico_Ex = Index
Exit Function

errhandler:
    'Ocurrió un error, la causa más probable es que el archivo no exista
    'If General_File_Exists(Windows_Temp_Dir & Archivo & ".bmp", vbNormal) Then
    '    Delete_File Windows_Temp_Dir & Archivo & ".bmp"
    'End If

End Function

Private Sub BorraMenosUsado()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'Borra el grafico menos usado y mueve la lista hacia atrás para evitar el hueco
'**************************************************************
    Dim Valor As Long
    Dim i As Long
    Dim MenosUsado As Long
    
    'Inicializamos todo
    Valor = mGraficos(1).UltimoAcceso
    MenosUsado = 1
    
    'Buscamos cual es el que lleva más tiempo sin ser utilizado
    For i = 1 To mCantidadGraficos
        If mGraficos(i).UltimoAcceso < Valor Then
            Valor = mGraficos(i).UltimoAcceso
            MenosUsado = i
        End If
    Next i
    
    'Borramos la texture
    Set mGraficos(MenosUsado).texture = Nothing
    mFreeMemoryBytes = mFreeMemoryBytes + mGraficos(MenosUsado).Size
    
    'Mover la lista para atrás
    If MenosUsado < mCantidadGraficos Then
        For i = MenosUsado To mCantidadGraficos - 1
            mGraficos(i) = mGraficos(i + 1)
        Next i
    End If
    
    'Disminuimos el contador
    mCantidadGraficos = mCantidadGraficos - 1
    
    ' Redimensionamos el array
    ReDim Preserve mGraficos(1 To mCantidadGraficos) As TEXT_DB_ENTRY
End Sub

Public Property Get MaxEntries() As Integer
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'
'**************************************************************
    MaxEntries = mMaxEntries
End Property

Public Property Let MaxEntries(ByVal vNewValue As Integer)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'
'**************************************************************
    mMaxEntries = vNewValue
End Property

Public Property Get CantidadGraficos() As Integer
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'
'**************************************************************
    CantidadGraficos = mCantidadGraficos
End Property

Public Sub BorraTimer()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'Cada un minuto revisamos que texturas se pueden eliminar
'**************************************************************
    Dim loopc As Long
    Dim offset As Integer
    Dim TiempoActual As Long
    
    'Barrin: we delete only if we have less than 10 MB of free memory
    If mFreeMemoryBytes > 10485760 Then Exit Sub
    
    TiempoActual = GetTickCount
    
    For loopc = 1 To mCantidadGraficos
        'Revisar cuales usamos
        If (mGraficos(loopc).Dimension <= 32) Or (mGraficos(loopc).UltimoAcceso > TiempoActual - 60000) Then
            'Si el offset está seteado, movemos para atrás la lista
            If offset Then
                mGraficos(loopc - offset) = mGraficos(loopc)
            End If
        Else
            'Eliminar de memoria
            Set mGraficos(loopc).texture = Nothing
            mGraficos(loopc).Dimension = 0
            mFreeMemoryBytes = mFreeMemoryBytes + mGraficos(loopc).Size
            mGraficos(loopc).Size = 0
            offset = offset + 1
        End If
    Next loopc
    
    'Actualizar el número de gráficos
    mCantidadGraficos = mCantidadGraficos - offset
End Sub

Public Function CargarGrafico(ByVal FileName As Integer) As Boolean
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 8/17/2005
'
'**************************************************************
    Dim Index As Integer
    
    If FileName <= 0 Then Exit Function
    Index = ObtenerIndice(FileName)
    
    If Index > 0 Then
        CargarGrafico = False
    Else    'Gráfico no cargado
        'Vemos si puedo agregar uno a la lista
        If mMaxEntries = mCantidadGraficos Then
            'Sacamos el que hace más que no usamos
            BorraMenosUsado
        End If
        
        'Agrego una textura nueva a la lista
        Call CrearGrafico_Ex(FileName, Index)
        CargarGrafico = True
    End If
End Function
