VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTileEngineX"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************
'clsTileEngineX - ORE Direct3D 8 Tile Engine - v0.5.0
'
'Graphics Engine.
'
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
'Contributors History
'   When releasing modifications to this source file please add your
'   date of release, name, email, and any info to the top of this list.
'   Follow this template:
'    XX/XX/200X - Your Name Here (Your Email Here)
'       - Your Description Here
'       Sub Release Contributors:
'           XX/XX/2003 - Sub Contributor Name Here (SC Email Here)
'               - SC Description Here
'*****************************************************************
'
'18/1/2005 - Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'   - Add: Particle_Type_Get method
'   - Add: Chats over head
'   - Add: Engine_Resolution_Change method
'   - Add: Particle_Data_Load_All method and related type
'   - Add: Grh_Render_Advance which allows grh resizing
'   - Add: Light_Move_By_Head method
'   - Change: Map_Render nolonger attempts to render anything outside map bounds, so it's faster, several FPS gained
'   - Change: Removed the clipping around edges code at the end of Map_Render since it was useless. Related vars and methods were removed as well
'   - Change: Engine now renders VSYNC since it's FPS have been really improved
'   - Change: Removed Engine_Deinitialize method, now it's all done when setting the engine object to Nothing
'   - Change: Map format was redesigned
'   - Change: Map_Base_Light_Fill was edited to store the base light color to the map structure
'   - Change: Font_Make now saves the StdFont so it can be recreated after changing resolution
'   - Change: Geometry_Create_Box method was slightly modified to improve speed
'   - Change: Light_Render method was recoded so now it's MUCH faster
'   - Change: Alpha blending is enabled only once at start up (or on reset) instead of every time something is rendered
'   - Change: the alpha rendering mode is only changed when alpha blending is required, and then set back to normal (less code executed, faster code)
'   - Change: GUI_Grh_Render was enhaced to support further customization on the way the grh is displayed
'   - Change: Almost every variable type was changed to save memory. Some variables were directly removed
'   - Change: Removed layer 5 since it wasn't used at all
'   - Change: When placing a NPC (map editor) it creates a char to be rendered, no need to set show special tile  to true anymore
'   - Change: Items are rendered after layer 2, no need to set show special tile  to true anymore
'   - Change: show_special_tiles is now show_exit tiles, related methods were renamed appropiately
'   - Change: Particle_Render was completely recoded, now it's faster and simpler
'   - Change: FontAlignment was improved to support properly the 9 possible alignments
'   - Change: Got rid of the header in the grh.dat file
'   - Change: The array of events in Input_Mouse_Poll was set to 1000 instead of 5000 since that's the limit set when creating the device
'   - Change: Several more small changes all over
'   - Fix: The mouse map pos is now updated even if the mouse is not being moved (since we may be scrolling the map)
'   - Fix: Engine_Render_Start now is able to avoid being called without having ended the scene
'   - Fix: Fonts are now destroyed to avoid memory leak
'   - Fix: Input_Mouse_Poll can now tell if a mouse button is pressed or not properly
'   - Fix: Memory-leak issue when deiitializing
'   - Fix: Memory-leak issue when rendering to a hDC
'   - Fix: The device is now properly reset
'   - Fix: The engine no longer crushes every now and then when you hit Ctrl-Alt-Del
'
'05/15/2003 - Aaron Perkins (aaron@baronsoft.com)
'   -Add: GUI_Grh_Render
'   -Change: Text_Render is now called GUI_Text_Render
'       -Added word_wrap option
'   Sub Release Contributors:
'       05/15/2003 - Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'           - Add: GUI_Box_Filled_Render
'           - Add: GUI_Box_Outline_Render
'
'05/12/2003 - Aaron Perkins (aaron@baronsoft.com)
'   -Add: Char_Data_Grh_Index_Get
'
'05/08/2003 - Aaron Perkins (aaron@baronsoft.com)
'   -Change: Moved all the General_X methods to a external modGeneral.bas
'
'05/07/2003 - Aaron Perkins (aaron@baronsoft.com)
'   -Fix: Map_Save and Map_Load format changed. Added angle and alpha values.
'   Sub Release Contributors:
'       05/07/2003 - Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'           - Fix: Geometry_Create_Sqaure - fixed rotation
'
'05/02/2003 - Aaron Perkins (aaron@baronsoft.com)
'   - Fix: bug with Light_Move and base light values
'   - Fix: Grh_Get_Info and Grh_Render_To_Hdc when using animated grhs
'   - Add: Base light values can now be specfied for each corner
'      - Changed .map format. Old maps will not work anymore
'      - Modified Map_Base_Light_Set
'   - Remove: Light_Remove_XY (Use Light_Remove Map_Light_Get(X,Y) instead)
'
'04/29/2003 - Aaron Perkins (aaron@baronsoft.com)
'   - Add: Grh_Render_To_Hdc
'   - Add: Map_Save_To_File and Map_Load_From_File
'   - Add: Text_Render method
'   - Change: file name to clsTileEngineX.cls
'   - Change: Split Engine_Render into Engine_Render_Start and Engine_Render_End
'   - Change: Method Names
'       -Engine_Stats_Render
'       -Engine_Blocked_Tiles_Show_Toggle
'   Sub Release Contributors:
'       04/28/2003 - Fredrik Alexandersson (fredrik@oraklet.zzn.com)
'           - Add: Grh_Add_GrhList_To_ListBox method
'           - Change: Particle_Render (added 3 new stream types)
'
'04/07/2003 - Aaron Perkins (aaron@baronsoft.com)
'   - First release
'   Sub Release Contributors:
'       03/31/2003 - Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'           - Add: Light_Erase method
'
'04/11/2005 - Augusto José Rando (barrin@imperiumao.com.ar)
'           - Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'
'   - Changes for ImperiumAO (new features, functions and procedures)
'   - Special ORE 1.0 Ed. by Juan Martín Sotuyo & Augusto José Rando
'*****************************************************************

'***************************
'Required Externals
'***************************
'Reference to dx8vb.dll
'   - URL: http://www.microsoft.com/directx
'***************************

Option Explicit

'***************************
'Added by Augusto José Rando
'***************************

Private Type Fuente
    Tamanio As Integer
    Caracteres(0 To 255) As Integer 'indice de cada letra
End Type

Private Fuentes(1 To 2) As Fuente

Private MinXBorder As Byte
Private MaxXBorder As Byte
Private MinYBorder As Byte
Private MaxYBorder As Byte

Private Const XMaxMapSize As Integer = 100
Private Const XMinMapSize As Integer = 1
Private Const YMaxMapSize As Integer = 100
Private Const YMinMapSize As Integer = 1

Private Const blind_color As Long = -13487566

Private Const fire_grh_index As Long = 1521

Private engine_erase_graphics As Byte

Private Type Position
    x As Integer
    y As Integer
End Type

Private Type Grh
    GrhIndex As Long
    framecounter As Single
    speed As Single
    started As Byte
    alpha_blend As Boolean
    angle As Single
End Type

Private Type GrhData
    sX As Integer
    sY As Integer
    FileNum As Integer
    pixelWidth As Integer
    pixelHeight As Integer
    TileWidth As Single
    TileHeight As Single
    NumFrames As Integer
    Frames() As Integer
    speed As Integer
    mini_map_color As Long
End Type

Private Type BodyData
    Walk(1 To 8) As Grh
    HeadOffset As Position
End Type

Private Type HeadData
    head(1 To 8) As Grh
End Type

Private Type tIndiceCabeza
    head(1 To 4) As Integer
End Type

Private Type WeaponAnimData
    WeaponWalk(1 To 8) As Grh
    WeaponAttack As Byte
End Type

Private Type ShieldAnimData
    ShieldWalk(1 To 8) As Grh
End Type

Private Type FxData
    fx As Grh
    offsetX As Long
    OffsetY As Long
End Type

Private NumBodies As Integer
Private Numheads As Integer
Private NumFxs As Integer

Private NumChars As Integer
Private LastChar As Integer
Private NumWeaponAnims As Integer
Private NumShieldAnims As Integer

Private GrhData() As GrhData 'Guarda todos los grh
Private BodyData() As BodyData
Private HeadData() As HeadData
Private FxData() As FxData
Private WeaponAnimData() As WeaponAnimData
Private ShieldAnimData() As ShieldAnimData
Private CascoAnimData() As HeadData

Private Type char
    id As Integer
    Heading As Byte
    Pos As Position
    iHead As Integer
    iBody As Integer
    Body As BodyData
    head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    UsandoArma As Boolean
    fx As Integer
    FxLoopTimes As Integer
    CharType As Byte
    nombre As String
    Pie As Boolean
    Muerto As Boolean
    invisible As Byte
    navigating As Byte
    fly As Byte
    scroll_on As Boolean
    scroll_offset_counter_x As Single
    scroll_offset_counter_y As Single
    scroll_direction_x As Long
    scroll_direction_y As Long
        
    group_index As Integer
        
    particle_count As Integer
    particle_group() As Long
    
    dialog As String
    dialog_color As Long
    dialog_life As Byte
    dialog_font_index As Integer
    
End Type

Private CharList() As char

Private char_current As Integer
Private char_current_overwater As Boolean
Private char_current_onhorse As Boolean
Private engine_label_render As Boolean
Private char_current_blind As Boolean
Private under_roof As Boolean
Private map_base_light As Long
Private Inventory_Render As Boolean
Private meteo_particle As Long
Private dx_event_handle As Long

'***************************
'Constants
'***************************

Public Enum FontAlignment
    fa_center = DT_CENTER Or DT_VCENTER
    fa_top = DT_TOP Or DT_CENTER
    fa_left = DT_LEFT Or DT_VCENTER
    fa_topleft = DT_TOP Or DT_LEFT
    fa_bottomleft = DT_BOTTOM Or DT_LEFT
    fa_bottom = DT_BOTTOM Or DT_CENTER
    fa_right = DT_RIGHT Or DT_VCENTER
    fa_bottomright = DT_BOTTOM Or DT_RIGHT
    fa_topright = DT_TOP Or DT_RIGHT
End Enum

Private Type Particle
    friction As Single
    x As Single
    y As Single
    vector_x As Single
    vector_y As Single
    angle As Single
    Grh As Grh
    alive_counter As Long
    x1 As Integer
    x2 As Integer
    y1 As Integer
    y2 As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Integer
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    rgb_list(0 To 3) As Long
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type

'Modified by: Ryan Cain (Onezero)
'Last modify date: 5/14/2003
Private Type particle_group
    active As Boolean
    id As Long
    map_x As Integer
    map_y As Integer
    char_index As Long

    frame_counter As Single
    frame_speed As Single
    
    stream_type As Byte

    particle_stream() As Particle
    particle_count As Long
    
    grh_index_list() As Long
    grh_index_count As Long
    
    alpha_blend As Boolean
    
    alive_counter As Long
    never_die As Boolean
    
    x1 As Integer
    x2 As Integer
    y1 As Integer
    y2 As Integer
    angle As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Long
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    rgb_list(0 To 3) As Long
    
    'Added by Juan Martín Sotuyo Dodero
    speed As Single
    life_counter As Long
    
    'Added by David Justus
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type

Private Const COLOR_KEY As Long = &HFF000000
Private Const FVF As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

'***************************
'Types
'***************************

Private Type D3D8Textures
    texture As Direct3DTexture8
    Dimension As Long
End Type

'This structure describes a transformed and lit vertex.
Private Type TLVERTEX
    x As Single
    y As Single
    z As Single
    rhw As Single
    color As Long
    specular As Long
    tu As Single
    tv As Single
End Type

'The light structure
Private Type Light
    active As Boolean 'Do we ignore this light?
    id As Long
    map_x As Integer 'Coordinates
    map_y As Integer
    color As Long 'Start colour
    range As Byte
End Type

'Particle system
Dim particle_group_list() As particle_group
Dim particle_group_count As Long
Dim particle_group_last As Long

'***************************
'Variables
'***************************
'Major DX Objects
Dim dX As DirectX8
Dim D3D As Direct3D8
Dim ddevice As Direct3DDevice8
Dim D3dx As D3DX8
Dim d3dpp As D3DPRESENT_PARAMETERS
Dim d3dcaps As D3DCAPS8
Dim d3ddm As D3DDISPLAYMODE
Dim di As DirectInput8
Dim ddevice_mouse As DirectInputDevice8

'Dim ddevice_keyboard As DirectInputDevice8
'Dim key_state As DIKEYBOARDSTATE

'The app path
Dim Path_Resources As String

'Main form handle
Dim form_hwnd As Long

'Display variables
Dim screen_hwnd As Long
Dim screen_width As Integer
Dim screen_height As Integer

'Map view area (where the game is played)
Dim view_screen_top As Integer 'In pixels
Dim view_screen_left As Integer 'In pixels
Dim view_screen_bottom As Integer
Dim view_screen_right As Integer
Dim view_screen_tile_width As Integer 'In tiles
Dim view_screen_tile_height As Integer 'In tiles
Dim view_screen_width As Integer
Dim view_screen_height As Integer

'Buffer area (used to draw object outside the map area but may still show up on the screen)
Dim view_tile_buffer As Integer 'In tiles

'Base tile size (smallest possible tile size: must be square)
Dim base_tile_size As Integer 'In pixels

'View position: In tiles
Dim view_pos_x As Integer
Dim view_pos_y As Integer

'Scrolling stuff
Dim scroll_on As Boolean
Dim scroll_direction_x As Integer
Dim scroll_direction_y As Integer
Dim scroll_offset_counter_x As Single
Dim scroll_offset_counter_y As Single
Dim scroll_pixels_per_frame As Single

'FPS Counters
Dim fps_last_time As Long 'When did we last check the frame rate?
Dim fps_frame_counter As Long 'How many frames have been drawn
Dim fps As Long 'What the current frame rate is.....

'time between frames
Dim timer_elapsed_time As Single
'ticks per frame
Dim timer_ticks_per_frame As Single

'base speed for the engine
Dim engine_base_speed As Single

Dim engine_render_started As Boolean

'windowed or not windowed
Dim engine_windowed As Boolean

'vsync or not vsync
Dim engine_vsync As Boolean

'show engine stats
Dim engine_show_stats As Boolean

'show blocked tiles on map
Dim engine_show_blocked_tiles As Boolean
Dim engine_show_special_tiles As Boolean
Dim engine_show_triggers As Boolean
Dim engine_show_layer_4 As Boolean

'mouse input
Dim input_mouse_button_left As Boolean
Dim input_mouse_button_right As Boolean
Dim input_mouse_moved As Boolean
Dim input_mouse_screen_x As Integer
Dim input_mouse_screen_y As Integer
Dim input_mouse_view_x As Integer
Dim input_mouse_view_y As Integer
Dim input_mouse_map_x As Integer
Dim input_mouse_map_y As Integer
Dim input_mouse_heading As Byte
'Added by Juan Martín Sotuyo Dodero
Dim input_mouse_subtile As Integer
Dim input_mouse_left_click_last As Integer
Dim input_mouse_left_double_click As Boolean
Dim input_mouse_right_click_last As Integer
Dim input_mouse_right_double_click As Boolean
Dim input_mouse_double_click_time As Integer
Dim input_mouse_left_click As Boolean
Dim input_mouse_right_click As Boolean


'***************************
'Arrays
'***************************
'Array of surfaces
'Modified by Augusto José Rando
Dim Textures As CTextureMan

'Light list
Dim light_list() As Light
Dim light_count As Long
Dim light_last As Long

'Font List
Dim font_primitive_list() As StdFont  'This contains all basic data on fonts
Dim font_list() As D3DXFont
Dim font_count As Long
Dim font_last As Long

'Loaded Particle groups list
Dim StreamData() As particle_group
Dim TotalStreams As Long

'Added by Barrin

Private Type MapInfo
    NumUsers As Integer
    Music As String
    Name As String
    Pk As Boolean
    Terreno As String
    Zona As String
    Restringir As String
    BackUp As Byte
    SonidoAmbiente As String
    base_light As Long
End Type

Private MapInfo As MapInfo

'The only current map
Private MapData() As tMapBlock

Private Type tItemData
    GrhIndex As Long
    Type As Integer
    Name As String
End Type

Private Type tNPCData
    Name As String
    BodyIndex As Integer
    HeadIndex As Integer
    Heading As Byte
End Type

'Item data type
Private Item_Data() As tItemData

'NPC data type
Private NPC_Data() As tNPCData

'NPC hostile data type
Private NPC_Hostile_Data() As tNPCData

Private Type WorldPos
    Map As Integer
    x As Integer
    y As Integer
End Type

Private Type Obj
    OBJIndex As Integer
    Amount As Integer
End Type

'Tipo de las celdas del mapa
Private Type tMapBlock
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjInfo As Obj
    ObjGrh As Grh
    NPCIndex As Integer
    TileExit As WorldPos
    Blocked As Byte
    Trigger As Integer
    fx As Integer
    light_value(0 To 3) As Long
    particle_group_index As Long
End Type

Private Type tTileMap
    bloqueado As Byte
    graf1 As Integer
    graf2 As Integer
    graf3 As Integer
    graf4 As Integer
    Trigger As Integer
    t1 As Integer 'espacio al pedo
    l1 As Long 'luz 0
    l2 As Long 'luz 1
    l3 As Long 'luz 2
    l4 As Long 'luz 3
    p1 As Long 'particula
End Type

Private letter_text As String
Private letter_grh As Grh
Private letter_x_pos As Integer
Private letter_y_pos As Integer

'***************************
'External Functions
'***************************
'KeyInput
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Long

'MouseInput
Private Type PointAPI
    x As Long
    y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long

'For getting the display size in windowed mode
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

'Gets number of ticks since windows started
Private Declare Function GetTickCount Lib "kernel32" () As Long

'Old fashion BitBlt function
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
'Added by Juan Martín Sotuyo Dodero
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
'Added by Barrin
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDest As Any, ByRef pSource As Any, ByVal Length As Long)

Private Sub Class_Initialize()
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
End Sub

Private Sub Engine_Stats_Render()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim x As Integer
    Dim y As Integer
    
    Dim offset_x As Integer
    
    offset_x = screen_width - 110
    
    'fps
    Device_Text_Render font_list(1), fps & " FPS", 0, offset_x, 100, 20, &HFFFFFFFF, DT_TOP Or DT_LEFT
    
    'delay
    Device_Text_Render font_list(1), timer_ticks_per_frame & " tpf", 20, offset_x, 100, 20, &HFFFFFFFF, DT_TOP Or DT_LEFT
    
    'cursor screen pos
    Device_Text_Render font_list(1), input_mouse_screen_x & " sx, " & input_mouse_screen_y & " sy", 40, offset_x, 100, 20, &HFFFFFFFF, DT_TOP Or DT_LEFT
    Device_Text_Render font_list(1), input_mouse_view_x & " vx, " & input_mouse_view_y & " vy", 60, offset_x, 100, 20, &HFFFFFFFF, DT_TOP Or DT_LEFT
    Device_Text_Render font_list(1), input_mouse_map_x & " mx, " & input_mouse_map_y & " my", 80, offset_x, 100, 20, &HFFFFFFFF, DT_TOP Or DT_LEFT
    Device_Text_Render font_list(1), input_mouse_heading & " heading", 100, offset_x, 100, 20, &HFFFFFFFF, DT_TOP Or DT_LEFT
    
End Sub

Public Sub Engine_Stats_Show_Toggle()
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Toggles engine stats
'*****************************************************************
    If engine_show_stats Then
        engine_show_stats = False
    Else
        engine_show_stats = True
    End If
End Sub

Public Sub Engine_Special_Tiles_Show_Toggle()
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2003

'*****************************************************************
    If engine_show_special_tiles Then
        engine_show_special_tiles = False
    Else
        engine_show_special_tiles = True
    End If
End Sub

Public Sub Layer_4_Show_Toggle()
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2003

'*****************************************************************
    If engine_show_layer_4 Then
        engine_show_layer_4 = False
    Else
        engine_show_layer_4 = True
    End If
End Sub

Public Sub Engine_Blocked_Tiles_Show_Toggle()
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/17/2003

'*****************************************************************
    If engine_show_blocked_tiles Then
        engine_show_blocked_tiles = False
    Else
        engine_show_blocked_tiles = True
    End If
End Sub

Public Function Engine_Triggers_Show_Toggle()
'*************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modified: 5/27/03
'
'*************************************
    If engine_show_triggers Then
        engine_show_triggers = False
    Else
        engine_show_triggers = True
    End If
End Function

Public Function Engine_View_Pos_Set(ByVal map_x As Integer, ByVal map_y As Integer) As Boolean
'*****************************************************************
'Author: Aaron Perkins - Edited by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'Sets the user postion. If valid, returns True, else False
'*****************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        view_pos_x = map_x
        view_pos_y = map_y
        If letter_text <> "" Then Letter_UnSet
        under_roof = ((Map_Trigger_Get(map_x, map_y) = 1) Or (Map_Trigger_Get(map_x, map_y) = 4))
        Engine_View_Pos_Set = True
    End If
End Function

Public Sub Engine_View_Pos_Get(ByRef map_x As Integer, ByRef map_y As Integer)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the user postion
'*****************************************************************
    map_x = view_pos_x
    map_y = view_pos_y
End Sub

Public Sub Engine_Base_Speed_Set(ByVal speed As Single)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'*****************************************************************
    engine_base_speed = speed
End Sub

Public Sub Engine_Label_Render_Set()
'*****************************************************************
'Author: Augusto José Rando
'Last Modify Date: 6/11/2002
'*****************************************************************
    engine_label_render = Not engine_label_render
End Sub

Public Function Engine_Label_Render_Get() As Boolean
'*****************************************************************
'Author: Augusto José Rando
'Last Modify Date: 6/11/2002
'*****************************************************************
    Engine_Label_Render_Get = engine_label_render
End Function

Public Sub Engine_Inventory_Render_Set()
'*****************************************************************
'Author: Augusto José Rando
'Last Modify Date: 6/11/2002
'*****************************************************************
    Inventory_Render = True
End Sub

Public Function Engine_Meteo_Particle_Get() As Long
'*****************************************************************
'Author: Augusto José Rando
'Last Modify Date: 6/11/2002
'*****************************************************************
    Engine_Meteo_Particle_Get = meteo_particle
End Function

Public Function Engine_DirectX_Event_Handle_Get() As Long
'*****************************************************************
'Author: Augusto José Rando
'Last Modify Date: 6/11/2002
'*****************************************************************
    Engine_DirectX_Event_Handle_Get = dx_event_handle
End Function

Public Sub Engine_Meteo_Particle_Set(ByVal meteo_part As Long)
'*****************************************************************
'Author: Augusto José Rando
'Last Modify Date: 6/11/2002
'*****************************************************************
    
    If (meteo_part = -1) And (meteo_particle <> 0) Then
        Call Particle_Group_Remove(meteo_particle)
    ElseIf (meteo_part <> -1) Then
        If meteo_particle <> 0 Then Call Particle_Group_Remove(meteo_particle)
        meteo_particle = General_Particle_Create(meteo_part, -1, -1)
    End If
    
End Sub

Public Function Engine_Inventory_Render_Get() As Boolean
'*****************************************************************
'Author: Augusto José Rando
'Last Modify Date: 6/11/2002
'*****************************************************************
    Engine_Inventory_Render_Get = Inventory_Render
End Function

Public Function Engine_Base_Speed_Get() As Single
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the target fps that the engine
'*****************************************************************
    Engine_Base_Speed_Get = engine_base_speed
End Function

Public Function Engine_Scroll_Pixels_Get() As Single
'*****************************************************************
'Author: Augusto José Rando
'Last Modify Date: 7/15/2005
'*****************************************************************
    Engine_Scroll_Pixels_Get = scroll_pixels_per_frame
End Function

Public Function Engine_Scroll_Pixels_Set(ByVal scroll_pixels As Single)
'*****************************************************************
'Author: Augusto José Rando
'Last Modify Date: 7/15/2005
'*****************************************************************
    scroll_pixels_per_frame = scroll_pixels
End Function

Public Function Engine_FPS_Get() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the target fps that the engine
'*****************************************************************
    Engine_FPS_Get = fps
End Function

Public Sub Engine_DeInitialize()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 3/28/2004
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
On Error Resume Next
    
    Dim i As Long
    
    'Set no Textures to standard stage to avoid memory leak
    ddevice.SetTexture 0, Nothing
    
    'Destroy all Texturess
    Call Textures.BorrarTodo
    Set Textures = Nothing
    
    'Destroy all fonts
    For i = 1 To UBound(font_list)
        Set font_list(i) = Nothing
    Next i
    
    For i = 1 To UBound(font_primitive_list)
        Set font_primitive_list(i) = Nothing
    Next i
    
    'ddevice_mouse.Unacquire
    Set ddevice_mouse = Nothing
    Set di = Nothing
    
    Set D3dx = Nothing
    Set ddevice = Nothing
    Set D3D = Nothing
    Set dX = Nothing

End Sub

Public Function Engine_Initialize(ByVal f_hwnd As Long, ByVal s_hwnd As Long, ByVal windowed As Boolean, _
                                    ByVal r_path As String, Optional ByVal s_width As Integer, Optional ByVal s_height As Integer, _
                                    Optional ByVal v_left As Integer = 0, Optional ByVal v_top As Integer = 0, Optional ByVal v_width_in_tiles As Integer = 0, _
                                    Optional ByVal v_height_in_tiles As Integer = 0, Optional ByVal tile_size As Integer = 32, Optional grh_files_need_unlocking As Boolean = False, _
                                    Optional delete_grh_files_after_load As Boolean = False, Optional vsync_enable As Boolean = True, Optional ByVal DevIndex As Long = D3DADAPTER_DEFAULT _
                                    ) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 7/12/2004
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
On Local Error GoTo ErrorHandler

    Dim flags As Long
    Dim ValidFormat As Boolean

    '****************************************************
    'Fill in global variables
    '****************************************************
    Engine_Initialize = True
    
    'Main display
    screen_hwnd = s_hwnd
    form_hwnd = f_hwnd
    
    fps = 1000
    
    'App path
    Path_Resources = r_path
    
    'Fill in base tile size (must always be square)
    base_tile_size = tile_size 'In pixels
    
    'Fill in viewport sizes. How large the game area is.
    If v_width_in_tiles = 0 Or v_height_in_tiles = 0 Then
        view_screen_tile_width = 17 'In tiles
        view_screen_tile_height = 13 'In tiles
    Else
        view_screen_tile_width = v_width_in_tiles 'In tiles
        view_screen_tile_height = v_height_in_tiles 'In tiles
    End If
    
    If windowed Then
        Dim target As RECT
        GetWindowRect screen_hwnd, target
        view_screen_top = target.top
        view_screen_left = target.left
        view_screen_right = target.Right
        view_screen_bottom = target.bottom
        view_screen_width = target.Right - target.left
        view_screen_height = target.bottom - target.top
        screen_width = view_screen_width
        screen_height = view_screen_height
        engine_windowed = True
    Else
        screen_width = s_width
        screen_height = s_height
        view_screen_left = v_left 'In pixels
        view_screen_top = v_top 'In pixels
        view_screen_width = view_screen_tile_width * base_tile_size
        view_screen_height = view_screen_tile_height * base_tile_size
        view_screen_right = view_screen_left + view_screen_width - 1
        view_screen_bottom = view_screen_top + view_screen_height - 1
        engine_windowed = False
    End If
        
    '****************************************************
    'Setup Map
    '****************************************************
    'Buffer area
    view_tile_buffer = 8
    
    'How many pixels to move per frame when scrolling
    scroll_pixels_per_frame = 4.5
    
    'Added by Augusto José Rando
    engine_vsync = vsync_enable
    
    'User start position
    view_pos_x = 1
    view_pos_y = 1
        
    '*******************************
    'Initialize root DirectX8 objects
    '*******************************
    Set dX = New DirectX8
    'Create the Direct3D object
    Set D3D = dX.Direct3DCreate
    'Create helper class
    Set D3dx = New D3DX8
    'Barrin: Textures manager
    Set Textures = New CTextureMan
    'Input object
    Set di = dX.DirectInputCreate
    
    '*******************************
    'Initialize video device
    '*******************************
    Dim DevType As CONST_D3DDEVTYPE
    DevType = D3DDEVTYPE_HAL
    
    'Get the capabilities of the Direct3D device that we specify. In this case,
    'we'll be using the adapter default (the primiary card on the system).
    
    'Modified by Augusto José Rando
    'Now uses a specified device in setup
    
    On Local Error Resume Next
    
    Call D3D.GetDeviceCaps(DevIndex, DevType, d3dcaps)
    
    If Err.number <> 0 Then
        If MsgBox("El engine no ha encontrado un dispositivo acelerador compatible con Direct3D 8.1 - se utilizará un dispositivo de referencia. Probablemente habrá problemas, ¿desea continuar?", vbExclamation + vbYesNo, "Inicialización Direct3D") = vbYes Then
            DevType = D3DDEVTYPE_REF
            Call D3D.GetDeviceCaps(DevIndex, DevType, d3dcaps)
            Err.Clear
        Else
            GoTo ErrorHandler
        End If
    End If
    
    On Local Error GoTo ErrorHandler
        
    'Grab some information about the current display mode.
    Call D3D.GetAdapterDisplayMode(DevIndex, d3ddm)
    'Now we'll go ahead and fill the D3DPRESENT_PARAMETERS type.
        
    'If d3dcaps.Caps3 = int(0x00000020L) Then Debug.Print "TRUE"
    'Debug.Print D3DCAPS3_ALPHA_FULLSCREEN_FLIP_OR_DISCARD
    
    With d3dpp
        If engine_windowed Then
            .windowed = 1
            
            If engine_vsync Then
                .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
            Else
                .SwapEffect = D3DSWAPEFFECT_COPY
            End If
                        
            .BackBufferFormat = d3ddm.format 'current display depth
        Else
            .SwapEffect = D3DSWAPEFFECT_FLIP
            'Set refresh rate
            .FullScreen_RefreshRateInHz = D3DPRESENT_RATE_DEFAULT
            'Turn off vsync
            .FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_IMMEDIATE
            'set color depth
            .BackBufferFormat = D3DFMT_R5G6B5 '16 bit
            'Back buffer size
            .BackBufferWidth = screen_width
            .BackBufferHeight = screen_height
        End If
    End With
        
    On Local Error Resume Next
    
    ' Confirm the device for HW vertex processing
    flags = D3DCREATE_HARDWARE_VERTEXPROCESSING
    If d3dcaps.DevCaps And d3dcaps.VertexProcessingCaps Then
       If VerifyDevice(flags, d3dpp.BackBufferFormat) Then ValidFormat = True
    End If
    
    'Try Software VertexProcesing
    If Not ValidFormat Then
        flags = D3DCREATE_SOFTWARE_VERTEXPROCESSING
        If VerifyDevice(flags, d3dpp.BackBufferFormat) Then ValidFormat = True
    End If
    
    'Fail
    If (Not ValidFormat) Or (Err.number <> 0) Then
        flags = D3DCREATE_SOFTWARE_VERTEXPROCESSING
        Err.Clear
    End If
            
    'create device
    Set ddevice = D3D.CreateDevice(DevIndex, DevType, screen_hwnd, flags, d3dpp)
    
    If (ddevice Is Nothing) Or (Err.number <> 0) Then
        flags = D3DCREATE_SOFTWARE_VERTEXPROCESSING
        Err.Clear
        Set ddevice = D3D.CreateDevice(DevIndex, DevType, screen_hwnd, flags, d3dpp)
    End If
        
    On Error GoTo ErrorHandler
    
    'setup device
    With ddevice
        'Set the vertex shader to an FVF that contains Textures coords,
        'and transformed and lit vertex coords.
        Call .SetVertexShader(FVF)
        'Turn off lighting
        Call .SetRenderState(D3DRS_LIGHTING, 0)
        'Set the render state that uses the alpha component as the source for blending.
        Call .SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
        'Set the render state that uses the inverse alpha component as the destination blend.
        Call .SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
        'Enable alpha blending
        Call .SetRenderState(D3DRS_ALPHABLENDENABLE, 1)
    End With
    
    MinXBorder = XMinMapSize + (view_screen_tile_width \ 2)
    MaxXBorder = XMaxMapSize - (view_screen_tile_width \ 2)
    MinYBorder = YMinMapSize + (view_screen_tile_height \ 2)
    MaxYBorder = YMaxMapSize - (view_screen_tile_height \ 2)

    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As tMapBlock
    
    'Added by Augusto José Rando
    Call Grh_Load_All
    Call Grh_Load_Body
    Call Grh_Load_Heads
    Call Grh_Load_Helmets
    Call Grh_Load_Shields
    Call Grh_Load_Weapons
    Call Grh_Load_Fx
    
    Call Engine_Font_Initialize
    
    Textures.Path = App.Path & "\Recursos"
    Call Textures.Init(D3dx, ddevice, General_Get_Free_Ram_Bytes, 700)
    
    '*******************************
    'Initialize input devices
    '*******************************
    Set ddevice_mouse = di.CreateDevice("guid_SysMouse")
    'Set ddevice_keyboard = di.CreateDevice("GUID_SysKeyboard")
    
    Call ddevice_mouse.SetCommonDataFormat(DIFORMAT_MOUSE)
    Call ddevice_mouse.SetCooperativeLevel(form_hwnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE)
    
    'Call ddevice_keyboard.SetCommonDataFormat(DIFORMAT_KEYBOARD)
    'Call ddevice_keyboard.SetCooperativeLevel(form_hwnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE)
    
    'dx_event_handle = dX.CreateEvent(frmMain)
    'ddevice_keyboard.SetEventNotification dx_event_handle
    
    'Set the buffer size
    Dim diProp As DIPROPLONG
    diProp.lHow = DIPH_DEVICE
    diProp.lObj = 0
    diProp.lData = 1000
    Call ddevice_mouse.SetProperty("DIPROP_BUFFERSIZE", diProp)
    
    ddevice_mouse.Acquire
    'ddevice_keyboard.Acquire
    
    '****************************************************
    'Misc stuff
    '****************************************************
    'Set up default font
    Font_Create "Arial", 11, False, False
    
    'Clears the buffer
    Device_Clear
    
Exit Function

ErrorHandler:
    MsgBox "Error al iniciar el engine: " & Err.number & ": " & Err.Description, vbCritical, "Inicialización de Direct3D"
    Engine_Initialize = False

End Function

Public Function Engine_View_Move(ByVal Heading As Byte) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim x As Integer
    Dim y As Integer
    Dim nX As Integer
    Dim nY As Integer
    
    'Remove letter text if neccesary
    If letter_text <> "" Then Letter_UnSet
    
    'Don't move if we are already moving....
    If scroll_on Then
        Engine_View_Move = False
        Exit Function
    End If
    
    'Invalid heading
    If Heading < 1 Or Heading > 8 Then
        Engine_View_Move = False
        Exit Function
    End If
    
    x = view_pos_x
    y = view_pos_y
    nX = x
    nY = y
    Convert_Heading_to_Direction Heading, nX, nY
    
    'See if out new position is legal
    If Map_In_Bounds(nX, nY) Then
        'start the scrolling process
        view_pos_x = nX
        view_pos_y = nY
                
        'bTecho = IIf(MapData(CurrentUser.X, CurrentUser.Y).Trigger = 1 Or _
        '        MapData(CurrentUser.X, CurrentUser.Y).Trigger = 2 Or _
        '        MapData(CurrentUser.X, CurrentUser.Y).Trigger = 4, True, False)
        
        scroll_offset_counter_x = (base_tile_size * (x - nX))
        scroll_offset_counter_y = (base_tile_size * (y - nY))
        scroll_direction_x = nX - x
        scroll_direction_y = nY - y
        scroll_on = True
        
        'See if mouse is in view
        If Input_Mouse_In_View Then
            'Update pos
            Convert_View_To_Map input_mouse_view_x, input_mouse_view_y, input_mouse_map_x, input_mouse_map_y
        End If
        
        Engine_View_Move = True
    Else
        'not legal don't move
        scroll_direction_x = 0
        scroll_direction_y = 0
    End If
End Function

Public Sub Engine_Render_Mini_Map_To_hDC(ByVal desthDC As Long)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/17/2004
'Renders the mini map to the given hDC
'**************************************************************
    Dim map_x As Long
    Dim map_y As Long
    Dim screen_x As Long
    Dim screen_y As Long
    Dim grh_index As Long
    
    '*********************
    'Layer 1 (Floor level) and walls
    '*********************
    For map_y = XMinMapSize To XMaxMapSize
        For map_x = YMinMapSize To YMaxMapSize
            screen_x = (map_x - 1) * 2
            screen_y = (map_y - 1) * 2
            '*** Start Layer 1 ***
            If MapData(map_x, map_y).Graphic(1).GrhIndex Then
                grh_index = MapData(map_x, map_y).Graphic(1).GrhIndex
                SetPixel desthDC, map_x, map_y, GrhData(grh_index).mini_map_color
            End If
            '*** End Layer 1 ***
        Next map_x
    Next map_y
    
    For map_y = XMinMapSize To XMaxMapSize
        For map_x = YMinMapSize To YMaxMapSize
            screen_x = (map_x - 1) * 2
            screen_y = (map_y - 1) * 2
            '*** Start Layer 2 ***
            If MapData(map_x, map_y).Graphic(2).GrhIndex Then
                grh_index = MapData(map_x, map_y).Graphic(2).GrhIndex
                SetPixel desthDC, map_x, map_y, GrhData(grh_index).mini_map_color
            End If
            '*** End Layer 2 ***
        Next map_x
    Next map_y
End Sub

Public Function Engine_Render_Start() As Boolean

'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/29/2003
'Edited by Juan Martín Sotuyo Dodero
'**************************************************************
On Error Resume Next
    If engine_render_started Then Exit Function
    
    Engine_Render_Start = True
    
    '****************************************************
    'Handle Alt-Tab
    '****************************************************
    'Test Cooperative Level
    If Not Device_Test_Cooperative_Level Then Exit Function
    
    '****************************************************
    'Render
    '****************************************************
    
    '*******************************
    'get the screen_rect if windowed
    If engine_windowed Then
        Dim target As RECT
        GetWindowRect screen_hwnd, target
        view_screen_top = target.top
        view_screen_left = target.left
        view_screen_right = target.Right
        view_screen_bottom = target.bottom
        view_screen_width = target.Right - target.left
        view_screen_height = target.bottom - target.top
    End If
    '*******************************

    '*******************************
    'Erase the backbuffer so that it can be drawn on again
    Device_Clear
    '*******************************
    
    '*******************************
    'Start the scene
    ddevice.BeginScene
    '*******************************
    
    '*******************************
    'Render lights
    Light_Render_All
    '*******************************
        
    '*******************************
    'Draw Map
    Map_Render
    '*******************************
    
    engine_render_started = True
Exit Function

ErrorHandler:
    Engine_Render_Start = False
    MsgBox "Error in Engine_Render_Start: " & Err.number & ": " & Err.Description

End Function

Public Function Engine_Render_End() As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/29/2003
'Edited by Juan Martín Sotuyo Dodero
'**************************************************************
On Error Resume Next
    
'Static ticks_elapsed As Single
Dim i As Integer

    Engine_Render_End = True

    If Not engine_render_started Then Exit Function

    '*******************************
    'Draw engine stats
    If engine_show_stats Then
        Engine_Stats_Render
    End If
    '*******************************
    
    '*******************************
    'End scene
    ddevice.EndScene
    '*******************************
    
    '*******************************
    'Flip the backbuffer to the screen
    Device_Flip
    '*******************************

    '*******************************
    'Calculate current frames per second
    If GetTickCount >= (fps_last_time + 1000) Then
        fps = fps_frame_counter
        fps_frame_counter = 0
        'ticks_elapsed = 1000 / fps
        fps_last_time = GetTickCount
        
        engine_erase_graphics = engine_erase_graphics + 1
        
        If engine_erase_graphics >= 60 Then
            Textures.BorraTimer
            engine_erase_graphics = 0
        End If
        
        For i = 1 To LastChar
            If Char_Check(i) Then
                If CharList(i).dialog <> "" Then
                    CharList(i).dialog_life = CharList(i).dialog_life - 1
                    If CharList(i).dialog_life <= 0 Then Char_Dialog_Remove (i)
                End If
            End If
        Next i
        
    Else
        fps_frame_counter = fps_frame_counter + 1
    End If
    '*******************************
        
    'Get timing info
    timer_elapsed_time = General_Get_Elapsed_Time()
    timer_ticks_per_frame = (timer_elapsed_time * engine_base_speed)
    
    '*******************************
    'Poll Inputs
    Input_Mouse_Poll
    '*******************************
    
    'If timer_ticks_per_frame < ticks_elapsed Then
    '    General_Sleep ticks_elapsed - timer_ticks_per_frame
    'End If
    
    engine_render_started = False

Exit Function

ErrorHandler:
    Engine_Render_End = False
    MsgBox "Error in Engine_Render_End: " & Err.number & ": " & Err.Description

End Function

Private Sub Device_Flip()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    'Draw the graphics to the front buffer.
    ddevice.Present ByVal 0&, ByVal 0&, 0, ByVal 0&
End Sub

Private Sub Device_Box_Textured_Render_Advance(ByVal GrhIndex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, _
                                            ByVal src_height As Integer, ByRef rgb_list() As Long, ByVal src_x As Integer, _
                                            ByVal src_y As Integer, ByVal dest_width As Integer, Optional ByVal dest_height As Integer, _
                                            Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/15/2003
'Copies the Textures allowing resizing
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
    Static src_rect As RECT
    Static dest_rect As RECT
    Static temp_verts(3) As TLVERTEX
    Static d3dTextures As D3D8Textures
    Static light_value(0 To 3) As Long

    If GrhIndex = 0 Then Exit Sub
    Set d3dTextures.texture = Textures.GetTexture(GrhData(GrhIndex).FileNum, d3dTextures.Dimension)
    
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)
        
    If Not char_current_blind Then
        If (light_value(0) = 0) Then light_value(0) = map_base_light
        If (light_value(1) = 0) Then light_value(1) = map_base_light
        If (light_value(2) = 0) Then light_value(2) = map_base_light
        If (light_value(3) = 0) Then light_value(3) = map_base_light
    Else
        light_value(0) = blind_color
        light_value(1) = blind_color
        light_value(2) = blind_color
        light_value(3) = blind_color
    End If
    
    'Set up the source rectangle
    With src_rect
        .bottom = src_y + src_height
        .left = src_x
        .Right = src_x + src_width
        .top = src_y
    End With
        
    'Set up the destination rectangle
    With dest_rect
        .bottom = dest_y + dest_height
        .left = dest_x
        .Right = dest_x + dest_width
        .top = dest_y
    End With
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), d3dTextures.Dimension, angle
        
    'Set Textures
    ddevice.SetTexture 0, d3dTextures.texture
    
    If alpha_blend Then
       'Set Rendering for alphablending
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    'Draw the triangles that make up our square Textures
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If

End Sub

Private Sub Device_Box_Textured_Render(ByVal GrhIndex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, _
                                            ByVal src_height As Integer, ByRef rgb_list() As Long, ByVal src_x As Integer, _
                                            ByVal src_y As Integer, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/12/2004
'Just copies the Textures
'**************************************************************
    Static src_rect As RECT
    Static dest_rect As RECT
    Static temp_verts(3) As TLVERTEX
    Static d3dTextures As D3D8Textures
    Static light_value(0 To 3) As Long
    
    If GrhIndex = 0 Then Exit Sub
    Set d3dTextures.texture = Textures.GetTexture(GrhData(GrhIndex).FileNum, d3dTextures.Dimension)
    
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)
    
    If Not char_current_blind Then
        If (light_value(0) = 0) Then light_value(0) = map_base_light
        If (light_value(1) = 0) Then light_value(1) = map_base_light
        If (light_value(2) = 0) Then light_value(2) = map_base_light
        If (light_value(3) = 0) Then light_value(3) = map_base_light
    Else
        light_value(0) = blind_color
        light_value(1) = blind_color
        light_value(2) = blind_color
        light_value(3) = blind_color
    End If
        
    'Set up the source rectangle
    With src_rect
        .bottom = src_y + src_height
        .left = src_x
        .Right = src_x + src_width
        .top = src_y
    End With
                
    'Set up the destination rectangle
    With dest_rect
        .bottom = dest_y + src_height
        .left = dest_x
        .Right = dest_x + src_width
        .top = dest_y
    End With
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), d3dTextures.Dimension, angle
    
    'Set Textures
    ddevice.SetTexture 0, d3dTextures.texture
    
    If alpha_blend Then
       'Set Rendering for alphablending
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    'Draw the triangles that make up our square Textures
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If

End Sub

Private Sub Device_Clear()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    'Clear the back buffer
    ddevice.Clear 0, ByVal 0&, D3DCLEAR_TARGET, 0, 0, 0
End Sub

Private Sub Device_Reset()
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'Resets the device
'**************************************************************
On Error GoTo ErrHandler
    
    'Destroy all fonts
    Dim loopc As Long
    For loopc = 1 To font_last
        Set font_list(loopc) = Nothing
    Next loopc
    
    font_last = 0
    
    'Reset the device
    ddevice.Reset d3dpp
    
    'Reset the device's rendering state
    With ddevice
        'Set the vertex shader to an FVF that contains Textures coords,
        'and transformed and lit vertex coords.
        Call .SetVertexShader(FVF)
        
        'Turn off lighting
        Call .SetRenderState(D3DRS_LIGHTING, 0)
                                    
        'Set the render state that uses the alpha component as the source for blending.
        Call .SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
            
        'Set the render state that uses the inverse alpha component as the destination blend.
        Call .SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
        
        'Enable alpha blending
        Call .SetRenderState(D3DRS_ALPHABLENDENABLE, 1)
        
    End With
    
    'Recreate all fonts from their primitives
    For loopc = 1 To UBound(font_primitive_list())
        Font_Create font_primitive_list(loopc).Name, font_primitive_list(loopc).Size, font_primitive_list(loopc).bold, font_primitive_list(loopc).italic
    Next loopc
Exit Sub

ErrHandler:
End Sub

Private Sub Device_Text_Render(ByVal font As D3DXFont, ByVal Text As String, ByVal top As Integer, ByVal left As Integer, _
                                ByVal Width As Integer, ByVal Height As Integer, ByVal color As Long, ByVal format As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim TextRect As RECT 'This defines where it will be
    
    TextRect.top = top
    TextRect.left = left
    TextRect.bottom = top + Height
    TextRect.Right = left + Width
    D3dx.DrawText font, color, Text, TextRect, format
End Sub

Private Function Geometry_Create_TLVertex(ByVal x As Single, ByVal y As Single, ByVal z As Single, _
                                            ByVal rhw As Single, ByVal color As Long, ByVal specular As Long, tu As Single, _
                                            ByVal tv As Single) As TLVERTEX
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'**************************************************************
    Geometry_Create_TLVertex.x = x
    Geometry_Create_TLVertex.y = y
    Geometry_Create_TLVertex.z = z
    Geometry_Create_TLVertex.rhw = rhw
    Geometry_Create_TLVertex.color = color
    Geometry_Create_TLVertex.specular = specular
    Geometry_Create_TLVertex.tu = tu
    Geometry_Create_TLVertex.tv = tv
End Function

Private Sub Geometry_Create_Box(ByRef verts() As TLVERTEX, ByRef dest As RECT, ByRef src As RECT, ByRef rgb_list() As Long, _
                                Optional ByRef Textures_size As Long, Optional ByVal angle As Single)
'**************************************************************
'Author: Aaron Perkins
'Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 11/17/2002
'
' * v1      * v3
' |\        |
' |  \      |
' |    \    |
' |      \  |
' |        \|
' * v0      * v2
'**************************************************************
    Dim x_center As Single
    Dim y_center As Single
    Dim radius As Single
    Dim x_Cor As Single
    Dim y_Cor As Single
    Dim left_point As Single
    Dim right_point As Single
    Dim temp As Single
    
    If angle > 0 Then
        'Center coordinates on screen of the square
        x_center = dest.left + (dest.Right - dest.left) / 2
        y_center = dest.top + (dest.bottom - dest.top) / 2
        
        'Calculate radius
        radius = Sqr((dest.Right - x_center) ^ 2 + (dest.bottom - y_center) ^ 2)
        
        'Calculate left and right points
        temp = (dest.Right - x_center) / radius
        right_point = Atn(temp / Sqr(-temp * temp + 1))
        left_point = PI - right_point
    End If
    
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.left
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-left_point - angle) * radius
        y_Cor = y_center - Sin(-left_point - angle) * radius
    End If
    
    
    '0 - Bottom left vertex
    If Textures_size Then
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.left / Textures_size, (src.bottom + 1) / Textures_size)
    Else
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.left
        y_Cor = dest.top
    Else
        x_Cor = x_center + Cos(left_point - angle) * radius
        y_Cor = y_center - Sin(left_point - angle) * radius
    End If
    
    
    '1 - Top left vertex
    If Textures_size Then
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, src.left / Textures_size, src.top / Textures_size)
    Else
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-right_point - angle) * radius
        y_Cor = y_center - Sin(-right_point - angle) * radius
    End If
    
    
    '2 - Bottom right vertex
    If Textures_size Then
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, (src.Right + 1) / Textures_size, (src.bottom + 1) / Textures_size)
    Else
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.top
    Else
        x_Cor = x_center + Cos(right_point - angle) * radius
        y_Cor = y_center - Sin(right_point - angle) * radius
    End If
    
    
    '3 - Top right vertex
    If Textures_size Then
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.Right + 1) / Textures_size, src.top / Textures_size)
    Else
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 0, 0)
    End If

End Sub

Public Sub Grh_Render_To_Hdc(ByVal grh_index As Long, desthDC As Long, ByVal screen_x As Integer, ByVal screen_y As Integer, Optional transparent As Boolean = False)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 8/30/2004
'This method is SLOW... Don't use in a loop if you care about
'speed!
'Modified by Juan Martín Sotuyo Dodero
'*************************************************************
    
    On Error GoTo ErrorHandler
    
    Dim file_path As String
    Dim src_x As Integer
    Dim src_y As Integer
    Dim src_width As Integer
    Dim src_height As Integer
    Dim hdcsrc As Long
    Dim MaskDC As Long
    Dim PrevObj As Long
    Dim PrevObj2 As Long

    If grh_index <= 0 Then Exit Sub

    'If it's animated switch grh_index to first frame
    If GrhData(grh_index).NumFrames <> 1 Then
        grh_index = GrhData(grh_index).Frames(1)
    End If

    If Extract_File(Graphics, App.Path & "\Recursos", GrhData(grh_index).FileNum & ".bmp", Windows_Temp_Dir, False) Then
        file_path = Windows_Temp_Dir & GrhData(grh_index).FileNum & ".bmp"
        
        src_x = GrhData(grh_index).sX
        src_y = GrhData(grh_index).sY
        src_width = GrhData(grh_index).pixelWidth
        src_height = GrhData(grh_index).pixelHeight
            
        hdcsrc = CreateCompatibleDC(desthDC)
        PrevObj = SelectObject(hdcsrc, LoadPicture(file_path))
        
        If transparent = False Then
            BitBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcCopy
        Else
            MaskDC = CreateCompatibleDC(desthDC)
            
            PrevObj2 = SelectObject(MaskDC, LoadPicture(file_path))
            
            Grh_Create_Mask hdcsrc, MaskDC, src_x, src_y, src_width, src_height
            
            'Render tranparently
            BitBlt desthDC, screen_x, screen_y, src_width, src_height, MaskDC, src_x, src_y, vbSrcAnd
            BitBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcPaint
            
            Call DeleteObject(SelectObject(MaskDC, PrevObj2))
            
            DeleteDC MaskDC
        End If
        
        Call DeleteObject(SelectObject(hdcsrc, PrevObj))
        DeleteDC hdcsrc
        
        Call Delete_File(Windows_Temp_Dir & GrhData(grh_index).FileNum & ".bmp")
    End If
    
    Exit Sub
    
ErrorHandler:
    If General_File_Exists(Windows_Temp_Dir & GrhData(grh_index).FileNum & ".bmp", vbNormal) Then
        Call Delete_File(Windows_Temp_Dir & GrhData(grh_index).FileNum & ".bmp")
    End If
    
End Sub

Public Sub Grh_Load(ByVal file_num As Integer)
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 7/16/2005
'We don't return nothing
'**************************************************************
Call Textures.CargarGrafico(file_num)
    
End Sub

Private Sub Grh_Render(ByRef Grh As Grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByRef rgb_list() As Long, Optional ByVal h_centered As Boolean = True, Optional ByVal v_centered As Boolean = True, Optional ByVal alpha_blend As Boolean = False)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'Modified by Juan Martín Sotuyo Dodero
'Added centering
'**************************************************************
    Dim tile_width As Integer
    Dim tile_height As Integer
    Dim grh_index As Long
    
    If Grh.GrhIndex = 0 Then Exit Sub
        
    'Animation
    If Grh.started Then
        Grh.framecounter = Grh.framecounter + (timer_ticks_per_frame * Grh.speed)
        If Grh.framecounter > GrhData(Grh.GrhIndex).NumFrames Then
            'If Grh.noloop Then
            '    Grh.FrameCounter = GrhData(Grh.GrhIndex).NumFrames
            'Else
                Grh.framecounter = 1
            'End If
        End If
    End If

        'particle_group_list(particle_group_index).frame_counter = particle_group_list(particle_group_index).frame_counter + timer_ticks_per_frame
        'If particle_group_list(particle_group_index).frame_counter > particle_group_list(particle_group_index).frame_speed Then
        '    particle_group_list(particle_group_index).frame_counter = 0
        '    no_move = False
        'Else
        '    no_move = True
        'End If

    'Figure out what frame to draw (always 1 if not animated)
    If Grh.framecounter = 0 Then Grh.framecounter = 1
    If Not Grh_Check(Grh.GrhIndex) Then Exit Sub
    grh_index = GrhData(Grh.GrhIndex).Frames(Grh.framecounter)
    If grh_index <= 0 Then Exit Sub
    If GrhData(grh_index).FileNum = 0 Then Exit Sub
        
    'Modified by Augusto José Rando
    'Simplier function - according to basic ORE engine
    If h_centered Then
        If GrhData(Grh.GrhIndex).TileWidth <> 1 Then
            screen_x = screen_x - Int(GrhData(Grh.GrhIndex).TileWidth * (base_tile_size \ 2)) + base_tile_size \ 2
        End If
    End If
    
    If v_centered Then
        If GrhData(Grh.GrhIndex).TileHeight <> 1 Then
            screen_y = screen_y - Int(GrhData(Grh.GrhIndex).TileHeight * base_tile_size) + base_tile_size
        End If
    End If
    
    'Draw it to device
    Device_Box_Textured_Render grh_index, _
        screen_x, screen_y, _
        GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, _
        rgb_list(), _
        GrhData(grh_index).sX, GrhData(grh_index).sY, _
        alpha_blend, _
        Grh.angle

End Sub

Private Sub Grh_Render_Char_FX(ByRef Grh As Grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByRef rgb_list() As Long, ByRef temp_char As char, Optional ByVal h_centered As Boolean, Optional ByVal v_centered As Boolean, Optional ByVal alpha_blend As Boolean = False)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'Modified by Augusto José Rando
'FX Char-Rendering
'**************************************************************
    Dim tile_width As Integer
    Dim tile_height As Integer
    Dim grh_index As Long
    
    'Animation
    If Grh.started Then
        Grh.framecounter = Grh.framecounter + (timer_ticks_per_frame * Grh.speed)
        If Grh.framecounter > GrhData(Grh.GrhIndex).NumFrames Then
            If temp_char.FxLoopTimes > 1 Then
                Grh.framecounter = 1
                temp_char.FxLoopTimes = temp_char.FxLoopTimes - 1
            Else
                Grh.framecounter = 1
                temp_char.fx = 0
                temp_char.FxLoopTimes = 0
                Exit Sub
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    If Grh.framecounter = 0 Then Grh.framecounter = 1
    grh_index = GrhData(Grh.GrhIndex).Frames(Grh.framecounter)
    If GrhData(grh_index).FileNum = 0 Then Exit Sub
        
    'Center Grh over X, Y pos
    If GrhData(Grh.GrhIndex).TileWidth <> 1 Then
        screen_x = screen_x - Int(GrhData(Grh.GrhIndex).TileWidth * (base_tile_size \ 2)) + base_tile_size \ 2
    End If
    
    If GrhData(Grh.GrhIndex).TileHeight <> 1 Then
        screen_y = screen_y - Int(GrhData(Grh.GrhIndex).TileHeight * base_tile_size) + base_tile_size
    End If
        
    'Draw it to device
    Device_Box_Textured_Render grh_index, _
        screen_x, screen_y, _
        GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, _
        rgb_list, _
        GrhData(grh_index).sX, GrhData(grh_index).sY, _
        alpha_blend, _
        Grh.angle
End Sub

Private Sub Grh_Render_MapFX(ByRef Grh As Grh, ByVal map_x As Integer, ByVal map_y As Integer, ByVal screen_x As Integer, ByVal screen_y As Integer, ByRef rgb_list() As Long, Optional ByVal alpha_blend As Boolean = False)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'Modified by Juan Martín Sotuyo Dodero
'Added centering
'**************************************************************
    Dim tile_width As Integer
    Dim tile_height As Integer
    Dim grh_index As Long
        
    'Animation
    If Grh.started Then
        Grh.framecounter = Grh.framecounter + (timer_ticks_per_frame * Grh.speed)
        If Grh.framecounter > GrhData(Grh.GrhIndex).NumFrames Then
            'If Grh.noloop Then
            '    Grh.FrameCounter = GrhData(Grh.GrhIndex).NumFrames
            'Else
                Grh.framecounter = 1
                MapData(map_x, map_y).fx = 0
                Exit Sub
            'End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    If Grh.framecounter = 0 Then Grh.framecounter = 1
    grh_index = GrhData(Grh.GrhIndex).Frames(Grh.framecounter)
    If GrhData(grh_index).FileNum = 0 Then Exit Sub
        
    If GrhData(Grh.GrhIndex).TileWidth <> 1 Then
        screen_x = screen_x - Int(GrhData(Grh.GrhIndex).TileWidth * (base_tile_size \ 2)) + base_tile_size \ 2
    End If
    
    If GrhData(Grh.GrhIndex).TileHeight <> 1 Then
        screen_y = screen_y - Int(GrhData(Grh.GrhIndex).TileHeight * (base_tile_size)) + base_tile_size
    End If
        
    'Draw it to device
    Device_Box_Textured_Render grh_index, _
        screen_x, screen_y, _
        GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, _
        rgb_list, _
        GrhData(grh_index).sX, GrhData(grh_index).sY, _
        alpha_blend, _
        Grh.angle
End Sub

Public Function Map_Particle_Group_Get(ByVal map_x As Integer, ByVal map_y As Integer) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/20/2003
'Checks to see if a tile position has a particle_group_index and return it
'*****************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        Map_Particle_Group_Get = MapData(map_x, map_y).particle_group_index
    Else
        Map_Particle_Group_Get = 0
    End If
End Function

Public Function Map_Base_Light_Get() As Long
'**************************************************************
'Author: Aaron Perkins - Modified by Augusto José Rando
'Last Modify Date: 6/12/2005
'
'**************************************************************
    Map_Base_Light_Get = map_base_light
End Function

Public Function Map_Name_Get() As String
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 7/9/2005
'
'**************************************************************
    Map_Name_Get = MapInfo.Name
End Function

Public Function Map_Sound_Get() As String
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 7/9/2005
'
'**************************************************************
    Map_Sound_Get = MapInfo.SonidoAmbiente
End Function

Public Function Map_Name_Set(ByVal map_name As String)
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 7/9/2005
'
'**************************************************************
    MapInfo.Name = map_name
End Function

Public Function Map_Fill(ByVal grh_index As Long, ByVal layer As Byte, Optional ByVal light_base_color As Long = -1, _
                        Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 1/04/2003
'
'**************************************************************
    Dim x As Integer
    Dim y As Integer
    
    'Base light color
    If light_base_color <> -1 Then
        If Not Map_Base_Light_Set(light_base_color) Then Exit Function
    End If
    
    For y = YMinMapSize To YMaxMapSize
        For x = XMinMapSize To XMaxMapSize
            'Grh
            If Not Map_Grh_Set(x, y, grh_index, layer, alpha_blend, angle) Then Exit Function
        Next x
    Next y
        
    Map_Fill = True
End Function

Public Function Map_Grh_Get(ByVal map_x As Integer, ByVal map_y As Integer, ByVal layer As Byte) As Integer
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 2/20/2003
'Checks to see if a tile position has a grh and return it
'*****************************************************************
    'Check
    If layer < 1 Or layer > 4 Then Exit Function
    
    If Map_In_Bounds(map_x, map_y) Then
        Map_Grh_Get = MapData(map_x, map_y).Graphic(layer).GrhIndex
    End If
End Function

Public Function Map_Grh_Set(ByVal map_x As Integer, ByVal map_y As Integer, ByVal grh_index As Long, _
                            ByVal layer As Byte, Optional ByVal alpha_blend As Boolean, _
                            Optional ByVal angle As Single) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 5/14/2003
'
'**************************************************************
    'Check
    If layer < 1 Or layer > 4 Then Exit Function
    
    If Not Map_In_Bounds(map_x, map_y) Then Exit Function
    
    If Not Grh_Check(grh_index) Then Exit Function
    
    MapData(map_x, map_y).Graphic(layer).GrhIndex = grh_index
    
    'Do it
    Grh_Initialize MapData(map_x, map_y).Graphic(layer), grh_index, alpha_blend, angle
    
    Map_Grh_Set = True
End Function

Public Function Map_Grh_UnSet(ByVal map_x As Integer, ByVal map_y As Integer, ByVal layer As Byte) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 5/14/2003
'
'**************************************************************
    Dim grh_index As Long
    
    'Check
    If layer < 1 Or layer > 5 Then Exit Function
    
    If Not Map_In_Bounds(map_x, map_y) Then Exit Function
    
    grh_index = MapData(map_x, map_y).Graphic(layer).GrhIndex
    If Not Grh_Check(grh_index) Then Exit Function
    
    'Do it
    Grh_Uninitialize MapData(map_x, map_y).Graphic(layer)
    
    Map_Grh_UnSet = True
End Function

Public Function Map_Blocked_Get(ByVal map_x As Integer, ByVal map_y As Integer) As Boolean
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'Checks to see if a tile position is blocked
'*****************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        Map_Blocked_Get = (MapData(map_x, map_y).Blocked = 1)
    Else
        Map_Blocked_Get = True
    End If
End Function

Public Function Map_Item_Add(ByVal s_map_x As Integer, ByVal s_map_y As Integer, ByVal s_item_data_index As Integer, ByVal s_item_amount As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo dodero
'Last Modify Date: 5/14/2002
'
'**************************************************************
    If Map_In_Bounds(s_map_x, s_map_y) Then
        If MapData(s_map_x, s_map_y).ObjInfo.OBJIndex = 0 Then
            MapData(s_map_x, s_map_y).ObjInfo.OBJIndex = s_item_data_index
            MapData(s_map_x, s_map_y).ObjInfo.Amount = s_item_amount
            Grh_Initialize MapData(s_map_x, s_map_y).ObjGrh, Item_Grh_Get(s_item_data_index)
            Map_Item_Add = True
        End If
    End If
End Function

Public Function Map_Item_Grh_Remove(ByVal s_map_x As Integer, ByVal s_map_y As Integer) As Boolean
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 7/9/2005
'
'**************************************************************
    If Map_In_Bounds(s_map_x, s_map_y) Then
        If MapData(s_map_x, s_map_y).ObjGrh.GrhIndex > 0 Then
            If MapData(s_map_x, s_map_y).ObjGrh.GrhIndex = fire_grh_index Then Call Light_Remove(Map_Light_Get(s_map_x, s_map_y))
            Grh_Uninitialize MapData(s_map_x, s_map_y).ObjGrh
            Map_Item_Grh_Remove = True
        End If
    End If
End Function

Public Function Map_Item_Grh_Add(ByVal s_map_x As Integer, ByVal s_map_y As Integer, ByVal s_item_ghr_index As Integer) As Boolean
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 7/9/2005
'
'**************************************************************
    If s_item_ghr_index <= 0 Then Exit Function
    
    If Map_In_Bounds(s_map_x, s_map_y) Then
        If MapData(s_map_x, s_map_y).ObjGrh.GrhIndex > 0 Then Call Map_Item_Grh_Remove(s_map_x, s_map_y)
        
        If s_item_ghr_index = fire_grh_index Then Call Light_Create(s_map_x, s_map_y, , 3)
        
        Grh_Initialize MapData(s_map_x, s_map_y).ObjGrh, s_item_ghr_index
        Call Textures.CargarGrafico(GrhData(s_item_ghr_index).FileNum)
        Map_Item_Grh_Add = True
    End If
End Function

Public Function Map_Fx_Add(ByVal s_map_x As Integer, ByVal s_map_y As Integer, ByVal fx As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo dodero
'Last Modify Date: 5/14/2002
'
'**************************************************************
    
    If Map_In_Bounds(s_map_x, s_map_y) Then
        If MapData(s_map_x, s_map_y).fx = 0 Then
            MapData(s_map_x, s_map_y).fx = fx
            Map_Fx_Add = True
        End If
    End If
End Function

Public Function Map_Item_Remove(ByVal s_map_x As Integer, ByVal s_map_y As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 5/14/2002
'
'**************************************************************
        
    If Map_In_Bounds(s_map_x, s_map_y) Then
        If MapData(s_map_x, s_map_y).ObjInfo.OBJIndex <> 0 Then
            MapData(s_map_x, s_map_y).ObjInfo.OBJIndex = 0
            MapData(s_map_x, s_map_y).ObjInfo.Amount = 0
            Grh_Uninitialize MapData(s_map_x, s_map_y).ObjGrh
            Map_Item_Remove = True
        End If
    End If
End Function

Public Function Map_Item_Get(ByVal s_map_x As Integer, ByVal s_map_y As Integer, ByRef r_item_data_index As Integer, ByRef r_item_amount As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo dodero
'Last Modify Date: 5/14/2002
'Returns item information
'**************************************************************
    If Map_In_Bounds(s_map_x, s_map_y) Then
        If MapData(s_map_x, s_map_y).ObjInfo.OBJIndex <> 0 Then
            r_item_data_index = MapData(s_map_x, s_map_y).ObjInfo.OBJIndex
            r_item_amount = MapData(s_map_x, s_map_y).ObjInfo.Amount
            Map_Item_Get = True
        End If
    End If
End Function

Public Function Map_Item_Grh_In_Current_Area(ByVal ghr_index As Integer, ByRef x_pos As Integer, ByRef y_pos As Integer) As Boolean
'*****************************************************************
'Author: Augusto José Rando
'*****************************************************************
    On Error GoTo ErrorHandler
    
    Dim map_x As Integer
    Dim map_y As Integer
    Dim x As Integer, y As Integer
    
    Call Char_Pos_Get(char_current, map_x, map_y)
    
    If Map_In_Bounds(map_x, map_y) Then
        For y = map_y - MinYBorder + 1 To map_y + MinYBorder - 1
          For x = map_x - MinXBorder + 1 To map_x + MinXBorder - 1
                If MapData(x, y).ObjGrh.GrhIndex = ghr_index Then
                    x_pos = x
                    y_pos = y
                    Map_Item_Grh_In_Current_Area = True
                    Exit Function
                End If
          Next x
        Next y
    End If
    
    Exit Function
    
ErrorHandler:
    Map_Item_Grh_In_Current_Area = False
    
End Function

Public Function Map_NPC_Add(ByVal s_map_x As Integer, ByVal s_map_y As Integer, ByVal s_npc_data_index As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo dodero
'Last Modify Date: 5/14/2002
'
'**************************************************************
    
    On Error GoTo Error_Handle
    
    Dim temp_arr() As Integer
    Dim temp_arr_lng() As Long
    
    If s_npc_data_index > 0 Then
        If Map_Legal_Char_Pos(s_map_x, s_map_y) Then
            If MapData(s_map_x, s_map_y).NPCIndex = 0 Then
            
                MapData(s_map_x, s_map_y).NPCIndex = s_npc_data_index
            
                If s_npc_data_index < 500 Then
                    'Create a char and plot in on the map
                    Char_Create s_map_x, s_map_y, 5, NPC_Data(s_npc_data_index).BodyIndex, NPC_Data(s_npc_data_index).HeadIndex, 0, 0, 0, "", 0, 0, 0, 0, 0, 0, 0, temp_arr_lng(), temp_arr()
                Else
                    'Create a char and plot in on the map
                    Char_Create s_map_x, s_map_y, 5, NPC_Hostile_Data(s_npc_data_index).BodyIndex, NPC_Hostile_Data(s_npc_data_index).HeadIndex, 0, 0, 0, "", 0, 0, 0, 0, 0, 0, 0, temp_arr_lng(), temp_arr()
                End If
                
                Map_NPC_Add = True
            
            End If
        End If
    End If
    
    Exit Function

Error_Handle:
    Map_NPC_Add = False
    
End Function

Public Function Map_NPC_Remove(ByVal s_map_x As Integer, ByVal s_map_y As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 5/14/2002
'
'**************************************************************
    If Map_In_Bounds(s_map_x, s_map_y) Then
        If MapData(s_map_x, s_map_y).NPCIndex <> 0 Then
            MapData(s_map_x, s_map_y).NPCIndex = 0
            'Get rid of the char on the map
            Char_Remove Map_Char_Get(s_map_x, s_map_y)
            Map_NPC_Remove = True
        End If
    End If
End Function

Public Sub Map_Edges_Blocked_Set(ByVal edge_distance_x As Integer, ByVal edge_distance_y As Integer, ByVal Blocked As Byte)
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 2/27/2003
'
'**************************************************************
    Dim x As Integer
    Dim y As Integer
    
    For y = YMinMapSize To YMaxMapSize
        For x = XMinMapSize To XMaxMapSize
            If x <= edge_distance_x Or y <= edge_distance_y Then
                MapData(x, y).Blocked = Blocked
            End If

            If x > XMaxMapSize - edge_distance_x Or y > YMinMapSize - edge_distance_y Then
                MapData(x, y).Blocked = Blocked
            End If
        Next x
    Next y
End Sub

Public Function Particle_Type_Get(ByVal particle_index As Long) As Long
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modify Date: 8/27/2003
'Returns the stream type of a particle stream
'*****************************************************************
    If Particle_Group_Check(particle_index) Then
        Particle_Type_Get = particle_group_list(particle_index).stream_type
    End If
End Function

Public Function Map_NPC_Get(ByVal s_map_x As Integer, ByVal s_map_y As Integer) As Integer
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo dodero
'Last Modify Date: 5/14/2002
'Returns NPC's data index
'**************************************************************
    If Map_In_Bounds(s_map_x, s_map_y) Then
        If MapData(s_map_x, s_map_y).NPCIndex <> 0 Then
            Map_NPC_Get = MapData(s_map_x, s_map_y).NPCIndex
        End If
    End If
End Function

Public Function Map_Exit_Add(ByVal s_map_x As Integer, ByVal s_map_y As Integer, ByVal s_exit_map_num As Integer, ByVal s_exit_map_x As Integer, ByVal s_exit_map_y As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 5/14/2002
'
'**************************************************************
    If Map_In_Bounds(s_map_x, s_map_y) Then
        If MapData(s_map_x, s_map_y).TileExit.Map = 0 Then
            MapData(s_map_x, s_map_y).TileExit.Map = s_exit_map_num
            MapData(s_map_x, s_map_y).TileExit.x = s_exit_map_x
            MapData(s_map_x, s_map_y).TileExit.y = s_exit_map_y
            Map_Exit_Add = True
        End If
    End If
End Function

Public Function Map_Exit_Remove(ByVal s_map_x As Integer, ByVal s_map_y As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 5/14/2002
'
'**************************************************************
    If Map_In_Bounds(s_map_x, s_map_y) Then
        MapData(s_map_x, s_map_y).TileExit.Map = 0
        MapData(s_map_x, s_map_y).TileExit.x = 0
        MapData(s_map_x, s_map_y).TileExit.y = 0
        Map_Exit_Remove = True
    End If
End Function

Public Function Map_Exit_Get(ByVal s_map_x As Integer, ByVal s_map_y As Integer, ByRef r_exit_map_num As Integer, ByRef r_exit_map_x As Integer, ByRef r_exit_map_y As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 5/14/2002
'Returns exit information
'**************************************************************
    If Map_In_Bounds(s_map_x, s_map_y) Then
        If MapData(s_map_x, s_map_y).TileExit.Map <> 0 Then
            r_exit_map_num = MapData(s_map_x, s_map_y).TileExit.Map
            r_exit_map_x = MapData(s_map_x, s_map_y).TileExit.x
            r_exit_map_y = MapData(s_map_x, s_map_y).TileExit.y
            Map_Exit_Get = True
        End If
    End If
End Function

Public Function Map_Trigger_Set(ByVal map_x As Integer, ByVal map_y As Integer, ByVal Trigger As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins
'
'**************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        MapData(map_x, map_y).Trigger = Trigger
        Map_Trigger_Set = True
    Else
        Map_Trigger_Set = False
    End If
End Function

Public Function Map_Trigger_UnSet(ByVal map_x As Integer, ByVal map_y As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins
'
'**************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        If MapData(map_x, map_y).Trigger <> 0 Then
            MapData(map_x, map_y).Trigger = 0
            Map_Trigger_UnSet = True
        End If
    Else
        Map_Trigger_UnSet = False
    End If
End Function

Public Function Map_Trigger_Get(ByVal map_x As Integer, ByVal map_y As Integer) As Integer
'**************************************************************
'Author: Aaron Perkins
'
'**************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        Map_Trigger_Get = MapData(map_x, map_y).Trigger
    End If
End Function

Public Sub Map_Bounds_Get(ByRef map_x_max As Integer, ByRef map_y_max As Integer)
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 2/17/2003
'
'**************************************************************
    'Get size
    map_x_max = XMaxMapSize
    map_y_max = YMaxMapSize
End Sub

Public Function Map_Blocked_Set(ByVal map_x As Integer, ByVal map_y As Integer, ByVal Blocked As Byte) As Boolean
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'Sets a tile position to blocked
'*****************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        MapData(map_x, map_y).Blocked = Blocked
        Map_Blocked_Set = True
    End If
End Function

Function Map_Water_Get(map_x As Integer, map_y As Integer) As Boolean
'**************************************************************
'Author: Augusto José Rando
'**************************************************************

If Map_In_Bounds(map_x, map_y) Then
    If MapData(map_x, map_y).Graphic(1).GrhIndex >= 1505 And _
       MapData(map_x, map_y).Graphic(1).GrhIndex <= 1520 And _
       MapData(map_x, map_y).Graphic(2).GrhIndex = 0 Then
            Map_Water_Get = True
    Else
        Map_Water_Get = False
    End If
End If

End Function

Public Function Map_Base_Light_Set(ByVal base_light As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    If map_base_light <> base_light Then
        map_base_light = base_light
        MapInfo.base_light = base_light
    End If
End Function

Public Function Map_In_Bounds(ByVal x As Integer, ByVal y As Integer) As Boolean

If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
    Map_In_Bounds = False
    Exit Function
End If

Map_In_Bounds = True

End Function

Public Function Char_Current_Pos_Refresh(ByVal map_x As Integer, ByVal map_y As Integer)
'*****************************************************************
'Author: Augusto José Rando
'*****************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        If MapData(CharList(char_current).Pos.x, CharList(char_current).Pos.y).CharIndex = char_current Then MapData(CharList(char_current).Pos.x, CharList(char_current).Pos.y).CharIndex = 0
        CharList(char_current).Pos.x = map_x
        CharList(char_current).Pos.y = map_y
        MapData(map_x, map_y).CharIndex = char_current
    End If
End Function

Public Function Map_Legal_Current_Char_Pos(ByVal map_x As Integer, ByVal map_y As Integer) As Boolean
'*****************************************************************
'Author: Augusto José Rando
'Last Modify Date: 10/6/2005
'Checks to see if a map position is a legal pos for the current
'character
'*****************************************************************
    Dim water_flag As Boolean
    water_flag = Map_Water_Get(map_x, map_y)
    
    If Map_In_Bounds(map_x, map_y) = False Then Exit Function
    
    If MapData(map_x, map_y).Blocked Then Exit Function
    
    If MapData(map_x, map_y).CharIndex Then Exit Function
    
    If water_flag And (char_current_overwater = False) Then Exit Function
    
    If Not water_flag And (char_current_overwater = True) Then Exit Function
    
    If ((Map_Trigger_Get(map_x, map_y) = 1) Or (Map_Trigger_Get(map_x, map_y) = 4)) And char_current_onhorse Then Exit Function
    
    Map_Legal_Current_Char_Pos = True
End Function

Public Function Map_Legal_Char_Pos(ByVal map_x As Integer, ByVal map_y As Integer) As Boolean
'*****************************************************************
'Author: Aaron Perkins - Modified by Augusto José Rando
'Last Modify Date: 10/6/2005
'Checks to see if a map position is a legal pos for a char
'*****************************************************************
    If Map_In_Bounds(map_x, map_y) = False Then Exit Function
    
    If MapData(map_x, map_y).Blocked Then Exit Function
    
    If MapData(map_x, map_y).CharIndex Then Exit Function
        
    Map_Legal_Char_Pos = True
End Function

Public Function Map_Legal_Char_Pos_By_Heading(ByVal CharIndex As Integer, ByVal Heading As Byte) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/17/2003
'Checks to see if a map position is a legal pos for a char
'*****************************************************************
    'Invalid heading
    If Heading < 1 Or Heading > 8 Then
        Exit Function
    End If
    
    'Make sure it's a legal CharIndex
    If CharIndex <> 0 Then
        Dim nX As Integer
        Dim nY As Integer
        nX = CharList(CharIndex).Pos.x
        nY = CharList(CharIndex).Pos.y
        Convert_Heading_to_Direction Heading, nX, nY
            
        Map_Legal_Char_Pos_By_Heading = Map_Legal_Char_Pos(nX, nY)
    End If
End Function

Public Function Map_Char_Get(ByVal map_x As Integer, ByVal map_y As Integer) As Integer
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/20/2003
'Checks to see if a tile position has a CharIndex and return it
'*****************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        Map_Char_Get = MapData(map_x, map_y).CharIndex
    Else
        Map_Char_Get = 0
    End If
End Function

Public Function Map_Char_Set(ByVal map_x As Integer, ByVal map_y As Integer, ByVal char_index As Integer) As Boolean
    If Map_In_Bounds(map_x, map_y) Then
        MapData(map_x, map_y).CharIndex = char_index
        Map_Char_Set = True
    Else
        Map_Char_Set = False
    End If
End Function

Public Function Map_Char_UnSet(ByVal map_x As Integer, ByVal map_y As Integer) As Boolean
    If Map_In_Bounds(map_x, map_y) Then
        MapData(map_x, map_y).CharIndex = 0
        Map_Char_UnSet = True
    Else
        Map_Char_UnSet = False
    End If
End Function

Public Function Map_Light_Get(ByVal map_x As Integer, ByVal map_y As Integer) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/20/2003
'Checks to see if a tile position has a light_index and return it
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    'We start from the back, to get the last light to be placed on the tile first
    loopc = light_last
    Do Until light_list(loopc).map_x = map_x And light_list(loopc).map_y = map_y
        If loopc = 0 Then
            Map_Light_Get = 0
            Exit Function
        End If
        loopc = loopc - 1
    Loop
    
    Map_Light_Get = loopc
Exit Function
ErrorHandler:
    Map_Light_Get = 0
End Function

Private Sub Map_Render()
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 5/14/2003
'
'**************************************************************
    Dim map_x As Integer
    Dim map_y As Integer
    Dim view_x As Integer
    Dim view_y As Integer
    Dim screen_x As Integer
    Dim screen_y As Integer
    
    Dim view_min_y As Integer
    Dim view_max_y As Integer
    Dim view_min_x As Integer
    Dim view_max_x As Integer
    
    Dim view_buffer_min_y As Integer
    Dim view_buffer_max_y As Integer
    Dim view_buffer_min_x As Integer
    Dim view_buffer_max_x As Integer
    
    Dim view_buffer_offset_x As Integer
    Dim view_buffer_offset_y As Integer
    
    Dim view_x_start_value As Integer
    Dim view_y_start_value As Integer
    
    Dim Grh As Grh
    
    Dim temp_array(3) As Long
    temp_array(0) = &HFFFFFF
    temp_array(1) = &HFFFFFF
    temp_array(2) = &HFFFFFF
    temp_array(3) = &HFFFFFF
    
    '*********************
    'Handle scrolling
    'counters
    '*********************
    If scroll_on Then
        '****** Move screen Left and Right if needed ******
        If scroll_direction_x <> 0 Then
            scroll_offset_counter_x = scroll_offset_counter_x + (scroll_pixels_per_frame * timer_ticks_per_frame * scroll_direction_x)
            If Sgn(scroll_offset_counter_x) = scroll_direction_x Then
                scroll_offset_counter_x = 0
                scroll_direction_x = 0
            End If
        End If
        '****** Move screen Up and Down if needed ******
        If scroll_direction_y <> 0 Then
            scroll_offset_counter_y = scroll_offset_counter_y + (scroll_pixels_per_frame * timer_ticks_per_frame * scroll_direction_y)
            If Sgn(scroll_offset_counter_y) = scroll_direction_y Then
                scroll_offset_counter_y = 0
                scroll_direction_y = 0
            End If
        End If
        'End scrolling if needed
        If scroll_direction_x = 0 And scroll_direction_y = 0 Then
            scroll_on = False
        End If
    End If
    
    'Figure out ends and starts of view area
    view_min_y = view_pos_y - view_screen_tile_height \ 2
    view_max_y = view_pos_y + view_screen_tile_height \ 2
    view_min_x = view_pos_x - view_screen_tile_width \ 2
    view_max_x = view_pos_x + view_screen_tile_width \ 2
    
    'Add the buffer
    view_buffer_min_y = view_min_y - view_tile_buffer
    view_buffer_max_y = view_max_y + view_tile_buffer
    view_buffer_min_x = view_min_x - view_tile_buffer
    view_buffer_max_x = view_max_x + view_tile_buffer
    
    'Reduce buffer size to avoid going off limits
    If view_buffer_min_y < YMinMapSize Then
        view_buffer_offset_y = YMinMapSize - view_buffer_min_y
        view_buffer_min_y = YMinMapSize
    End If
    
    If view_buffer_max_y > YMaxMapSize Then view_buffer_max_y = YMaxMapSize
    
    If view_buffer_min_x < XMinMapSize Then
        view_buffer_offset_x = XMinMapSize - view_buffer_min_x
        view_buffer_min_x = XMinMapSize
    End If
    
    If view_buffer_max_x > XMaxMapSize Then view_buffer_max_x = XMaxMapSize
    
    view_buffer_offset_x = view_buffer_offset_x - view_tile_buffer + 1
    view_buffer_offset_y = view_buffer_offset_y - view_tile_buffer + 1
    
    'Only attempt to render layer floor beyond edges if in map bounds
    If view_min_y < YMinMapSize Then
        view_y_start_value = -view_min_y + 2
        view_min_y = YMinMapSize
    ElseIf view_min_y > YMinMapSize Then
        view_min_y = view_min_y - 1
    Else
        view_y_start_value = 1
    End If
    
    If view_max_y > YMaxMapSize Then
        view_max_y = YMaxMapSize
    ElseIf view_max_y < YMaxMapSize Then
        view_max_y = view_max_y + 1
    End If
    
    If view_min_x < XMinMapSize Then
        view_x_start_value = -view_min_x + 2
        view_min_x = XMinMapSize
    ElseIf view_min_x > XMinMapSize Then
        view_min_x = view_min_x - 1
    Else
        view_x_start_value = 1
    End If
    
    If view_max_x > XMaxMapSize Then
        view_max_x = XMaxMapSize
    ElseIf view_max_x < XMaxMapSize Then
        view_max_x = view_max_x + 1
    End If
    
    '*********************
    'Layer 1
    '*********************
    view_y = view_y_start_value
    For map_y = view_min_y To view_max_y
        view_x = view_x_start_value
        For map_x = view_min_x To view_max_x
            '*** Start Layer 1 ***
            If MapData(map_x, map_y).Graphic(1).GrhIndex Then
                screen_x = Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x
                screen_y = Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y
                Grh_Render MapData(map_x, map_y).Graphic(1), screen_x, screen_y, MapData(map_x, map_y).light_value()
            End If
            '*** End Layer 1 ***
            view_x = view_x + 1
        Next map_x
        view_y = view_y + 1
    Next map_y
    
    '*********************
    'Layer 2
    '*********************
    view_y = view_buffer_offset_y
    For map_y = view_buffer_min_y To view_buffer_max_y
        view_x = view_buffer_offset_x
        For map_x = view_buffer_min_x To view_buffer_max_x
            '*** Start Layer 2 ***
            If MapData(map_x, map_y).Graphic(2).GrhIndex Then
                screen_x = Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x
                screen_y = Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y
                Grh_Render MapData(map_x, map_y).Graphic(2), screen_x, screen_y, MapData(map_x, map_y).light_value()
            End If
            '*** End Layer 2 ***
                        
            '*** Start Item Layer ***
            If MapData(map_x, map_y).ObjGrh.GrhIndex Then
                If GrhData(MapData(map_x, map_y).ObjGrh.GrhIndex).pixelHeight = 32 And GrhData(MapData(map_x, map_y).ObjGrh.GrhIndex).pixelWidth = 32 Then
                    screen_x = Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x
                    screen_y = Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y
                    Grh_Render MapData(map_x, map_y).ObjGrh, screen_x, screen_y, MapData(map_x, map_y).light_value(), True, True
                End If
            End If
            '*** End Item Layer ***
                        
            view_x = view_x + 1
        Next map_x
        view_y = view_y + 1
    Next map_y
    
    '*********************
    'Middle layer
    '*********************
    view_y = view_buffer_offset_y
    For map_y = view_buffer_min_y To view_buffer_max_y
        view_x = view_buffer_offset_x
        For map_x = view_buffer_min_x To view_buffer_max_x
            '*** Start Layer 3 ***
            If MapData(map_x, map_y).Graphic(3).GrhIndex Then
                screen_x = Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x
                screen_y = Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y
                Grh_Render MapData(map_x, map_y).Graphic(3), screen_x, screen_y, MapData(map_x, map_y).light_value()
            End If
            '*** End Layer 3 ***
            
            '*** Start Characters ***
            If MapData(map_x, map_y).CharIndex Then
                'Figure out screen position
                screen_x = Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x
                screen_y = Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y
                
                If (CharList(MapData(map_x, map_y).CharIndex).invisible) Then
                    If CharList(MapData(map_x, map_y).CharIndex).navigating = 0 Then
                        If Char_Alpha_View(CharList(MapData(map_x, map_y).CharIndex), CharList(char_current)) Then
                            Char_Render CharList(MapData(map_x, map_y).CharIndex), MapData(map_x, map_y).CharIndex, screen_x, screen_y, MapData(map_x, map_y).light_value, True
                        Else
                            Char_Invisible_Render CharList(MapData(map_x, map_y).CharIndex), MapData(map_x, map_y).CharIndex, screen_x, screen_y, MapData(map_x, map_y).light_value, True
                        End If
                    Else
                        Char_Render CharList(MapData(map_x, map_y).CharIndex), MapData(map_x, map_y).CharIndex, screen_x, screen_y, MapData(map_x, map_y).light_value
                    End If
                Else
                    Char_Render CharList(MapData(map_x, map_y).CharIndex), MapData(map_x, map_y).CharIndex, screen_x, screen_y, MapData(map_x, map_y).light_value
                End If
            End If
            '*** End Characters ***
            
            '*** Start MAP Item Layer ***
            If MapData(map_x, map_y).ObjGrh.GrhIndex Then
                If GrhData(MapData(map_x, map_y).ObjGrh.GrhIndex).pixelHeight <> 32 Or GrhData(MapData(map_x, map_y).ObjGrh.GrhIndex).pixelWidth <> 32 Then
                    screen_x = Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x
                    screen_y = Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y
                    Grh_Render MapData(map_x, map_y).ObjGrh, screen_x, screen_y, MapData(map_x, map_y).light_value(), True, True
                End If
            End If
            '*** End MAP Item Layer ***
            
            view_x = view_x + 1
        Next map_x
        view_y = view_y + 1
    Next map_y
    
    '*********************
    'Layer 4
    '*********************
    view_y = view_buffer_offset_y
    For map_y = view_buffer_min_y To view_buffer_max_y
        view_x = view_buffer_offset_x
        For map_x = view_buffer_min_x To view_buffer_max_x
            '*** Start particle effects ***
            If MapData(map_x, map_y).particle_group_index Then
                screen_x = (Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x) + base_tile_size / 2
                screen_y = (Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y) + base_tile_size / 2
                Particle_Group_Render MapData(map_x, map_y).particle_group_index, screen_x, screen_y
            End If
            '*** End particle effects ***
            '*** Start fx effects ***
            If MapData(map_x, map_y).fx <> 0 Then
                screen_x = (Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x)
                screen_y = (Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y)
                Grh_Render_MapFX FxData(MapData(map_x, map_y).fx).fx, map_x, map_y, screen_x + FxData(MapData(map_x, map_y).fx).offsetX, screen_y + FxData(MapData(map_x, map_y).fx).OffsetY, MapData(map_x, map_y).light_value(), True
            End If
            '*** End fx effects ***
            
            '*** Start Layer 4 ***
            If MapData(map_x, map_y).Graphic(4).GrhIndex And engine_show_layer_4 Then
                screen_x = Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x
                screen_y = Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y
                'Grh_Render MapData(map_x, map_y).Graphic(4), screen_x, screen_y, MapData(map_x, map_y).light_value(), MapData(map_x, map_y).Graphic(4).h_centered, MapData(map_x, map_y).Graphic(4).v_centered
                
                If Not under_roof Then
                    Grh_Render MapData(map_x, map_y).Graphic(4), screen_x, screen_y, MapData(map_x, map_y).light_value()
                Else
                    temp_array(0) = 850703540
                    temp_array(1) = 850703540
                    temp_array(2) = 850703540
                    temp_array(3) = 850703540
                    Grh_Render MapData(map_x, map_y).Graphic(4), screen_x, screen_y, temp_array(), , , True
                End If
            
            End If
            '*** End Layer 4 ***
            
            '*** Start Dialogs ***
            If MapData(map_x, map_y).CharIndex Then
                If CharList(MapData(map_x, map_y).CharIndex).dialog <> "" Then
                    'Figure out screen position
                    screen_x = Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x + CharList(MapData(map_x, map_y).CharIndex).scroll_offset_counter_x
                    screen_y = Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y + CharList(MapData(map_x, map_y).CharIndex).scroll_offset_counter_y
                    'Set colors
                    temp_array(0) = CharList(MapData(map_x, map_y).CharIndex).dialog_color
                    temp_array(1) = CharList(MapData(map_x, map_y).CharIndex).dialog_color
                    temp_array(2) = CharList(MapData(map_x, map_y).CharIndex).dialog_color
                    temp_array(3) = CharList(MapData(map_x, map_y).CharIndex).dialog_color
                    'Render dialog
                    Engine_Text_Render CharList(MapData(map_x, map_y).CharIndex).dialog, screen_x + 14 - Engine_Text_Width(CharList(MapData(map_x, map_y).CharIndex).dialog, True) / 2, screen_y + (CharList(MapData(map_x, map_y).CharIndex).Body.HeadOffset.y) - Engine_Text_Height(CharList(MapData(map_x, map_y).CharIndex).dialog, True), temp_array, CharList(MapData(map_x, map_y).CharIndex).dialog_font_index, True
                End If
            End If
            '*** End Dialogs ***
            
            view_x = view_x + 1
        Next map_x
        
        view_y = view_y + 1
    Next map_y

    '*********************
    'Special Layer
    '*********************
    If engine_show_blocked_tiles Or engine_show_special_tiles Or engine_show_triggers Then
        
        view_y = view_buffer_offset_y
        
        For map_y = view_buffer_min_y To view_buffer_max_y
            view_x = view_buffer_offset_x
            For map_x = view_buffer_min_x To view_buffer_max_x
                If Map_In_Bounds(map_x, map_y) Then
                    screen_x = Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x
                    screen_y = Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y
                    
                    If engine_show_blocked_tiles Then
                        '*** Start Blocked Tiles ***
                        If MapData(map_x, map_y).Blocked Then
                            Grh.GrhIndex = 4
                            Grh_Render Grh, screen_x, screen_y, temp_array()
                        End If
                        '*** End Blocked Tiles ***
                    End If
                    
                   If engine_show_special_tiles Then
                        '*** Start Special Tiles ***
                        If MapData(map_x, map_y).TileExit.Map Then
                            Grh.GrhIndex = 2
                            Grh_Render Grh, screen_x, screen_y, temp_array
                        End If
                    End If
                    
                    If engine_show_triggers Then
                        '*** Start Trigger ***
                        Device_Text_Render font_list(1), CStr(MapData(map_x, map_y).Trigger), screen_y, screen_x, base_tile_size, base_tile_size, &HFFFFFFFF, DT_VCENTER Or DT_CENTER
                    End If
                    
                End If
                
                view_x = view_x + 1
            Next map_x
            view_y = view_y + 1
        Next map_y
    End If
        
    'Rendering of meteorologic particles
    If meteo_particle <> 0 Then
        'Screen positions were hardcoded by now
        screen_x = 250
        screen_y = 0
        Call Particle_Group_Render(meteo_particle, screen_x, screen_y)
    End If
        
    If letter_text <> "" Then
        temp_array(0) = &HFFFFFF
        temp_array(1) = &HFFFFFF
        temp_array(2) = &HFFFFFF
        temp_array(3) = &HFFFFFF
        Grh_Render letter_grh, 250, 300, temp_array()
        Engine_Text_Render letter_text, 250 + 14 - Engine_Text_Width(letter_text, True) / 2, 300 - Engine_Text_Height(letter_text, True), temp_array(), 1, True
    End If

End Sub

Public Function Font_Create(ByVal style As String, ByVal Size As Long, ByVal bold As Boolean, _
                            ByVal italic As Boolean) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/27/2003
'Returns the font_index if successful, else 0
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
On Error GoTo ErrorHandler:
    Font_Create = Font_Next_Open
    
    Font_Make Font_Create, style, bold, italic, Size
    
    Exit Function
ErrorHandler:
    Font_Create = 0
End Function

Private Sub Font_Make(ByVal font_index As Long, ByVal style As String, ByVal bold As Boolean, _
                        ByVal italic As Boolean, ByVal Size As Long)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 3/28/2004
'Modified by Juan Martín Sotuyo Dodero
'*****************************************************************
    'Update array size
    If font_index > font_last Then
        font_last = font_index
        ReDim Preserve font_list(1 To font_last)
    End If
    font_count = font_count + 1
    
    Dim font_desc As IFont 'We use this temporarily to setup the font
    Dim fnt As New StdFont 'This is also used to describe and setup the font
    fnt.Name = style
    fnt.Size = Size
    fnt.bold = bold
    fnt.italic = italic
    
    If font_last = 1 And font_count = 1 Then
        ReDim Preserve font_primitive_list(1 To font_last)
        Set font_primitive_list(font_last) = fnt
    ElseIf font_last > UBound(font_primitive_list()) Then
        ReDim Preserve font_primitive_list(1 To font_last)
        Set font_primitive_list(font_last) = fnt
    End If
    
    Set font_desc = fnt
    Set font_list(font_index) = D3dx.CreateFont(ddevice, font_desc.hFont)
End Sub

Private Function Font_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
    Font_Next_Open = font_last + 1
End Function

Private Function Font_Check(ByVal font_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/27/2003
'
'**************************************************************
    'check font_index
    If font_index > 0 And font_index <= font_last Then
        Font_Check = True
    End If
End Function

Public Function Char_Move_By_Pos(CharIndex As Integer, nX As Integer, nY As Integer)

Dim x As Integer
Dim y As Integer
Dim addX As Integer
Dim addY As Integer
Dim nHeading As Byte

x = CharList(CharIndex).Pos.x
y = CharList(CharIndex).Pos.y

addX = nX - x
addY = nY - y

If Sgn(addX) = 1 Then
    nHeading = EAST
End If

If Sgn(addX) = -1 Then
    nHeading = WEST
End If

If Sgn(addY) = -1 Then
    nHeading = NORTH
End If

If Sgn(addY) = 1 Then
    nHeading = SOUTH
End If

Call Char_Move(CharIndex, nHeading)

End Function

Public Function Char_Move(ByVal CharIndex As Integer, ByVal Heading As Byte) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Returns true if successful, else false
'**************************************************************
    Dim temp_x As Integer
    Dim temp_y As Integer
    
    'Invalid heading
    If Heading < 1 Or Heading > 8 Then
        Char_Move = False
        Exit Function
    End If
    
    'Make sure it's a legal CharIndex
    If CharIndex <> 0 Then
        'Make sure it's a legal move
        temp_x = CharList(CharIndex).Pos.x
        temp_y = CharList(CharIndex).Pos.y
        Convert_Heading_to_Direction Heading, temp_x, temp_y
        If Map_In_Bounds(temp_x, temp_y) Then
            'check for another CharIndex
            If (MapData(temp_x, temp_y).CharIndex = 0) Or (MapData(temp_x, temp_y).CharIndex = char_current) Or (Char_Type_Get(MapData(temp_x, temp_y).CharIndex) = 4 And Char_Body_Get(MapData(temp_x, temp_y).CharIndex) = 0) Then
                'Move it
                Char_Move_By_Head CharIndex, Heading
                
                If char_current = CharIndex Then
                    under_roof = ((Map_Trigger_Get(temp_x, temp_y) = 1) Or (Map_Trigger_Get(temp_x, temp_y) = 4))
                End If
                
                Char_Move = True
            End If
            
        End If
    End If
End Function

Public Function Char_Data_Grh_Index_Get(ByVal char_data_index As Long, ByVal s_body_index As Long, ByVal s_heading As Byte) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/12/2003
'Returns the grh_index for a character given a body_index and heading
'*****************************************************************
    'Invalid heading
    If s_heading < 1 Or s_heading > 8 Then
        Exit Function
    End If

    'Make sure it's a legal CharIndex
    If char_data_index <= UBound(CharList) And char_data_index >= LBound(CharList) Then
        Char_Data_Grh_Index_Get = CharList(char_data_index).Body.Walk(s_heading).GrhIndex
    End If
End Function

Private Sub Char_Move_By_Head(ByVal CharIndex As Integer, ByVal Heading As Byte)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Starts the movement of a character in nHeading direction
'*****************************************************************
    Dim x As Integer
    Dim y As Integer
    Dim nX As Integer
    Dim nY As Integer
    
    x = CharList(CharIndex).Pos.x
    y = CharList(CharIndex).Pos.y
    
    nX = x
    nY = y
    Convert_Heading_to_Direction Heading, nX, nY
    
    MapData(nX, nY).CharIndex = CharIndex
    CharList(CharIndex).Pos.x = nX
    CharList(CharIndex).Pos.y = nY
    MapData(x, y).CharIndex = 0
    
    CharList(CharIndex).scroll_offset_counter_x = (base_tile_size * (x - nX))
    CharList(CharIndex).scroll_offset_counter_y = (base_tile_size * (y - nY))
    CharList(CharIndex).scroll_direction_x = nX - x
    CharList(CharIndex).scroll_direction_y = nY - y
    
    CharList(CharIndex).scroll_on = True
    CharList(CharIndex).Heading = Heading
    'Set char to walk
    CharList(CharIndex).Body.Walk(Heading).started = 1 '???
End Sub

Public Function Char_Heading_Set(ByVal char_index As Integer, ByVal Heading As Byte) As Boolean
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 2/20/2003
'Changes the character heading
'*****************************************************************
   'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        CharList(char_index).Heading = Heading
        CharList(char_index).Body.Walk(Heading).started = 0
        CharList(char_index).Body.Walk(Heading).framecounter = 1
        
        CharList(char_index).Arma.WeaponWalk(Heading).started = 0
        CharList(char_index).Arma.WeaponWalk(Heading).framecounter = 1
        
        CharList(char_index).Escudo.ShieldWalk(Heading).started = 0
        CharList(char_index).Escudo.ShieldWalk(Heading).framecounter = 1
        
        Char_Heading_Set = True
    End If
End Function

Public Function Char_Pos_Set(ByVal char_index As Integer, ByVal map_x As Integer, ByVal map_y As Integer) As Boolean
'*****************************************************************
'Author: Aaron Perkins - Modified by Augusto José Rando
'Last Modify Date: 2/20/2003
'Changes the character position
'*****************************************************************
   'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        Call Map_Char_UnSet(CharList(char_index).Pos.x, CharList(char_index).Pos.y)
        Call Map_Char_Set(map_x, map_y, char_index)
        CharList(char_index).Pos.x = map_x
        CharList(char_index).Pos.y = map_y
        
        If char_index = char_current Then
            Call Engine_View_Pos_Set(map_x, map_y)
            Call Char_Current_Pos_Refresh(map_x, map_y)
        End If
        
        Char_Pos_Set = True
    End If
End Function

Public Function Char_Type_Set(ByVal char_index As Integer, ByVal char_type As Byte) As Boolean
'*****************************************************************
'Author: Augusto José Rando
'Last Modify Date: 6/12/2005
'Changes the character type
'*****************************************************************
   'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        CharList(char_index).CharType = char_type
        Char_Type_Set = True
    End If
End Function

Public Function Char_FX_Set(ByVal char_index As Integer, ByVal fx As Integer, ByVal fx_loops As Integer) As Boolean
'*****************************************************************
'Author: Augusto José Rando
'Adds an FX to the character
'*****************************************************************
      
   'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        CharList(char_index).fx = fx
        CharList(char_index).FxLoopTimes = fx_loops
        Char_FX_Set = True
    End If
End Function

Public Function Char_Invisible_Set(ByVal char_index As Integer, ByVal char_invisible As Byte) As Boolean
'*****************************************************************
'Author: Augusto José Rando
'Makes (or not) invisible a character
'*****************************************************************
   'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        CharList(char_index).invisible = char_invisible
        Char_Invisible_Set = True
    End If
End Function

Public Function Char_Fly_Set(ByVal char_index As Integer, ByVal char_fly As Byte) As Boolean
'*****************************************************************
'Author: Augusto José Rando
'Makes (or not) flying a character
'*****************************************************************
   'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        CharList(char_index).fly = char_fly
        Char_Fly_Set = True
    End If
End Function

Public Function Char_Current_OverWater_Set(ByVal overwater As Boolean) As Boolean
'*****************************************************************
'Author: Augusto José Rando
'Makes (or not) able to cross water current char
'*****************************************************************
    char_current_overwater = overwater
    Char_Current_OverWater_Set = True
End Function

Public Function Char_Current_OnHorse_Set(ByVal onhorse As Boolean) As Boolean
'*****************************************************************
'Author: Augusto José Rando
'Makes (or not) able to cross water current char
'*****************************************************************
    char_current_onhorse = onhorse
    Char_Current_OnHorse_Set = True
End Function

Public Function Char_Current_Blind_Set(ByVal blind As Boolean) As Boolean
'*****************************************************************
'Author: Augusto José Rando
'Makes (or not) able to cross water current char
'*****************************************************************
    char_current_blind = blind
    Char_Current_Blind_Set = True
End Function

Public Function Char_Current_Get() As Integer
'*****************************************************************
'Author: Augusto José Rando
'*****************************************************************
    Char_Current_Get = char_current
End Function

Public Function Char_Current_Set(ByVal char_index As Integer)
'*****************************************************************
'Author: Augusto José Rando
'*****************************************************************
    char_current = char_index
End Function

Public Function Char_Distance_Get(ByVal char_index As Integer) As Integer
'*****************************************************************
'Author: Augusto José Rando
'Gets the distance from the current char to another
'*****************************************************************

    On Error GoTo ErrorHandler

    Dim map_x As Integer
    Dim map_y As Integer
    Dim map_current_x As Integer
    Dim map_current_y As Integer
    
    If Char_Check(char_index) Then
        Call Char_Pos_Get(char_index, map_x, map_y)
        Call Char_Pos_Get(char_index, map_current_x, map_current_y)
        
        Char_Distance_Get = Abs(map_x - map_current_x) + Abs(map_y - map_current_y)
        
    End If

ErrorHandler:
    Char_Distance_Get = 0

End Function

Public Function Char_In_Current_Area(ByVal char_index As Integer) As Boolean
'*****************************************************************
'Author: Augusto José Rando
'*****************************************************************
    On Error GoTo ErrorHandler
    
    Dim map_x As Integer
    Dim map_y As Integer
    Dim x As Integer, y As Integer
    
    Call Char_Pos_Get(char_current, map_x, map_y)
    If map_x = 0 Or map_y = 0 Then Exit Function
    
    For y = map_y - MinYBorder + 1 To map_y + MinYBorder - 1
      For x = map_x - MinXBorder + 1 To map_x + MinXBorder - 1
            If MapData(x, y).CharIndex = char_index Then
                Char_In_Current_Area = True
                Exit Function
            End If
      Next x
    Next y
    
    Exit Function
    
ErrorHandler:
    Char_In_Current_Area = False
    
End Function

Public Function Char_Feet_Switch(ByVal char_index As Integer) As Integer
'*****************************************************************
'Author: Augusto José Rando
'*****************************************************************
    If Char_Check(char_index) Then
        If CharList(char_index).fly = 0 Then
            If CharList(char_index).Pie Then
                CharList(char_index).Pie = False
                Char_Feet_Switch = 1
            Else
                CharList(char_index).Pie = True
                Char_Feet_Switch = 0
            End If
        Else
            Char_Feet_Switch = -1
        End If
    End If

End Function

Public Function Char_Body_Get(ByVal char_index As Integer) As Integer
'*****************************************************************
'Author: Aaron Perkins
'*****************************************************************
   'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        Char_Body_Get = CharList(char_index).iBody
    End If
End Function

Public Function Char_Body_Ship_Get(ByVal char_body As Integer) As Byte
'*****************************************************************
'Author: Augusto José Rando
'Used to check if user is on a ship. Bodys are hardcoded.
'*****************************************************************
Char_Body_Ship_Get = ((char_body = 84) Or (char_body = 85) Or (char_body = 86) Or (char_body = 87))
End Function

Public Function Char_Particle_Group_Create(ByVal char_index As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                        Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                        Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                        Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                        Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                        Optional grh_resizex As Integer, Optional grh_resizey As Integer)
'**************************************************************
'Author: Augusto José Rando
'**************************************************************
    Dim char_part_free_index As Integer
    
    'If Char_Particle_Group_Find(char_index, stream_type) Then Exit Function ' hay que ver si dejar o sacar esto...
    If Not Char_Check(char_index) Then Exit Function
    char_part_free_index = Char_Particle_Group_Next_Open(char_index)
    
    If char_part_free_index > 0 Then
        Char_Particle_Group_Create = Particle_Group_Next_Open
        Char_Particle_Group_Make Char_Particle_Group_Create, char_index, char_part_free_index, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
    End If

End Function

Public Function Char_Particle_Group_Remove(ByVal char_index As Integer, ByVal stream_type As Long)
'**************************************************************
'Author: Augusto José Rando
'**************************************************************
    Dim char_part_index As Integer
    
    If Char_Check(char_index) Then
        char_part_index = Char_Particle_Group_Find(char_index, stream_type)
        If char_part_index = -1 Then Exit Function
        Call Particle_Group_Remove(char_part_index)
    End If

End Function

Public Function Char_Particle_Group_Remove_All(ByVal char_index As Integer)
'**************************************************************
'Author: Augusto José Rando
'**************************************************************
    Dim i As Integer
    
    If Char_Check(char_index) Then
        For i = 1 To CharList(char_index).particle_count
            If CharList(char_index).particle_group(i) <> 0 Then Call Particle_Group_Remove(CharList(char_index).particle_group(i))
        Next i
    End If
    
End Function

Private Function Char_Particle_Group_Find(ByVal char_index As Integer, ByVal stream_type As Long) As Integer
'*****************************************************************
'Author: Augusto José Rando
'Modified: returns slot or -1
'*****************************************************************

Dim i As Integer

For i = 1 To CharList(char_index).particle_count
    If particle_group_list(CharList(char_index).particle_group(i)).stream_type = stream_type Then
        Char_Particle_Group_Find = CharList(char_index).particle_group(i)
        Exit Function
    End If
Next i

Char_Particle_Group_Find = -1

End Function

Private Function Char_Particle_Group_Next_Open(ByVal char_index As Integer) As Integer
'*****************************************************************
'Author: Augusto José Rando
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until CharList(char_index).particle_group(loopc) = 0
        If loopc = CharList(char_index).particle_count Then
            Char_Particle_Group_Next_Open = CharList(char_index).particle_count + 1
            CharList(char_index).particle_count = Char_Particle_Group_Next_Open
            ReDim Preserve CharList(char_index).particle_group(1 To Char_Particle_Group_Next_Open) As Long
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Char_Particle_Group_Next_Open = loopc

Exit Function

ErrorHandler:
    CharList(char_index).particle_count = 1
    ReDim CharList(char_index).particle_group(1 To 1) As Long
    Char_Particle_Group_Next_Open = 1

End Function

Public Function Char_Heading_Get(ByVal char_index As Integer) As Byte
'*****************************************************************
'Author: Aaron Perkins
'*****************************************************************
   'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        Char_Heading_Get = CharList(char_index).Heading
    End If
End Function

Public Function Char_Dead_Get(ByVal char_index As Integer) As Boolean
'*****************************************************************
'Author: Augusto José Rando
'*****************************************************************
   'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        Char_Dead_Get = CharList(char_index).Muerto
    End If
End Function

Public Function Char_User_Roof_Get() As Boolean
'*****************************************************************
'Author: Augusto José Rando
'*****************************************************************
   'Make sure it's a legal char_current
    If Char_Check(char_current) Then
        Char_User_Roof_Get = under_roof
    End If
End Function

Public Function Char_Pos_Get(ByVal char_index As Integer, ByRef map_x As Integer, ByRef map_y As Integer) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'*****************************************************************
   'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        map_x = CharList(char_index).Pos.x
        map_y = CharList(char_index).Pos.y
        Char_Pos_Get = True
    End If
End Function

Public Function Char_Type_Get(ByVal char_index As Integer) As Byte
'*****************************************************************
'Author: Augusto José Rando
'*****************************************************************
   'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        Char_Type_Get = CharList(char_index).CharType
    End If
End Function

Public Function Char_Name_Get(ByVal char_index As Integer) As String
'*****************************************************************
'Author: Augusto José Rando
'Gets the character name
'*****************************************************************
   'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        Char_Name_Get = CharList(char_index).nombre
    End If
End Function

Public Function Char_Name_Set(ByVal char_index As Integer, ByVal char_name As String) As Boolean
'*****************************************************************
'Author: Augusto José Rando
'Changes the character name
'*****************************************************************
   'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        CharList(char_index).nombre = char_name
        Char_Name_Set = True
    End If
End Function

Public Function Char_Group_Set(ByVal char_index As Integer, ByVal char_group_id As Integer) As Boolean
'*****************************************************************
'Author: Augusto José Rando
'Changes the character group
'*****************************************************************
   'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        CharList(char_index).group_index = char_group_id
        Char_Group_Set = True
    End If
End Function

Private Function Char_Check(ByVal char_index As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check char_index
    If char_index > 0 And char_index <= LastChar Then
        Char_Check = (CharList(char_index).Heading > 0)
    End If
    
End Function

Public Function Char_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim Index As Long
    
    For Index = 1 To LastChar
        'Make sure it's a legal index
        If Char_Check(Index) Then
            Char_Destroy Index
        End If
    Next Index
    
    Char_Remove_All = True
End Function

Public Function Char_Find(ByVal id As Integer) As Integer
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    If id <= 0 Then Exit Function
    
    loopc = 1
    Do Until CharList(loopc).id = id
        If loopc = LastChar Then Exit Function
        loopc = loopc + 1
    Loop
    
    Char_Find = loopc
Exit Function

ErrorHandler:
End Function

Public Function Char_Remove(ByVal char_index As Integer) As Boolean
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Char_Check(char_index) Then
        Char_Particle_Group_Remove_All (char_index)
        Char_Destroy char_index
        Char_Remove = True
    End If
End Function

Private Sub Char_Destroy(ByVal char_index As Integer)
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim temp As char
        
    If Map_In_Bounds(CharList(char_index).Pos.x, CharList(char_index).Pos.y) Then MapData(CharList(char_index).Pos.x, CharList(char_index).Pos.y).CharIndex = 0
    CharList(char_index) = temp
    
    'Update array size
    If char_index = LastChar Then
        Do Until CharList(LastChar).Heading > 0
            LastChar = LastChar - 1
            If LastChar = 0 Then
                NumChars = 0
                Exit Sub
            End If
        Loop
        ReDim Preserve CharList(1 To LastChar)
    End If
    
    NumChars = NumChars - 1
End Sub

Private Function Char_Next_Open() As Integer
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until CharList(loopc).Heading = 0
        If loopc = LastChar Then
            Char_Next_Open = LastChar + 1
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Char_Next_Open = loopc
Exit Function

ErrorHandler:
    Char_Next_Open = 1
End Function

Public Function Char_Create(ByVal map_x As Integer, ByVal map_y As Integer, ByVal Heading As Byte, _
                            ByVal body_index As Integer, ByVal head_index As Integer, _
                            ByVal helmet_index As Integer, ByVal weapon_index As Integer, ByVal shield_index As Integer, _
                            ByVal Name As String, ByVal CharType As Integer, ByVal fx As Integer, ByVal FxLoopTimes As Integer, _
                            ByVal id As Integer, ByVal char_invisible As Byte, ByVal char_fly As Byte, ByVal particle_count As Integer, _
                            ByRef particle_life_counter() As Long, ByRef particle_type() As Integer) As Integer
                            
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Modified by Augusto José Rando
'Last Modify Date: 6/9/2005
'Returns the char_index if successful, else 0
'**************************************************************
    'Invalid heading
    If Heading < 1 Or Heading > 8 Then Exit Function
    
    If Map_Char_Get(map_x, map_y) = 0 Then
        Char_Create = Char_Next_Open
        Char_Make Char_Create, map_x, map_y, Heading, body_index, head_index, weapon_index, shield_index, helmet_index, Name, CharType, fx, FxLoopTimes, id, char_invisible, char_fly, particle_count, particle_life_counter(), particle_type()
    End If

End Function

Public Function Char_Change(ByVal id As Integer, Optional ByVal Heading As Byte, _
                            Optional ByVal body_index As Integer, Optional ByVal head_index As Integer, _
                            Optional ByVal helmet_index As Integer, Optional ByVal weapon_index As Integer, Optional ByVal shield_index As Integer, _
                            Optional ByVal fx As Integer, Optional ByVal FxLoopTimes As Integer) As Boolean
                            
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 6/9/2005
'**************************************************************
    Dim char_index As Integer
    
    'Invalid heading
    If Heading < 1 Or Heading > 8 Then Exit Function
    
    char_index = Char_Find(id)
    
    If Char_Check(char_index) Then
        Char_Data_Change char_index, Heading, body_index, head_index, weapon_index, shield_index, helmet_index, "", 0, fx, FxLoopTimes
    End If

End Function

Private Sub Char_Data_Change(ByVal char_index As Integer, Optional ByVal Heading As Byte, _
                        Optional ByVal body_index As Integer, Optional ByVal head_index As Integer, _
                        Optional ByVal weapon_index As Integer, Optional ByVal shield_index As Integer, Optional ByVal helmet_index As Integer, _
                        Optional ByVal Name As String, Optional ByVal CharType As Integer, Optional ByVal fx As Integer, Optional ByVal FxLoopTimes As Integer)
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 6/9/2005
'**************************************************************

    If Name <> "" Then CharList(char_index).nombre = Name
    If CharType <> 0 Then CharList(char_index).CharType = CharType
    If Heading <> 0 Then CharList(char_index).Heading = Heading
    
    CharList(char_index).iBody = body_index
    CharList(char_index).navigating = Char_Body_Ship_Get(body_index)
    
    CharList(char_index).Body = BodyData(body_index)
    CharList(char_index).Body.Walk(Heading).framecounter = 1
    CharList(char_index).Body.Walk(Heading).started = 0
    CharList(char_index).Muerto = (head_index = 500)
    
    CharList(char_index).iHead = head_index
    CharList(char_index).head = HeadData(head_index)

    If weapon_index <> 0 Then
        CharList(char_index).Arma = WeaponAnimData(weapon_index)
        CharList(char_index).Arma.WeaponWalk(Heading).framecounter = 1
        CharList(char_index).Arma.WeaponWalk(Heading).started = 0
    End If
    
    If shield_index <> 0 Then
        CharList(char_index).Escudo = ShieldAnimData(shield_index)
        CharList(char_index).Escudo.ShieldWalk(Heading).framecounter = 1
        CharList(char_index).Escudo.ShieldWalk(Heading).started = 0
    End If
    
    If helmet_index <> 0 Then CharList(char_index).Casco = CascoAnimData(helmet_index)

    If fx <> 0 Then CharList(char_index).fx = fx
    If FxLoopTimes <> 0 Then CharList(char_index).FxLoopTimes = FxLoopTimes

End Sub

Private Sub Char_Make(ByVal char_index As Integer, ByVal map_x As Integer, ByVal map_y As Integer, ByVal Heading As Byte, _
                        ByVal body_index As Integer, ByVal head_index As Integer, _
                        ByVal weapon_index As Integer, ByVal shield_index As Integer, ByVal helmet_index As Integer, _
                        ByVal Name As String, ByVal CharType As Integer, ByVal fx As Integer, ByVal FxLoopTimes As Integer, _
                        ByVal id As Integer, ByVal char_invisible As Byte, ByVal char_fly As Byte, ByVal particle_count As Integer, _
                        ByRef particle_life() As Long, ByRef particle_type() As Integer)
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Modified by Augusto José Rando
'Last Modify Date: 10/07/2002
'Makes a new character and puts it on the map
'*****************************************************************
    
    On Error GoTo ErrorHandler
    
    Dim i As Integer, file_num As Integer
    file_num = -1
    
    'Update array size
    If char_index > LastChar Then
        LastChar = char_index
        ReDim Preserve CharList(1 To LastChar)
    End If
    NumChars = NumChars + 1
    
    'Barrin: preloading of all characters data resources when needed
    If weapon_index = 0 Then
        weapon_index = 2
    ElseIf weapon_index <> 2 Then
        If GrhData(WeaponAnimData(weapon_index).WeaponWalk(Heading).GrhIndex).FileNum <> 0 Then
            If Textures.CargarGrafico(GrhData(WeaponAnimData(weapon_index).WeaponWalk(Heading).GrhIndex).FileNum) Then DoEvents
        End If
    End If
    
    If shield_index = 0 Then
        shield_index = 2
    ElseIf shield_index <> 2 Then
        If GrhData(ShieldAnimData(shield_index).ShieldWalk(Heading).GrhIndex).FileNum <> 0 Then
            If Textures.CargarGrafico(GrhData(ShieldAnimData(shield_index).ShieldWalk(Heading).GrhIndex).FileNum) Then DoEvents
        End If
    End If
    
    If helmet_index = 0 Then
        helmet_index = 2
    ElseIf helmet_index <> 2 Then
        If GrhData(CascoAnimData(helmet_index).head(Heading).GrhIndex).FileNum <> 0 Then
            If Textures.CargarGrafico(GrhData(CascoAnimData(helmet_index).head(Heading).GrhIndex).FileNum) Then DoEvents
        End If
    End If
    
    If head_index <> 0 Then
        If GrhData(HeadData(head_index).head(Heading).GrhIndex).FileNum <> 0 Then
            If Textures.CargarGrafico(GrhData(HeadData(head_index).head(Heading).GrhIndex).FileNum) Then DoEvents
        End If
    End If
    
    If body_index <> 0 Then
        If GrhData(BodyData(body_index).Walk(1).GrhIndex).FileNum <> 0 Then
            If GrhData(BodyData(body_index).Walk(1).GrhIndex).FileNum <> file_num Then
                file_num = GrhData(BodyData(body_index).Walk(1).GrhIndex).FileNum
                If Textures.CargarGrafico(file_num) Then DoEvents
            End If
        End If
        
        If GrhData(BodyData(body_index).Walk(3).GrhIndex).FileNum <> 0 Then
            If GrhData(BodyData(body_index).Walk(3).GrhIndex).FileNum <> file_num Then
                file_num = GrhData(BodyData(body_index).Walk(3).GrhIndex).FileNum
                If Textures.CargarGrafico(file_num) Then DoEvents
            End If
        End If

        If GrhData(BodyData(body_index).Walk(5).GrhIndex).FileNum <> 0 Then
            If GrhData(BodyData(body_index).Walk(5).GrhIndex).FileNum <> file_num Then
                file_num = GrhData(BodyData(body_index).Walk(5).GrhIndex).FileNum
                If Textures.CargarGrafico(file_num) Then DoEvents
            End If
        End If

        If GrhData(BodyData(body_index).Walk(7).GrhIndex).FileNum <> 0 Then
            If GrhData(BodyData(body_index).Walk(7).GrhIndex).FileNum <> file_num Then
                file_num = GrhData(BodyData(body_index).Walk(7).GrhIndex).FileNum
                If Textures.CargarGrafico(file_num) Then DoEvents
            End If
        End If
    End If
    
    'Heading
    CharList(char_index).Heading = Heading
    
    'Update char data
    CharList(char_index).nombre = Name
    CharList(char_index).id = id
    CharList(char_index).CharType = CharType
        
    CharList(char_index).iBody = body_index
    CharList(char_index).navigating = Char_Body_Ship_Get(body_index)
    
    CharList(char_index).iHead = head_index

    CharList(char_index).head = HeadData(head_index)
    CharList(char_index).Body = BodyData(body_index)
    CharList(char_index).Arma = WeaponAnimData(weapon_index)
    CharList(char_index).Arma.WeaponAttack = 0
    CharList(char_index).Escudo = ShieldAnimData(shield_index)
    CharList(char_index).Casco = CascoAnimData(helmet_index)

    CharList(char_index).Body.Walk(Heading).framecounter = 1
    CharList(char_index).Body.Walk(Heading).started = 0
    CharList(char_index).Arma.WeaponWalk(Heading).framecounter = 1
    CharList(char_index).Arma.WeaponWalk(Heading).started = 0
    CharList(char_index).Escudo.ShieldWalk(Heading).framecounter = 1
    CharList(char_index).Escudo.ShieldWalk(Heading).started = 0
    
    CharList(char_index).fx = fx
    CharList(char_index).FxLoopTimes = FxLoopTimes

    CharList(char_index).invisible = char_invisible
    CharList(char_index).fly = char_fly
    CharList(char_index).Muerto = (head_index = 500)

    CharList(char_index).particle_count = particle_count

    If particle_count > 0 Then
        ReDim CharList(char_index).particle_group(1 To particle_count) As Long
        For i = 1 To particle_count
            CharList(char_index).particle_group(i) = General_Char_Particle_Create(particle_type(i), char_index, particle_life(i))
        Next i
    End If

    'Label Offset
    'CharList(char_index).label_offset_x = char_data_list(char_data_index).label_offset_x
    'CharList(char_index).label_offset_y = char_data_list(char_data_index).label_offset_y
    
    'Reset moving stats
    CharList(char_index).scroll_on = False
    CharList(char_index).scroll_direction_x = 0
    CharList(char_index).scroll_direction_y = 0
    CharList(char_index).scroll_offset_counter_y = 0
    CharList(char_index).scroll_offset_counter_x = 0
    
    'Update position
    CharList(char_index).Pos.x = map_x
    CharList(char_index).Pos.y = map_y
        
    'Plot on map
    MapData(map_x, map_y).CharIndex = char_index
    
    Exit Sub
    
ErrorHandler:
    Char_Remove (char_index)
    
End Sub

'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 8/06/2005
'
'**************************************************************
Private Function Char_Alpha_View(TempChar As char, UserChar As char) As Boolean

Dim TempCharClanName As String, UserCharClanName As String

TempCharClanName = General_Field_Read(2, TempChar.nombre, "<")
UserCharClanName = General_Field_Read(2, UserChar.nombre, "<")

If (UserChar.nombre = TempChar.nombre) Then
    Char_Alpha_View = True
    Exit Function
ElseIf TempChar.iHead = 500 And (TempChar.iBody = 8 Or TempChar.iBody = 145) Then
    Char_Alpha_View = True
    Exit Function
ElseIf ((UserChar.CharType = 0 And TempChar.CharType = 0) Or UserChar.CharType = 4) Then
    Char_Alpha_View = True
    Exit Function
ElseIf (UserCharClanName <> "") And (UserCharClanName = TempCharClanName) Then
    Char_Alpha_View = True
    Exit Function
ElseIf (UserChar.group_index > 0) And (TempChar.group_index = UserChar.group_index) Then
    Char_Alpha_View = True
    Exit Function
End If

Char_Alpha_View = False

End Function

Private Sub Char_Render(ByRef temp_char As char, ByVal char_index As Integer, ByVal screen_x As Integer, ByVal screen_y As Integer, ByRef light_value() As Long, Optional ByVal alpha_blend As Boolean = False)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 8/05/2004
'Renders a character at a paticular screen point
'Edited By Juan Martín Sotuyo Dodero
'*****************************************************************
    
    Dim head As Boolean
    Dim i As Integer
    
    If temp_char.Heading = 0 Then Exit Sub

    'Scrolling
    If temp_char.scroll_on Then
        '****** Move Left and Right if needed ******
        If temp_char.scroll_direction_x <> 0 Then
            temp_char.scroll_offset_counter_x = temp_char.scroll_offset_counter_x + (IIf(char_index = char_current, scroll_pixels_per_frame, 4.3) * timer_ticks_per_frame * Sgn(temp_char.scroll_direction_x))
            If Sgn(temp_char.scroll_offset_counter_x) = temp_char.scroll_direction_x Then
                temp_char.scroll_offset_counter_x = 0
                temp_char.scroll_direction_x = 0
            End If
        End If
        '****** Move Up and Down if needed ******
        If temp_char.scroll_direction_y <> 0 Then
            temp_char.scroll_offset_counter_y = temp_char.scroll_offset_counter_y + (IIf(char_index = char_current, scroll_pixels_per_frame, 4.3) * timer_ticks_per_frame * Sgn(temp_char.scroll_direction_y))
            If Sgn(temp_char.scroll_offset_counter_y) = temp_char.scroll_direction_y Then
                temp_char.scroll_offset_counter_y = 0
                temp_char.scroll_direction_y = 0
            End If
        End If
        'End scrolling if needed
        If temp_char.scroll_direction_x = 0 And temp_char.scroll_direction_y = 0 Then
            'Turn off scrolling
            temp_char.scroll_on = False
            'Set char to stand
            temp_char.Body.Walk(temp_char.Heading).framecounter = 1
            temp_char.Body.Walk(temp_char.Heading).started = 0
            temp_char.Arma.WeaponWalk(temp_char.Heading).framecounter = 1
            temp_char.Arma.WeaponWalk(temp_char.Heading).started = 0
            temp_char.Escudo.ShieldWalk(temp_char.Heading).framecounter = 1
            temp_char.Escudo.ShieldWalk(temp_char.Heading).started = 0
        Else
            temp_char.Body.Walk(temp_char.Heading).started = 1
            temp_char.Arma.WeaponWalk(temp_char.Heading).started = 1
            temp_char.Escudo.ShieldWalk(temp_char.Heading).started = 1
        End If
    End If
    
    head = (temp_char.head.head(temp_char.Heading).GrhIndex > 0)

    'Find screen position
    screen_x = screen_x + temp_char.scroll_offset_counter_x
    screen_y = screen_y + temp_char.scroll_offset_counter_y
    
    If temp_char.Body.Walk(temp_char.Heading).GrhIndex Then
        Grh_Render temp_char.Body.Walk(temp_char.Heading), screen_x, screen_y, light_value(), True, True, alpha_blend
                                                                                        
        If temp_char.nombre <> "" And engine_label_render And (Not char_current_blind) Then
            If head Or temp_char.navigating Then
                Char_Label_Render temp_char.nombre, temp_char.CharType, screen_x, screen_y
            End If
        End If
    End If
    
    If head Then
                
        If temp_char.Casco.head(temp_char.Heading).GrhIndex <> 0 Then
            Grh_Render temp_char.head.head(temp_char.Heading), screen_x + temp_char.Body.HeadOffset.x, screen_y + temp_char.Body.HeadOffset.y, light_value(), True, True, alpha_blend
            Grh_Render temp_char.Casco.head(temp_char.Heading), screen_x + temp_char.Body.HeadOffset.x, screen_y + temp_char.Body.HeadOffset.y, light_value(), True, True, alpha_blend
        Else
            Grh_Render temp_char.head.head(temp_char.Heading), screen_x + temp_char.Body.HeadOffset.x, screen_y + temp_char.Body.HeadOffset.y, light_value(), True, True, alpha_blend
        End If
        If temp_char.Arma.WeaponWalk(temp_char.Heading).GrhIndex <> 0 Then _
            Grh_Render temp_char.Arma.WeaponWalk(temp_char.Heading), screen_x, screen_y, light_value(), True, True, alpha_blend
        If temp_char.Escudo.ShieldWalk(temp_char.Heading).GrhIndex <> 0 Then _
            Grh_Render temp_char.Escudo.ShieldWalk(temp_char.Heading), screen_x, screen_y, light_value(), True, True, alpha_blend
    End If
    
    If temp_char.particle_count > 0 Then
        For i = 1 To temp_char.particle_count
            If temp_char.particle_group(i) > 0 Then
                Particle_Group_Render temp_char.particle_group(i), screen_x + (base_tile_size / 2), screen_y
            End If
        Next i
    End If
    
    If temp_char.fx <> 0 Then
        Grh_Render_Char_FX FxData(temp_char.fx).fx, screen_x + FxData(temp_char.fx).offsetX, screen_y + FxData(temp_char.fx).OffsetY, light_value(), temp_char, True, True, True
    End If
    
End Sub

Private Sub Char_Invisible_Render(ByRef temp_char As char, ByVal char_index As Integer, ByVal screen_x As Integer, ByVal screen_y As Integer, ByRef light_value() As Long, Optional ByVal alpha_blend As Boolean = False)
'*****************************************************************
'Author: Augusto José Rando
'Last Modify Date: 8/05/2004
'Renders a invisible character at a paticular screen point
'*****************************************************************
    
    Dim i As Long
    
    If temp_char.Heading = 0 Then Exit Sub

    'Scrolling
    If temp_char.scroll_on Then
        '****** Move Left and Right if needed ******
        If temp_char.scroll_direction_x <> 0 Then
            temp_char.scroll_offset_counter_x = temp_char.scroll_offset_counter_x + (IIf(char_index = char_current, scroll_pixels_per_frame, 4.3) * timer_ticks_per_frame * Sgn(temp_char.scroll_direction_x))
            If Sgn(temp_char.scroll_offset_counter_x) = temp_char.scroll_direction_x Then
                temp_char.scroll_offset_counter_x = 0
                temp_char.scroll_direction_x = 0
            End If
        End If
        '****** Move Up and Down if needed ******
        If temp_char.scroll_direction_y <> 0 Then
            temp_char.scroll_offset_counter_y = temp_char.scroll_offset_counter_y + (IIf(char_index = char_current, scroll_pixels_per_frame, 4.3) * timer_ticks_per_frame * Sgn(temp_char.scroll_direction_y))
            If Sgn(temp_char.scroll_offset_counter_y) = temp_char.scroll_direction_y Then
                temp_char.scroll_offset_counter_y = 0
                temp_char.scroll_direction_y = 0
            End If
        End If
        'End scrolling if needed
        If temp_char.scroll_direction_x = 0 And temp_char.scroll_direction_y = 0 Then
            'Turn off scrolling
            temp_char.scroll_on = False
            'Set char to stand
            temp_char.Body.Walk(temp_char.Heading).framecounter = 1
            temp_char.Body.Walk(temp_char.Heading).started = 0
            temp_char.Arma.WeaponWalk(temp_char.Heading).framecounter = 1
            temp_char.Arma.WeaponWalk(temp_char.Heading).started = 0
            temp_char.Escudo.ShieldWalk(temp_char.Heading).framecounter = 1
            temp_char.Escudo.ShieldWalk(temp_char.Heading).started = 0
        Else
            temp_char.Body.Walk(temp_char.Heading).started = 1
            temp_char.Arma.WeaponWalk(temp_char.Heading).started = 1
            temp_char.Escudo.ShieldWalk(temp_char.Heading).started = 1
        End If
    End If

    'Find screen position
    screen_x = screen_x + temp_char.scroll_offset_counter_x
    screen_y = screen_y + temp_char.scroll_offset_counter_y
        
    If temp_char.particle_count > 0 Then
        For i = 1 To temp_char.particle_count
            If temp_char.particle_group(i) > 0 Then
                Particle_Group_Render temp_char.particle_group(i), screen_x + (base_tile_size / 2), screen_y
            End If
        Next i
    End If
    
    If temp_char.fx <> 0 Then
        Grh_Render_Char_FX FxData(temp_char.fx).fx, screen_x + FxData(temp_char.fx).offsetX, screen_y + FxData(temp_char.fx).OffsetY, light_value(), temp_char, True, True, True
    End If
    
End Sub

'Creado por Barrin
Private Function Char_Label_Render(ByVal TempCharName As String, ByVal TempCharType As Byte, ByVal screen_x As Integer, ByVal screen_y As Integer) As Boolean

Dim lCenter As Long
Dim sClan As String
Dim tempa(3) As Long
Dim tempb(3) As Long
Dim a As Integer

If InStr(TempCharName, "<") > 0 And InStr(TempCharName, ">") > 0 Then
    
    sClan = mid(TempCharName, InStr(TempCharName, "<"))
    If TempCharType = 1 Then
        For a = 0 To 3
            tempa(a) = -65536
        Next a
    ElseIf TempCharType = 2 Then
        For a = 0 To 3
            tempa(a) = -16744193
        Next a
    ElseIf TempCharType = 3 Then
        For a = 0 To 3
            tempa(a) = -7436129
        Next a
    Else
        For a = 0 To 3
            tempa(a) = -16729856
        Next a
    End If
    a = screen_x + 16 - Engine_Text_Width(left(TempCharName, InStr(TempCharName, "<") - 2)) / 2
    Engine_Text_Render left(TempCharName, InStr(TempCharName, "<") - 1), a, screen_y + 30, tempa
    a = screen_x + 16 - Engine_Text_Width(sClan) / 2
    Engine_Text_Render sClan, a, screen_y + 45, tempa
Else
    screen_x = screen_x + 16 - Engine_Text_Width(TempCharName) / 2
    If TempCharType = 1 Then
        For a = 0 To 3
            tempa(a) = -65536
        Next a
    ElseIf TempCharType = 2 Then
        For a = 0 To 3
            tempa(a) = -16744193
        Next a
    ElseIf TempCharType = 3 Then
        For a = 0 To 3
            tempa(a) = -7436129
        Next a
    Else
        For a = 0 To 3
            tempa(a) = -16729856
        Next a
    End If
    Engine_Text_Render TempCharName, screen_x, screen_y + 30, tempa
End If

End Function

Public Function Light_Remove(ByVal light_index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Light_Check(light_index) Then
        Light_Destroy light_index
        Light_Remove = True
    End If
End Function

Public Function Light_Color_Value_Get(ByVal light_index As Long, ByRef color_value As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Light_Check(light_index) Then
        color_value = light_list(light_index).color
        Light_Color_Value_Get = True
    End If
End Function

Public Function Light_Create(ByVal map_x As Integer, ByVal map_y As Integer, Optional ByVal color_value As Long = &HFFFFFF, _
                            Optional ByVal range As Byte = 1, Optional ByVal id As Long) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Returns the light_index if successful, else 0
'Edited by Juan Martín Sotuyo Dodero
'**************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        'Make sure there is no light in the given map pos
        If Map_Light_Get(map_x, map_y) <> 0 Then
            Light_Create = 0
            Exit Function
        End If
        Light_Create = Light_Next_Open
        Light_Make Light_Create, map_x, map_y, color_value, range, id
    End If
End Function

Public Function Light_Move(ByVal light_index As Long, ByVal map_x As Integer, ByVal map_y As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Returns true if successful, else false
'**************************************************************
    'Make sure it's a legal CharIndex
    If Light_Check(light_index) Then
        'Make sure it's a legal move
        If Map_In_Bounds(map_x, map_y) Then
        
            'Move it
            Light_Erase light_index
            light_list(light_index).map_x = map_x
            light_list(light_index).map_y = map_y
    
            Light_Move = True
            
        End If
    End If
End Function

Public Function Light_Move_By_Head(ByVal light_index As Long, ByVal Heading As Byte) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 15/05/2002
'Returns true if successful, else false
'**************************************************************
    Dim map_x As Integer
    Dim map_y As Integer
    Dim nX As Integer
    Dim nY As Integer
    
    'Check for valid heading
    If Heading < 1 Or Heading > 8 Then
        Light_Move_By_Head = False
        Exit Function
    End If

    'Make sure it's a legal CharIndex
    If Light_Check(light_index) Then
    
        map_x = light_list(light_index).map_x
        map_y = light_list(light_index).map_y
        
        nX = map_x
        nY = map_y
        
        Convert_Heading_to_Direction Heading, nX, nY
        
        'Make sure it's a legal move
        If Map_In_Bounds(nX, nY) Then
        
            'Move it
            Light_Erase light_index

            light_list(light_index).map_x = nX
            light_list(light_index).map_y = nY
    
            Light_Move_By_Head = True
            
        End If
    End If
End Function

Private Sub Light_Make(ByVal light_index As Long, ByVal map_x As Integer, ByVal map_y As Integer, ByVal rgb_value As Long, _
                        ByVal range As Long, Optional ByVal id As Long)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
    'Update array size
    If light_index > light_last Then
        light_last = light_index
        ReDim Preserve light_list(1 To light_last)
    End If
    light_count = light_count + 1
    
    'Make active
    light_list(light_index).active = True
    
    light_list(light_index).map_x = map_x
    light_list(light_index).map_y = map_y
    light_list(light_index).color = rgb_value
    light_list(light_index).range = range
    light_list(light_index).id = id
End Sub

Private Function Light_Check(ByVal light_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check light_index
    If light_index > 0 And light_index <= light_last Then
        If light_list(light_index).active Then
            Light_Check = True
        End If
    End If
End Function

Private Sub Light_Render_All()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim loop_counter As Long
            
    For loop_counter = 1 To light_count
        
        If light_list(loop_counter).active Then
            Light_Render loop_counter
        End If
    
    Next loop_counter
End Sub

Private Sub Light_Render(ByVal light_index As Long)
'***************************************'
'Author: Juan Martín Sotuyo Dodero
'Last modified: 11/11/2004
'Renders a light
'***************************************'
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim x As Integer
    Dim y As Integer
    Dim color As Long
    
    'Set up light borders
    min_x = light_list(light_index).map_x - light_list(light_index).range
    min_y = light_list(light_index).map_y - light_list(light_index).range
    max_x = light_list(light_index).map_x + light_list(light_index).range
    max_y = light_list(light_index).map_y + light_list(light_index).range
    
    'Set color
    color = light_list(light_index).color
    
    'Arrange corners
    'NE
    If Map_In_Bounds(min_x, min_y) Then
        MapData(min_x, min_y).light_value(2) = color
    End If
    'NW
    If Map_In_Bounds(max_x, min_y) Then
        MapData(max_x, min_y).light_value(0) = color
    End If
    'SW
    If Map_In_Bounds(max_x, max_y) Then
        MapData(max_x, max_y).light_value(1) = color
    End If
    'SE
    If Map_In_Bounds(min_x, max_y) Then
        MapData(min_x, max_y).light_value(3) = color
    End If
    
    'Arrange borders
    'Upper border
    For x = min_x + 1 To max_x - 1
        If Map_In_Bounds(x, min_y) Then
            MapData(x, min_y).light_value(0) = color
            MapData(x, min_y).light_value(2) = color
        End If
    Next x
    
    'Lower border
    For x = min_x + 1 To max_x - 1
        If Map_In_Bounds(x, max_y) Then
            MapData(x, max_y).light_value(1) = color
            MapData(x, max_y).light_value(3) = color
        End If
    Next x
    
    'Left border
    For y = min_y + 1 To max_y - 1
        If Map_In_Bounds(min_x, y) Then
            MapData(min_x, y).light_value(2) = color
            MapData(min_x, y).light_value(3) = color
        End If
    Next y
    
    'Right border
    For y = min_y + 1 To max_y - 1
        If Map_In_Bounds(max_x, y) Then
            MapData(max_x, y).light_value(0) = color
            MapData(max_x, y).light_value(1) = color
        End If
    Next y
    
    'Set the inner part of the light
    For x = min_x + 1 To max_x - 1
        For y = min_y + 1 To max_y - 1
            If Map_In_Bounds(x, y) Then
                MapData(x, y).light_value(0) = color
                MapData(x, y).light_value(1) = color
                MapData(x, y).light_value(2) = color
                MapData(x, y).light_value(3) = color
            End If
        Next y
    Next x
End Sub

Private Function Light_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until light_list(loopc).active = False
        If loopc = light_last Then
            Light_Next_Open = light_last + 1
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Light_Next_Open = loopc
Exit Function
ErrorHandler:
    Light_Next_Open = 1
End Function

Public Function Light_Find(ByVal id As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until light_list(loopc).id = id
        If loopc = light_last Then
            Light_Find = 0
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Light_Find = loopc
Exit Function
ErrorHandler:
    Light_Find = 0
End Function

Public Function Light_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim Index As Long
    
    For Index = 1 To light_last
        'Make sure it's a legal index
        If Light_Check(Index) Then
            Light_Destroy Index
        End If
    Next Index
    
    Light_Remove_All = True
End Function

Private Sub Light_Destroy(ByVal light_index As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim temp As Light
    
    Light_Erase light_index
    
    light_list(light_index) = temp
    
    'Update array size
    If light_index = light_last Then
        Do Until light_list(light_last).active
            light_last = light_last - 1
            If light_last = 0 Then
                light_count = 0
                Exit Sub
            End If
        Loop
        ReDim Preserve light_list(1 To light_last)
    End If
    light_count = light_count - 1
End Sub

Private Sub Light_Erase(ByVal light_index As Long)
'***************************************'
'Author: Juan Martín Sotuyo Dodero
'Last modified: 3/31/2003
'Correctly erases a light
'***************************************'
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim x As Integer
    Dim y As Integer
    
    'Set up light borders
    min_x = light_list(light_index).map_x - light_list(light_index).range
    min_y = light_list(light_index).map_y - light_list(light_index).range
    max_x = light_list(light_index).map_x + light_list(light_index).range
    max_y = light_list(light_index).map_y + light_list(light_index).range
    
    'Arrange corners
    'NE
    If Map_In_Bounds(min_x, min_y) Then
        MapData(min_x, min_y).light_value(2) = map_base_light
    End If
    'NW
    If Map_In_Bounds(max_x, min_y) Then
        MapData(max_x, min_y).light_value(0) = map_base_light
    End If
    'SW
    If Map_In_Bounds(max_x, max_y) Then
        MapData(max_x, max_y).light_value(1) = map_base_light
    End If
    'SE
    If Map_In_Bounds(min_x, max_y) Then
        MapData(min_x, max_y).light_value(3) = map_base_light
    End If
    
    'Arrange borders
    'Upper border
    For x = min_x + 1 To max_x - 1
        If Map_In_Bounds(x, min_y) Then
            MapData(x, min_y).light_value(0) = map_base_light
            MapData(x, min_y).light_value(2) = map_base_light
        End If
    Next x
    
    'Lower border
    For x = min_x + 1 To max_x - 1
        If Map_In_Bounds(x, max_y) Then
            MapData(x, max_y).light_value(1) = map_base_light
            MapData(x, max_y).light_value(3) = map_base_light
        End If
    Next x
    
    'Left border
    For y = min_y + 1 To max_y - 1
        If Map_In_Bounds(min_x, y) Then
            MapData(min_x, y).light_value(2) = map_base_light
            MapData(min_x, y).light_value(3) = map_base_light
        End If
    Next y
    
    'Right border
    For y = min_y + 1 To max_y - 1
        If Map_In_Bounds(max_x, y) Then
            MapData(max_x, y).light_value(0) = map_base_light
            MapData(max_x, y).light_value(1) = map_base_light
        End If
    Next y
    
    'Set the inner part of the light
    For x = min_x + 1 To max_x - 1
        For y = min_y + 1 To max_y - 1
            If Map_In_Bounds(x, y) Then
                MapData(x, y).light_value(0) = map_base_light
                MapData(x, y).light_value(1) = map_base_light
                MapData(x, y).light_value(2) = map_base_light
                MapData(x, y).light_value(3) = map_base_light
            End If
        Next y
    Next x
End Sub

Public Function Map_FileNum_Get(ByVal map_x As Integer, ByVal map_y As Integer, ByVal layer As Integer) As Integer
'*****************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 10/6/2005
'*****************************************************************
    On Error GoTo Error_Handler
    
    If layer < 1 Or layer > 4 Then
        Map_FileNum_Get = 0
        Exit Function
    End If
    
    If Map_In_Bounds(map_x, map_y) Then
        If MapData(map_x, map_y).Graphic(layer).GrhIndex > 0 Then
            Map_FileNum_Get = GrhData(MapData(map_x, map_y).Graphic(layer).GrhIndex).FileNum
        Else
            Map_FileNum_Get = 0
        End If
    End If

    Exit Function
    
Error_Handler:
    Map_FileNum_Get = 0

End Function

Public Function Map_Grh_Info_Get(ByVal map_x As Integer, ByVal map_y As Integer, ByVal layer As Integer, ByRef alpha_blending As Boolean, ByRef angle As Single, ByRef h_centered As Boolean, ByRef v_centered As Boolean) As Long
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modify Date: 12/10/2003
'If the tile has a Grh returns it´s index, and sets the angle
'and alpha_blending status to the given variables
'*****************************************************************
    'Check
    If layer < 1 Or layer > 4 Then
        Map_Grh_Info_Get = 0
        Exit Function
    End If
    
    If Map_In_Bounds(map_x, map_y) Then
        Map_Grh_Info_Get = MapData(map_x, map_y).Graphic(layer).GrhIndex
        angle = MapData(map_x, map_y).Graphic(layer).angle
        alpha_blending = MapData(map_x, map_y).Graphic(layer).alpha_blend
        'h_centered = MapData(map_x, map_y).Graphic(layer).h_centered
        'v_centered = MapData(map_x, map_y).Graphic(layer).v_centered
    Else
        Map_Grh_Info_Get = 0
    End If
End Function
 
Public Function Light_Info_Get(ByVal light_index As Long, ByRef map_x As Integer, ByRef map_y As Integer, ByRef color As Long, ByRef range As Byte) As Boolean
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modify Date: 7/19/2003
'Gets the data of the given light
'*****************************************************************
    If Not Light_Check(light_index) Then
        Light_Info_Get = False
        Exit Function
    End If
    
    map_x = light_list(light_index).map_x
    map_y = light_list(light_index).map_y
    color = light_list(light_index).color
    range = light_list(light_index).range
    
    Light_Info_Get = True
End Function

Public Function Light_Count_Get() As Long
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modify Date: 7/19/2003
'Gets the number of existing lights
'*****************************************************************
    Light_Count_Get = UBound(light_list)
End Function

Private Sub Grh_Render_Advance(ByRef Grh As Grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByVal Height As Integer, ByVal Width As Integer, ByRef rgb_list() As Long, Optional ByVal h_center As Boolean, Optional ByVal v_center As Boolean, Optional ByVal alpha_blend As Boolean = False)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modify Date: 11/19/2003
'Similar to Grh_Render, but let´s you resize the Grh
'**************************************************************
    Dim tile_width As Integer
    Dim tile_height As Integer
    Dim grh_index As Long
    
    'Animation
    If Grh.started Then
        Grh.framecounter = Grh.framecounter + (timer_ticks_per_frame * Grh.speed)
        If Grh.framecounter > GrhData(Grh.GrhIndex).NumFrames Then
            'If Grh.noloop Then
            '    Grh.FrameCounter = GrhData(Grh.GrhIndex).NumFrames
            'Else
                Grh.framecounter = 1
            'End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    If Grh.framecounter = 0 Then Grh.framecounter = 1
    grh_index = GrhData(Grh.GrhIndex).Frames(Grh.framecounter)
    
    'Center Grh over X, Y pos
    If GrhData(Grh.GrhIndex).TileWidth <> 1 Then
        screen_x = screen_x - Int(GrhData(Grh.GrhIndex).TileWidth * (base_tile_size \ 2)) + base_tile_size \ 2
    End If
    
    If GrhData(Grh.GrhIndex).TileHeight <> 1 Then
        screen_y = screen_y - Int(GrhData(Grh.GrhIndex).TileHeight * base_tile_size) + base_tile_size
    End If
    
    'Draw it to device
    Device_Box_Textured_Render_Advance grh_index, _
        screen_x, screen_y, _
        GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, _
        rgb_list, _
        GrhData(grh_index).sX, GrhData(grh_index).sY, _
        Width, Height, alpha_blend, Grh.angle
End Sub

Public Sub GUI_Box_Filled_Render(ByVal x As Integer, ByVal y As Integer, ByVal Width As Integer, ByVal Height As Integer, _
                            ByVal Color1 As Long, Optional ByVal Color2 As Long, Optional ByVal Color3 As Long, _
                            Optional ByVal Color4 As Long, Optional alpha_blend As Boolean, Optional ByVal angle As Single)
'**************************************************************
'Author: Fredrik Alexandersson - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 5/15/2003
'Draws a filled box
'**************************************************************
    If Color2 = 0 Then Color2 = Color1
    If Color3 = 0 Then Color3 = Color1
    If Color4 = 0 Then Color4 = Color1
    
    Static Vertex(3) As TLVERTEX
    Static box_rect As RECT
    Static rgb_list(3) As Long
    
    With box_rect
        .bottom = y + Height
        .left = x
        .Right = x + Width
        .top = y
    End With
    
    rgb_list(0) = Color1
    rgb_list(1) = Color2
    rgb_list(2) = Color3
    rgb_list(3) = Color4
    
    Geometry_Create_Box Vertex(), box_rect, box_rect, rgb_list(), , angle
    
    If alpha_blend Then
       'Set Rendering for alphablending
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    ddevice.SetTexture 0, Nothing
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex(0), Len(Vertex(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
End Sub

Public Sub GUI_Box_Outline_Render(ByVal x As Integer, ByVal y As Integer, ByVal Width As Integer, ByVal Height As Integer, _
                            ByVal border_width As Integer, ByVal Color1 As Long, Optional ByVal Color2 As Long, Optional ByVal Color3 As Long, _
                            Optional ByVal Color4 As Long, Optional alpha_blend As Boolean)
'**************************************************************
'Author: Fredrik Alexandersson - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 5/15/2003
'Draws a box outline
'**************************************************************
    If Color2 = 0 Then Color2 = Color1
    If Color3 = 0 Then Color3 = Color1
    If Color4 = 0 Then Color4 = Color1
    
    Static VertexU(3) As TLVERTEX
    Static VertexL(3) As TLVERTEX
    Static VertexR(3) As TLVERTEX
    Static VertexB(3) As TLVERTEX
    Static box_rect As RECT
    
    With box_rect
        .bottom = y + Height
        .left = x
        .Right = x + Width
        .top = y
    End With
    
    If alpha_blend Then
       'Set Rendering for alphablending
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    ddevice.SetTexture 0, Nothing
    'Upper Line
    VertexU(0) = Geometry_Create_TLVertex(box_rect.left, box_rect.top, 0, 1, Color1, 0, 0, 0)
    VertexU(1) = Geometry_Create_TLVertex(box_rect.Right, box_rect.top, 0, 1, Color1, 0, 0, 0)
    VertexU(2) = Geometry_Create_TLVertex(box_rect.left, box_rect.top + border_width, 0, 1, Color1, 0, 0, 0)
    VertexU(3) = Geometry_Create_TLVertex(box_rect.Right, box_rect.top + border_width, 0, 1, Color1, 0, 0, 0)
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexU(0), Len(VertexU(0))
    'Left Line
    VertexL(0) = Geometry_Create_TLVertex(box_rect.left + border_width, box_rect.top, 0, 1, Color2, 0, 0, 0)
    VertexL(1) = Geometry_Create_TLVertex(box_rect.left + border_width, box_rect.bottom, 0, 1, Color2, 0, 0, 0)
    VertexL(2) = Geometry_Create_TLVertex(box_rect.left, box_rect.top, 0, 2, Color1, 0, 0, 0)
    VertexL(3) = Geometry_Create_TLVertex(box_rect.left, box_rect.bottom, 0, 2, Color1, 0, 0, 0)
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexL(0), Len(VertexL(0))
    'Right Border
    VertexR(0) = Geometry_Create_TLVertex(box_rect.Right, box_rect.top, 0, 1, Color3, 0, 0, 0)
    VertexR(1) = Geometry_Create_TLVertex(box_rect.Right, box_rect.bottom, 0, 1, Color3, 0, 0, 0)
    VertexR(2) = Geometry_Create_TLVertex(box_rect.Right - border_width, box_rect.top, 0, 3, Color1, 0, 0, 0)
    VertexR(3) = Geometry_Create_TLVertex(box_rect.Right - border_width, box_rect.bottom, 0, 3, Color1, 0, 0, 0)
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexR(0), Len(VertexR(0))
    'Bottom Border
    VertexB(0) = Geometry_Create_TLVertex(box_rect.left, box_rect.bottom - border_width, 0, 1, Color4, 0, 0, 0)
    VertexB(1) = Geometry_Create_TLVertex(box_rect.Right, box_rect.bottom - border_width, 0, 1, Color4, 0, 0, 0)
    VertexB(2) = Geometry_Create_TLVertex(box_rect.left, box_rect.bottom, 0, 1, Color4, 0, 0, 0)
    VertexB(3) = Geometry_Create_TLVertex(box_rect.Right, box_rect.bottom, 0, 1, Color4, 0, 0, 0)
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexB(0), Len(VertexB(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
End Sub

Public Sub GUI_Text_Render(ByVal Text As String, ByVal font_index As Integer, ByVal x As Integer, _
                                ByVal y As Integer, ByVal Width As Integer, _
                                ByVal Height As Integer, ByVal color As Long, ByVal Alignment As FontAlignment, Optional ByVal word_wrap As Boolean)
'**************************************************************
'Author: Aaron Perkins - Modified by juan Martín Sotuyo Dodero
'Last Modify Date: 5/15/2003
'
'**************************************************************
    If Not Font_Check(font_index) Then Exit Sub
    
    If word_wrap Then
        Alignment = Alignment Or DT_WORDBREAK Or DT_EDITCONTROL
    End If
    
    Device_Text_Render font_list(font_index), Text, y, x, Width, Height, color, Alignment
End Sub

Public Sub GUI_Grh_Render(ByVal grh_index As Long, ByVal x As Integer, ByVal y As Integer, Optional ByVal Color1 As Long = &HFFFFFFFF, _
                        Optional ByVal Color2 As Long = &HFFFFFFFF, Optional ByVal Color3 As Long = &HFFFFFFFF, Optional ByVal Color4 As Long = &HFFFFFFFF, _
                        Optional ByVal Width As Integer, Optional ByVal Height As Integer, Optional ByVal angle As Single, Optional ByVal alpha_blend As Boolean)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/15/2003
'
'**************************************************************
    Dim temp_grh As Grh
    Dim rgb_list(3) As Long
    
    If Not Grh_Check(grh_index) Then Exit Sub
    
    rgb_list(0) = Color1
    rgb_list(1) = Color2
    rgb_list(2) = Color3
    rgb_list(3) = Color4
    
    Grh_Initialize temp_grh, grh_index, alpha_blend, angle
    
    If Height = 0 Then Height = GrhData(temp_grh.GrhIndex).pixelHeight
    If Width = 0 Then Width = GrhData(temp_grh.GrhIndex).pixelWidth
    
    Grh_Render_Advance temp_grh, x, y, Width, Height, rgb_list
End Sub

Public Function GUI_Box_Filled_Render_Advance(ByVal x As Integer, ByVal y As Integer, ByVal Width As Integer, ByVal Height As Integer, _
                            ByVal Color1 As Long, Optional ByVal Color2 As Long, Optional ByVal Color3 As Long, _
                            Optional ByVal Color4 As Long, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single) As Boolean
'**************************************************************
'Author: Juan MArtín Sotuyo Dodero
'Last Modify Date: 2/4/2004
'Draws a filled box with  some extra options
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
    If Not Color2 <> 0 Then
        Color2 = Color1
    End If
    If Not Color3 <> 0 Then
        Color3 = Color1
    End If
    If Not Color4 <> 0 Then
        Color4 = Color1
    End If
    
    Static Vertex(3) As TLVERTEX
    Static box_rect As RECT
    Static rgb_list(1 To 4) As Long
    
    With box_rect
        .bottom = y + Height
        .left = x
        .Right = x + Width
        .top = y
    End With
    
    rgb_list(1) = Color1
    rgb_list(2) = Color2
    rgb_list(3) = Color3
    rgb_list(4) = Color4
    
    'Add angle
    Geometry_Create_Box Vertex(), box_rect, box_rect, rgb_list(), , angle
    
    If alpha_blend Then
       'Set Rendering for alphablending
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    ddevice.SetTexture 0, Nothing
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex(0), Len(Vertex(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
        
    GUI_Box_Filled_Render_Advance = True
End Function

Public Function Engine_Resolution_Change(ByVal windowed As Boolean, Optional ByVal s_width As Integer, Optional ByVal s_height As Integer, _
                                        Optional ByVal v_left As Integer = 0, Optional ByVal v_top As Integer = 0, Optional ByVal v_width_in_tiles As Integer = 0, _
                                        Optional ByVal v_height_in_tiles As Integer = 0) As Boolean
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/29/2004
'Changes the resolution while in-game
'*****************************************************************
On Error GoTo ErrHandler
    If v_width_in_tiles = 0 Or v_height_in_tiles = 0 Then
        view_screen_tile_width = 20 'In tiles
        view_screen_tile_height = 15 'In tiles
    Else
        view_screen_tile_width = v_width_in_tiles 'In tiles
        view_screen_tile_height = v_height_in_tiles 'In tiles
    End If
    
    If windowed Then
        Dim target As RECT
        GetWindowRect screen_hwnd, target
        view_screen_top = target.top
        view_screen_left = target.left
        view_screen_right = target.Right
        view_screen_bottom = target.bottom
        view_screen_width = target.Right - target.left
        view_screen_height = target.bottom - target.top
        screen_width = view_screen_width
        screen_height = view_screen_height
        engine_windowed = True
    Else
        view_screen_left = v_left 'In pixels
        view_screen_top = v_top 'In pixels
        view_screen_width = view_screen_tile_width * base_tile_size
        view_screen_height = view_screen_tile_height * base_tile_size
        view_screen_right = view_screen_left + view_screen_width - 1
        view_screen_bottom = view_screen_top + view_screen_height - 1
        screen_width = s_width
        screen_height = s_height
        engine_windowed = False
    End If
    
    Dim d3dpp2 As D3DPRESENT_PARAMETERS
    
    With d3dpp2
        If engine_windowed Then
            .windowed = 1
            
            If engine_vsync Then
                .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
            Else
                .SwapEffect = D3DSWAPEFFECT_COPY
            End If
            
            .BackBufferFormat = d3ddm.format 'current display depth
        Else
            .SwapEffect = D3DSWAPEFFECT_FLIP
            'Set refresh rate
            .FullScreen_RefreshRateInHz = D3DPRESENT_RATE_DEFAULT
            'Turn off vsync
            .FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_IMMEDIATE
            'set color depth
            .BackBufferFormat = D3DFMT_R5G6B5 '16 bit
            'Back buffer size
            .BackBufferWidth = screen_width
            .BackBufferHeight = screen_height
        End If
        .hDeviceWindow = screen_hwnd
    End With
    
    d3dpp = d3dpp2
    
    Device_Reset
Exit Function

ErrHandler:
    MsgBox "Couldn't change resolution."

End Function

Private Sub Grh_Create_Mask(ByRef hdcsrc As Long, ByRef MaskDC As Long, ByVal src_x As Integer, ByVal src_y As Integer, ByVal src_width As Integer, ByVal src_height As Integer)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/30/2004
'Creates a Mask hDC, and sets the source hDC to work for trans bliting.
'**************************************************************
    Dim x As Integer
    Dim y As Integer
    Dim TransColor As Long
    Dim ColorKey As String
    
    'ColorKey = hex(COLOR_KEY)
    
    'Check if it has an alpha component
    'If Len(ColorKey) > 6 Then
         'get rid of alpha
    '    ColorKey = "&H" & Right$(ColorKey, 6)
    'End If
    'piluex prueba
    'TransColor = Val(ColorKey)
    ColorKey = "0"
    TransColor = &H0

    'Make it a mask (set background to black and foreground to white)
    'And set the sprite's background white
    For y = src_y To src_height + src_y
        For x = src_x To src_width + src_x
            If GetPixel(hdcsrc, x, y) = TransColor Then
                SetPixel MaskDC, x, y, vbWhite
                SetPixel hdcsrc, x, y, vbBlack
            Else
                SetPixel MaskDC, x, y, vbBlack
            End If
        Next x
    Next y
End Sub

Private Function Particle_Group_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until particle_group_list(loopc).active = False
        If loopc = particle_group_last Then
            Particle_Group_Next_Open = particle_group_last + 1
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Particle_Group_Next_Open = loopc
Exit Function
ErrorHandler:
    Particle_Group_Next_Open = 1
End Function

Private Function Particle_Group_Check(ByVal particle_group_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check index
    If particle_group_index > 0 And particle_group_index <= particle_group_last Then
        If particle_group_list(particle_group_index).active Then
            Particle_Group_Check = True
        End If
    End If
End Function

Public Function Particle_Group_Create(ByVal map_x As Integer, ByVal map_y As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                        Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                        Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                        Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                        Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                        Optional grh_resizex As Integer, Optional grh_resizey As Integer)
'**************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/14/2003
'Returns the particle_group_index if successful, else 0
'Modified by Juan Martín Sotuyo Dodero
'Modified by Augusto José Rando
'**************************************************************
    
    If (map_x <> -1) And (map_y <> -1) Then
        If Map_Particle_Group_Get(map_x, map_y) = 0 Then
            Particle_Group_Create = Particle_Group_Next_Open
            Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
        End If
    Else
        Particle_Group_Create = Particle_Group_Next_Open
        Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
    End If

End Function

Public Function Particle_Group_Remove(ByVal particle_group_index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Particle_Group_Check(particle_group_index) Then
        Particle_Group_Destroy particle_group_index
        Particle_Group_Remove = True
    End If
End Function

Public Function Particle_Group_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim Index As Long
    
    For Index = 1 To particle_group_last
        'Make sure it's a legal index
        If Particle_Group_Check(Index) Then
            Particle_Group_Destroy Index
        End If
    Next Index
    
    Particle_Group_Remove_All = True
End Function

Public Function Particle_Group_Find(ByVal id As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until particle_group_list(loopc).id = id
        If loopc = particle_group_last Then
            Particle_Group_Find = 0
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Particle_Group_Find = loopc
Exit Function
ErrorHandler:
    Particle_Group_Find = 0
End Function

Private Sub Particle_Group_Destroy(ByVal particle_group_index As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim temp As particle_group
    Dim i As Integer
    
    If particle_group_list(particle_group_index).map_x > 0 And particle_group_list(particle_group_index).map_y > 0 Then
        MapData(particle_group_list(particle_group_index).map_x, particle_group_list(particle_group_index).map_y).particle_group_index = 0
    ElseIf particle_group_list(particle_group_index).char_index Then
        If Char_Check(particle_group_list(particle_group_index).char_index) Then
            For i = 1 To CharList(particle_group_list(particle_group_index).char_index).particle_count
                If CharList(particle_group_list(particle_group_index).char_index).particle_group(i) = particle_group_index Then
                    CharList(particle_group_list(particle_group_index).char_index).particle_group(i) = 0
                    
                    'We don't resize arrays by now, it's really a waste...
                    'If i = UBound(CharList(particle_group_list(particle_group_index).char_index).particle_group) Then
                    '    CharList(particle_group_list(particle_group_index).char_index).particle_count = i - 1
                    '    ReDim Preserve CharList(particle_group_list(particle_group_index).char_index).particle_group(1 To (i - 1)) As Long
                    'End If
                    
                    Exit For
                End If
            Next i
        End If
    ElseIf particle_group_index = meteo_particle Then
        meteo_particle = 0
    End If
    
    particle_group_list(particle_group_index) = temp
    
    'Update array size
    If particle_group_index = particle_group_last Then
        Do Until particle_group_list(particle_group_last).active
            particle_group_last = particle_group_last - 1
            If particle_group_last = 0 Then
                particle_group_count = 0
                Exit Sub
            End If
        Loop
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count - 1
End Sub

Private Sub Particle_Group_Make(ByVal particle_group_index As Long, ByVal map_x As Integer, ByVal map_y As Integer, _
                                ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                Optional grh_resizex As Integer, Optional grh_resizey As Integer)
                                
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/15/2003
'Makes a new particle effect
'Modified by Juan Martín Sotuyo Dodero
'*****************************************************************
    'Update array size
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count + 1
    
    'Make active
    particle_group_list(particle_group_index).active = True
    
    'Map pos
    If (map_x <> -1) And (map_y <> -1) Then
        particle_group_list(particle_group_index).map_x = map_x
        particle_group_list(particle_group_index).map_y = map_y
    End If
    
    'Grh list
    ReDim particle_group_list(particle_group_index).grh_index_list(1 To UBound(grh_index_list))
    particle_group_list(particle_group_index).grh_index_list() = grh_index_list()
    particle_group_list(particle_group_index).grh_index_count = UBound(grh_index_list)
    
    'Sets alive vars
    If alive_counter = -1 Then
        particle_group_list(particle_group_index).alive_counter = -1
        particle_group_list(particle_group_index).never_die = True
    Else
        particle_group_list(particle_group_index).alive_counter = alive_counter
        particle_group_list(particle_group_index).never_die = False
    End If
    
    'alpha blending
    particle_group_list(particle_group_index).alpha_blend = alpha_blend
    
    'stream type
    particle_group_list(particle_group_index).stream_type = stream_type
    
    'speed
    particle_group_list(particle_group_index).frame_speed = frame_speed
    
    particle_group_list(particle_group_index).x1 = x1
    particle_group_list(particle_group_index).y1 = y1
    particle_group_list(particle_group_index).x2 = x2
    particle_group_list(particle_group_index).y2 = y2
    particle_group_list(particle_group_index).angle = angle
    particle_group_list(particle_group_index).vecx1 = vecx1
    particle_group_list(particle_group_index).vecx2 = vecx2
    particle_group_list(particle_group_index).vecy1 = vecy1
    particle_group_list(particle_group_index).vecy2 = vecy2
    particle_group_list(particle_group_index).life1 = life1
    particle_group_list(particle_group_index).life2 = life2
    particle_group_list(particle_group_index).fric = fric
    particle_group_list(particle_group_index).spin = spin
    particle_group_list(particle_group_index).spin_speedL = spin_speedL
    particle_group_list(particle_group_index).spin_speedH = spin_speedH
    particle_group_list(particle_group_index).gravity = gravity
    particle_group_list(particle_group_index).grav_strength = grav_strength
    particle_group_list(particle_group_index).bounce_strength = bounce_strength
    particle_group_list(particle_group_index).XMove = XMove
    particle_group_list(particle_group_index).YMove = YMove
    particle_group_list(particle_group_index).move_x1 = move_x1
    particle_group_list(particle_group_index).move_x2 = move_x2
    particle_group_list(particle_group_index).move_y1 = move_y1
    particle_group_list(particle_group_index).move_y2 = move_y2
    
    particle_group_list(particle_group_index).rgb_list(0) = rgb_list(0)
    particle_group_list(particle_group_index).rgb_list(1) = rgb_list(1)
    particle_group_list(particle_group_index).rgb_list(2) = rgb_list(2)
    particle_group_list(particle_group_index).rgb_list(3) = rgb_list(3)
    
    particle_group_list(particle_group_index).grh_resize = grh_resize
    particle_group_list(particle_group_index).grh_resizex = grh_resizex
    particle_group_list(particle_group_index).grh_resizey = grh_resizey
    
    'handle
    particle_group_list(particle_group_index).id = id
    
    'create particle stream
    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)
    
    'plot particle group on map
    If (map_x <> -1) And (map_y <> -1) Then
        MapData(map_x, map_y).particle_group_index = particle_group_index
    End If
    
End Sub

Private Sub Char_Particle_Group_Make(ByVal particle_group_index As Long, ByVal char_index As Integer, ByVal particle_char_index As Integer, _
                                ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                Optional grh_resizex As Integer, Optional grh_resizey As Integer)
                                
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/15/2003
'Makes a new particle effect
'Modified by Juan Martín Sotuyo Dodero
'*****************************************************************
    'Update array size
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count + 1
    
    'Make active
    particle_group_list(particle_group_index).active = True
    
    'Char index
    particle_group_list(particle_group_index).char_index = char_index
    
    'Grh list
    ReDim particle_group_list(particle_group_index).grh_index_list(1 To UBound(grh_index_list))
    particle_group_list(particle_group_index).grh_index_list() = grh_index_list()
    particle_group_list(particle_group_index).grh_index_count = UBound(grh_index_list)
    
    'Sets alive vars
    If alive_counter = -1 Then
        particle_group_list(particle_group_index).alive_counter = -1
        particle_group_list(particle_group_index).never_die = True
    Else
        particle_group_list(particle_group_index).alive_counter = alive_counter
        particle_group_list(particle_group_index).never_die = False
    End If
    
    'alpha blending
    particle_group_list(particle_group_index).alpha_blend = alpha_blend
    
    'stream type
    particle_group_list(particle_group_index).stream_type = stream_type
    
    'speed
    particle_group_list(particle_group_index).frame_speed = frame_speed
    
    particle_group_list(particle_group_index).x1 = x1
    particle_group_list(particle_group_index).y1 = y1
    particle_group_list(particle_group_index).x2 = x2
    particle_group_list(particle_group_index).y2 = y2
    particle_group_list(particle_group_index).angle = angle
    particle_group_list(particle_group_index).vecx1 = vecx1
    particle_group_list(particle_group_index).vecx2 = vecx2
    particle_group_list(particle_group_index).vecy1 = vecy1
    particle_group_list(particle_group_index).vecy2 = vecy2
    particle_group_list(particle_group_index).life1 = life1
    particle_group_list(particle_group_index).life2 = life2
    particle_group_list(particle_group_index).fric = fric
    particle_group_list(particle_group_index).spin = spin
    particle_group_list(particle_group_index).spin_speedL = spin_speedL
    particle_group_list(particle_group_index).spin_speedH = spin_speedH
    particle_group_list(particle_group_index).gravity = gravity
    particle_group_list(particle_group_index).grav_strength = grav_strength
    particle_group_list(particle_group_index).bounce_strength = bounce_strength
    particle_group_list(particle_group_index).XMove = XMove
    particle_group_list(particle_group_index).YMove = YMove
    particle_group_list(particle_group_index).move_x1 = move_x1
    particle_group_list(particle_group_index).move_x2 = move_x2
    particle_group_list(particle_group_index).move_y1 = move_y1
    particle_group_list(particle_group_index).move_y2 = move_y2
    
    particle_group_list(particle_group_index).rgb_list(0) = rgb_list(0)
    particle_group_list(particle_group_index).rgb_list(1) = rgb_list(1)
    particle_group_list(particle_group_index).rgb_list(2) = rgb_list(2)
    particle_group_list(particle_group_index).rgb_list(3) = rgb_list(3)
    
    particle_group_list(particle_group_index).grh_resize = grh_resize
    particle_group_list(particle_group_index).grh_resizex = grh_resizex
    particle_group_list(particle_group_index).grh_resizey = grh_resizey
    
    'handle
    particle_group_list(particle_group_index).id = id
    
    'create particle stream
    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)
    
    'plot particle group on char
    CharList(char_index).particle_group(particle_char_index) = particle_group_index
End Sub

Private Sub Particle_Group_Render(ByVal particle_group_index As Long, ByVal screen_x As Integer, ByVal screen_y As Integer)
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Modified by: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/15/2003
'Renders a particle stream at a paticular screen point
'*****************************************************************
    Dim loopc As Long
    Dim temp_rgb(0 To 3) As Long
    Dim no_move As Boolean
    
    'Set colors
    temp_rgb(0) = particle_group_list(particle_group_index).rgb_list(0)
    temp_rgb(1) = particle_group_list(particle_group_index).rgb_list(1)
    temp_rgb(2) = particle_group_list(particle_group_index).rgb_list(2)
    temp_rgb(3) = particle_group_list(particle_group_index).rgb_list(3)
    
    If particle_group_list(particle_group_index).alive_counter Then
    
        'See if it is time to move a particle
        particle_group_list(particle_group_index).frame_counter = particle_group_list(particle_group_index).frame_counter + timer_ticks_per_frame
        If particle_group_list(particle_group_index).frame_counter > particle_group_list(particle_group_index).frame_speed Then
            particle_group_list(particle_group_index).frame_counter = 0
            no_move = False
        Else
            no_move = True
        End If
    
    
        'If it's still alive render all the particles inside
        For loopc = 1 To particle_group_list(particle_group_index).particle_count
        
            'Render particle
            Particle_Render particle_group_list(particle_group_index).particle_stream(loopc), _
                            screen_x, screen_y, _
                            particle_group_list(particle_group_index).grh_index_list(Round(General_Random_Number(1, particle_group_list(particle_group_index).grh_index_count), 0)), _
                            temp_rgb(), _
                            particle_group_list(particle_group_index).alpha_blend, no_move, _
                            particle_group_list(particle_group_index).x1, particle_group_list(particle_group_index).y1, particle_group_list(particle_group_index).angle, _
                            particle_group_list(particle_group_index).vecx1, particle_group_list(particle_group_index).vecx2, _
                            particle_group_list(particle_group_index).vecy1, particle_group_list(particle_group_index).vecy2, _
                            particle_group_list(particle_group_index).life1, particle_group_list(particle_group_index).life2, _
                            particle_group_list(particle_group_index).fric, particle_group_list(particle_group_index).spin_speedL, _
                            particle_group_list(particle_group_index).gravity, particle_group_list(particle_group_index).grav_strength, _
                            particle_group_list(particle_group_index).bounce_strength, particle_group_list(particle_group_index).x2, _
                            particle_group_list(particle_group_index).y2, particle_group_list(particle_group_index).XMove, _
                            particle_group_list(particle_group_index).move_x1, particle_group_list(particle_group_index).move_x2, _
                            particle_group_list(particle_group_index).move_y1, particle_group_list(particle_group_index).move_y2, _
                            particle_group_list(particle_group_index).YMove, particle_group_list(particle_group_index).spin_speedH, _
                            particle_group_list(particle_group_index).spin, particle_group_list(particle_group_index).grh_resize, particle_group_list(particle_group_index).grh_resizex, particle_group_list(particle_group_index).grh_resizey
        Next loopc
        
        If no_move = False Then
            'Update the group alive counter
            If particle_group_list(particle_group_index).never_die = False Then
                particle_group_list(particle_group_index).alive_counter = particle_group_list(particle_group_index).alive_counter - 1
            End If
        End If
    
    Else
        'If it's dead destroy it
        Particle_Group_Destroy particle_group_index
    End If
End Sub

Private Sub Particle_Render(ByRef temp_particle As Particle, ByVal screen_x As Integer, ByVal screen_y As Integer, _
                            ByVal grh_index As Long, ByRef rgb_list() As Long, _
                            Optional ByVal alpha_blend As Boolean, Optional ByVal no_move As Boolean, _
                            Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                            Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                            Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                            Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                            Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                            Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                            Optional ByVal bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                            Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                            Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                            Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                            Optional grh_resizex As Integer, Optional grh_resizey As Integer)
'**************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Modified by: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/15/2003
'**************************************************************
    If no_move = False Then
        If temp_particle.alive_counter = 0 Then
            'Start new particle
            Grh_Initialize temp_particle.Grh, grh_index, alpha_blend
            temp_particle.x = General_Random_Number(x1, x2) - (base_tile_size / 2)
            temp_particle.y = General_Random_Number(y1, y2) - (base_tile_size / 2)
            temp_particle.vector_x = General_Random_Number(vecx1, vecx2)
            temp_particle.vector_y = General_Random_Number(vecy1, vecy2)
            temp_particle.angle = angle
            temp_particle.alive_counter = General_Random_Number(life1, life2)
            temp_particle.friction = fric
        Else
            'Continue old particle
            'Do gravity
            If gravity = True Then
                temp_particle.vector_y = temp_particle.vector_y + grav_strength
                If temp_particle.y > 0 Then
                    'bounce
                    temp_particle.vector_y = bounce_strength
                End If
            End If
            'Do rotation
            If spin = True Then temp_particle.Grh.angle = temp_particle.Grh.angle + (General_Random_Number(spin_speedL, spin_speedH) / 100)
            If temp_particle.angle >= 360 Then
                temp_particle.angle = 0
            End If
            
            If XMove = True Then temp_particle.vector_x = General_Random_Number(move_x1, move_x2)
            If YMove = True Then temp_particle.vector_y = General_Random_Number(move_y1, move_y2)
        End If
        
        'Add in vector
        temp_particle.x = temp_particle.x + (temp_particle.vector_x \ temp_particle.friction)
        temp_particle.y = temp_particle.y + (temp_particle.vector_y \ temp_particle.friction)
    
        'decrement counter
         temp_particle.alive_counter = temp_particle.alive_counter - 1
    End If
    
    'Draw it
    If grh_resize = True Then
        If temp_particle.Grh.GrhIndex Then
             Grh_Render_Advance temp_particle.Grh, temp_particle.x + screen_x, temp_particle.y + screen_y, grh_resizex, grh_resizey, rgb_list(), True, True, alpha_blend
            Exit Sub
        End If
    End If
    If temp_particle.Grh.GrhIndex Then
        Grh_Render temp_particle.Grh, temp_particle.x + screen_x, temp_particle.y + screen_y, rgb_list(), True, True, alpha_blend
    End If
End Sub

Public Function Input_Mouse_Moved_Get() As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Input_Mouse_Moved_Get = input_mouse_moved
    input_mouse_moved = False
End Function

Private Sub Input_Mouse_Poll()
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'**************************************************************
On Error Resume Next
    
    'Get the mouse event buffer
    Static DevData(1 To 1000) As DIDEVICEOBJECTDATA
    Dim nEvents As Long
    Dim i As Long
    
    nEvents = ddevice_mouse.GetDeviceData(DevData, DIGDD_DEFAULT)
    
    'Check buffer for clicks
    For i = 1 To nEvents
        Select Case DevData(i).lOfs
            Case DIMOFS_BUTTON0
                'Check if mouse button was pressed or released
                input_mouse_button_left = (DevData(i).lData And &H80)
                
            Case DIMOFS_BUTTON1
                'Check if mouse button was pressed or released
                input_mouse_button_right = (DevData(i).lData And &H80)
        End Select
    Next i
    
    'Use a API to get the mouse cordinates.
    Dim temp_point As PointAPI
    Dim temp_x As Integer
    Dim temp_y As Integer
    GetCursorPos temp_point
    
    'See if we moved
    If input_mouse_screen_x <> temp_point.x Or input_mouse_screen_y <> temp_point.y Then
        
        'We moved
        'Store new screen position
        input_mouse_screen_x = temp_point.x
        input_mouse_screen_y = temp_point.y
        
        'See if we are in the view area
        Convert_Screen_To_View input_mouse_screen_x, input_mouse_screen_y, input_mouse_view_x, input_mouse_view_y
        
        input_mouse_moved = True
    End If
    
    'Figure out where on the map the mouse is
    Convert_View_To_Map input_mouse_view_x, input_mouse_view_y, input_mouse_map_x, input_mouse_map_y
    
    'Figure out what direction from center of screen to where the mouse is
    Convert_Map_To_Direction input_mouse_map_x, input_mouse_map_y, temp_x, temp_y
    
    'Figure out heading from center of screen
    input_mouse_heading = Convert_Direction_To_Heading(temp_x, temp_y)
End Sub

Public Sub Input_Mouse_Screen_Get(ByRef screen_x As Integer, ByRef screen_y As Integer)
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo dodero
'Last Modify Date: 10/07/2002
'
'**************************************************************
    screen_x = input_mouse_screen_x
    screen_y = input_mouse_screen_y
End Sub

Public Sub Input_Mouse_View_Get(ByRef view_x As Integer, ByRef view_y As Integer)
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo dodero
'Last Modify Date: 10/07/2002
'
'**************************************************************
    view_x = input_mouse_view_x
    view_y = input_mouse_view_y
End Sub

Public Function Input_Mouse_In_View() As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Input_Mouse_In_View = (input_mouse_view_x >= 0 And input_mouse_view_x < view_screen_width And input_mouse_view_y >= 0 And input_mouse_view_y < view_screen_height)
End Function

Public Function Input_Mouse_Button_Left_Get() As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Input_Mouse_Button_Left_Get = input_mouse_button_left
End Function

Public Function Input_Mouse_Button_Right_Get() As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Input_Mouse_Button_Right_Get = input_mouse_button_right
End Function

Public Function Input_Mouse_Heading_Get() As Byte
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Input_Mouse_Heading_Get = input_mouse_heading
End Function

Public Sub Input_Mouse_Map_Get(ByRef map_x As Integer, ByRef map_y As Integer)
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'**************************************************************
    map_x = input_mouse_map_x
    map_y = input_mouse_map_y
End Sub

'Public Function Input_Key_Start()
'**************************************************************
'Author: Aaron Perkins - Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'**************************************************************
'    ddevice_keyboard.GetDeviceStateKeyboard key_state
'End Function

Public Function Input_Key_Get(ByVal key_code As Byte) As Boolean
'**************************************************************
'Author: Aaron Perkins - Juan Martín Sotuyo Dodero
'Modified by Augusto José Rando
'Now we use DirectInput Keyboard
'Last Modify Date: 10/07/2002
'
'**************************************************************
    'Input_Key_Get = (key_state.Key(key_code) > 0)
    Input_Key_Get = (GetKeyState(key_code) < 0)
End Function

Private Function Convert_Direction_To_Heading(ByVal direction_x As Integer, ByVal direction_y As Integer) As Byte
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'**************************************************************
    'North
    If direction_x = 0 And direction_y = -1 Then
        Convert_Direction_To_Heading = 1
    End If
    'NE
    If direction_x = 1 And direction_y = -1 Then
        Convert_Direction_To_Heading = 2
    End If
    'East
    If direction_x = 1 And direction_y = 0 Then
        Convert_Direction_To_Heading = 3
    End If
    'SE
    If direction_x = 1 And direction_y = 1 Then
        Convert_Direction_To_Heading = 4
    End If
    'South
    If direction_x = 0 And direction_y = 1 Then
        Convert_Direction_To_Heading = 5
    End If
    'SW
    If direction_x = -1 And direction_y = 1 Then
        Convert_Direction_To_Heading = 6
    End If
    'West
    If direction_x = -1 And direction_y = 0 Then
        Convert_Direction_To_Heading = 7
    End If
    'NW
    If direction_x = -1 And direction_y = -1 Then
        Convert_Direction_To_Heading = 8
    End If
End Function

Private Sub Convert_Heading_to_Direction(ByVal Heading As Byte, ByRef direction_x As Integer, ByRef direction_y As Integer)
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim addY As Integer
    Dim addX As Integer
    
    'Figure out which way to move
    Select Case Heading
        Case 1
            addY = -1
        
        Case 2
            addY = -1
            addX = 1
        
        Case 3
            addX = 1
        
        Case 4
            addX = 1
            addY = 1
        
        Case 5
            addY = 1
        
        Case 6
            addX = -1
            addY = 1
        
        Case 7
            addX = -1
        
        Case 8
            addX = -1
            addY = -1
    End Select
    
    direction_x = direction_x + addX
    direction_y = direction_y + addY
End Sub

Private Function Convert_Tile_To_View_X(ByVal x As Integer) As Integer
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'Convert tile position into position in view area
'**************************************************************
    If engine_windowed Then
        Convert_Tile_To_View_X = ((x - 1) * base_tile_size)
    Else
        Convert_Tile_To_View_X = view_screen_left + ((x - 1) * base_tile_size)
    End If
End Function

Private Function Convert_Tile_To_View_Y(ByVal y As Integer) As Integer
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'Convert tile position into position in view area
'**************************************************************
    If engine_windowed Then
        Convert_Tile_To_View_Y = ((y - 1) * base_tile_size)
    Else
        Convert_Tile_To_View_Y = view_screen_top + ((y - 1) * base_tile_size)
    End If
End Function

Private Sub Convert_Screen_To_View(ByVal screen_x As Integer, ByVal screen_y As Integer, ByRef view_x As Integer, ByRef view_y As Integer)
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'**************************************************************
    view_x = screen_x - view_screen_left
    view_y = screen_y - view_screen_top
End Sub

Private Sub Convert_View_To_Map(ByVal view_x As Integer, ByVal view_y As Integer, ByRef map_x As Integer, ByRef map_y As Integer)
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim half_view_tile_width As Integer
    Dim half_view_tile_height As Integer
    
    half_view_tile_width = (view_screen_tile_width \ 2)
    half_view_tile_height = (view_screen_tile_height \ 2)
    
    map_x = (view_x \ base_tile_size)
    map_y = (view_y \ base_tile_size)
    
    If map_x > half_view_tile_width Then
        map_x = (map_x - half_view_tile_width)
    Else
        If map_x < half_view_tile_width Then
            map_x = (0 - (half_view_tile_width - map_x))
        Else
            map_x = 0
        End If
    End If
    
    If map_y > half_view_tile_height Then
        map_y = (0 - (half_view_tile_height - map_y))
    Else
        If map_y < half_view_tile_height Then
            map_y = (map_y - half_view_tile_height)
        Else
            map_y = 0
        End If
    End If
    
    map_x = view_pos_x + map_x
    map_y = view_pos_y + map_y
End Sub

Private Sub Convert_Map_To_Direction(ByVal map_x As Integer, ByVal map_y As Integer, ByRef direction_x As Integer, ByRef direction_y As Integer)
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim temp_x As Integer
    Dim temp_y As Integer
    
    temp_x = map_x - view_pos_x
    temp_y = map_y - view_pos_y
    
    If temp_x <> 0 Then
        direction_x = temp_x \ Abs(temp_x)
    Else
        direction_x = 0
    End If
    If temp_y <> 0 Then
        direction_y = temp_y \ Abs(temp_y)
    Else
        direction_y = 0
    End If
End Sub

Private Sub Grh_Load_All()
'*****************************************************************
'Loads Grh.dat
'*****************************************************************

On Error GoTo ErrorHandler

Dim Grh As Integer
Dim frame As Integer
Dim TempInt As Integer

'Resize arrays / USAMOS CANT_GRH_INDEX POR DEFAULT
ReDim GrhData(1 To CANT_GRH_INDEX) As GrhData

If Not Extract_File(Scripts, App.Path & "\Recursos", "graficos.ind", Windows_Temp_Dir, False) Then
    Err.Description = "¡No se puede cargar el archivo de recurso!"
    GoTo ErrorHandler
End If

'Open files
Open Windows_Temp_Dir & "Graficos.ind" For Binary Access Read As #1

Seek #1, 1

Get #1, , TempInt
Get #1, , TempInt
Get #1, , TempInt
Get #1, , TempInt
Get #1, , TempInt

'Fill Grh List

'Get first Grh Number
Get #1, , Grh

Do Until Grh <= 0
    'Get number of frames
    Get #1, , GrhData(Grh).NumFrames
    
    If GrhData(Grh).NumFrames <= 0 Then
        GoTo ErrorHandler
    End If
    
    ReDim GrhData(Grh).Frames(1 To GrhData(Grh).NumFrames)
    
    If GrhData(Grh).NumFrames > 1 Then
    
        'Read a animation GRH set
        For frame = 1 To GrhData(Grh).NumFrames
        
            Get #1, , GrhData(Grh).Frames(frame)
            If GrhData(Grh).Frames(frame) <= 0 Or GrhData(Grh).Frames(frame) > CANT_GRH_INDEX Then
                GoTo ErrorHandler
            End If
        
        Next frame
    
        Get #1, , TempInt
        
        If TempInt <= 0 Then
            GoTo ErrorHandler
        End If
        
        GrhData(Grh).speed = CLng(TempInt)
        
        'Compute width and height
        GrhData(Grh).pixelHeight = GrhData(GrhData(Grh).Frames(1)).pixelHeight
        
        If GrhData(Grh).pixelHeight <= 0 Then
            GoTo ErrorHandler
        End If
        
        GrhData(Grh).pixelWidth = GrhData(GrhData(Grh).Frames(1)).pixelWidth
        
        If GrhData(Grh).pixelWidth <= 0 Then
            GoTo ErrorHandler
        End If
        
        GrhData(Grh).TileWidth = GrhData(GrhData(Grh).Frames(1)).TileWidth
        
        If GrhData(Grh).TileWidth <= 0 Then
            GoTo ErrorHandler
        End If
        
        GrhData(Grh).TileHeight = GrhData(GrhData(Grh).Frames(1)).TileHeight
        
        If GrhData(Grh).TileHeight <= 0 Then
            GoTo ErrorHandler
        End If
    
    Else
    
        'Read in normal GRH data
        Get #1, , GrhData(Grh).FileNum
        
        If GrhData(Grh).FileNum <= 0 Then
            GoTo ErrorHandler
        End If
        
        Get #1, , GrhData(Grh).sX
        
        If GrhData(Grh).sX < 0 Then
            GoTo ErrorHandler
        End If
        
        Get #1, , GrhData(Grh).sY
        
        If GrhData(Grh).sY < 0 Then
            GoTo ErrorHandler
        End If
            
        Get #1, , GrhData(Grh).pixelWidth
        
        If GrhData(Grh).pixelWidth <= 0 Then
            GoTo ErrorHandler
        End If
        
        Get #1, , GrhData(Grh).pixelHeight
        
        If GrhData(Grh).pixelHeight <= 0 Then
            GoTo ErrorHandler
        End If
        
        'Compute width and height
        'sinuhe_d3d
        GrhData(Grh).TileWidth = GrhData(Grh).pixelWidth / 32
        GrhData(Grh).TileHeight = GrhData(Grh).pixelHeight / 32
        
        GrhData(Grh).Frames(1) = Grh
            
    End If

    'Get Next Grh Number
    Get #1, , Grh

Loop
'************************************************

Close #1

    Dim count As Long
    Open App.Path & "\Init\minimap.dat" For Binary As #1
        Seek #1, 1
        For count = 1 To CANT_GRH_INDEX
            If Grh_Check(count) Then
                Get #1, , GrhData(count).mini_map_color
            End If
        Next count
    Close #1

    Delete_File Windows_Temp_Dir & "Graficos.ind"

Exit Sub

ErrorHandler:
    Close #1
    If General_File_Exists(Windows_Temp_Dir & "Graficos.ind", vbNormal) Then Delete_File Windows_Temp_Dir & "Graficos.ind"
    MsgBox "Error al cargar el recurso de índice de gráficos: " & Err.Description & " (" & Grh & ")", vbCritical, "Error al cargar"
    Call EndGame

End Sub

Sub Grh_Load_Weapons()

On Error GoTo ErrorHandler

Dim loopc As Integer
Dim Arch As String

If Not Extract_File(Scripts, App.Path & "\Recursos", "armas.dat", Windows_Temp_Dir, False) Then
    Err.Description = "¡No se puede cargar el archivo de recurso!"
    GoTo ErrorHandler
End If

Arch = Windows_Temp_Dir & "armas.dat"
NumWeaponAnims = Val(General_Var_Get(Arch, "INIT", "NumArmas"))

ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData

For loopc = 1 To NumWeaponAnims
    Grh_Initialize WeaponAnimData(loopc).WeaponWalk(NORTH), Val(General_Var_Get(Arch, "ARMA" & loopc, "Dir1"))
    Grh_Initialize WeaponAnimData(loopc).WeaponWalk(EAST), Val(General_Var_Get(Arch, "ARMA" & loopc, "Dir2"))
    Grh_Initialize WeaponAnimData(loopc).WeaponWalk(SOUTH), Val(General_Var_Get(Arch, "ARMA" & loopc, "Dir3"))
    Grh_Initialize WeaponAnimData(loopc).WeaponWalk(WEST), Val(General_Var_Get(Arch, "ARMA" & loopc, "Dir4"))
Next loopc

Delete_File Windows_Temp_Dir & "armas.dat"

Exit Sub

ErrorHandler:
    Close #1
    If General_File_Exists(Windows_Temp_Dir & "armas.dat", vbNormal) Then Delete_File Windows_Temp_Dir & "armas.dat"
    MsgBox "Error al cargar el recurso de índice de armas: " & Err.Description & " (" & Err.number & ")", vbCritical, "Error al cargar"
    Call EndGame

End Sub

Sub Grh_Load_Shields()

On Error GoTo ErrorHandler

Dim loopc As Integer
Dim Arch As String
Dim NumEscudosAnims As Integer

If Not Extract_File(Scripts, App.Path & "\Recursos", "escudos.dat", Windows_Temp_Dir, False) Then
    Err.Description = "¡No se puede cargar el archivo de recurso!"
    GoTo ErrorHandler
End If

Arch = Windows_Temp_Dir & "escudos.dat"

NumEscudosAnims = Val(General_Var_Get(Arch, "INIT", "NumEscudos"))

ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData

For loopc = 1 To NumEscudosAnims
    Grh_Initialize ShieldAnimData(loopc).ShieldWalk(NORTH), Val(General_Var_Get(Arch, "ESC" & loopc, "Dir1"))
    Grh_Initialize ShieldAnimData(loopc).ShieldWalk(EAST), Val(General_Var_Get(Arch, "ESC" & loopc, "Dir2"))
    Grh_Initialize ShieldAnimData(loopc).ShieldWalk(SOUTH), Val(General_Var_Get(Arch, "ESC" & loopc, "Dir3"))
    Grh_Initialize ShieldAnimData(loopc).ShieldWalk(WEST), Val(General_Var_Get(Arch, "ESC" & loopc, "Dir4"))
Next loopc

Delete_File Windows_Temp_Dir & "escudos.dat"

Exit Sub

ErrorHandler:
    Close #1
    If General_File_Exists(Windows_Temp_Dir & "escudos.dat", vbNormal) Then Delete_File Windows_Temp_Dir & "escudos.dat"
    MsgBox "Error al cargar el recurso de índice de escudos: " & Err.Description & " (" & Err.number & ")", vbCritical, "Error al cargar"
    Call EndGame

End Sub

Private Sub Grh_Load_Heads()

On Error GoTo ErrorHandler

Dim n As Integer, i As Integer, Numheads As Integer, Index As Integer

Dim Miscabezas() As tIndiceCabeza

n = FreeFile

If Not Extract_File(Scripts, App.Path & "\Recursos", "cabezas.ind", Windows_Temp_Dir, False) Then
    Err.Description = "No se ha logrado extraer el archivo de recurso."
    GoTo ErrorHandler
End If

Open Windows_Temp_Dir & "Cabezas.ind" For Binary Access Read As #n

'cabecera
Get #n, , MiCabecera

'num de cabezas
Get #n, , Numheads

'Resize array
ReDim HeadData(0 To Numheads + 1) As HeadData
ReDim Miscabezas(0 To Numheads + 1) As tIndiceCabeza

For i = 1 To Numheads
    Get #n, , Miscabezas(i)
    Grh_Initialize HeadData(i).head(NORTH), Miscabezas(i).head(1)
    Grh_Initialize HeadData(i).head(EAST), Miscabezas(i).head(2)
    Grh_Initialize HeadData(i).head(SOUTH), Miscabezas(i).head(3)
    Grh_Initialize HeadData(i).head(WEST), Miscabezas(i).head(4)
Next i

Close #n

Delete_File Windows_Temp_Dir & "Cabezas.ind"

Exit Sub

ErrorHandler:
    If General_File_Exists(General_Get_Temp_Dir & "Cabezas.ind", vbNormal) Then Delete_File General_Get_Temp_Dir & "Cabezas.ind"
    MsgBox "Error al cargar el recurso de índice de cabezas: " & Err.Description & " (" & Err.number & ")", vbCritical, "Error al cargar"
    Call EndGame

End Sub

Private Sub Grh_Load_Helmets()

On Error GoTo ErrorHandler

Dim n As Integer, i As Integer, NumCascos As Integer, Index As Integer

Dim Miscabezas() As tIndiceCabeza

n = FreeFile

If Not Extract_File(Scripts, App.Path & "\Recursos", "cascos.ind", Windows_Temp_Dir, False) Then
    Err.Description = "No se ha logrado extraer el archivo de recurso."
    GoTo ErrorHandler
End If

Open Windows_Temp_Dir & "Cascos.ind" For Binary Access Read As #n

'cabecera
Get #n, , MiCabecera

'num de cabezas
Get #n, , NumCascos

'Resize array
ReDim CascoAnimData(0 To NumCascos + 1) As HeadData
ReDim Miscabezas(0 To NumCascos + 1) As tIndiceCabeza

For i = 1 To NumCascos
    Get #n, , Miscabezas(i)
    Grh_Initialize CascoAnimData(i).head(NORTH), Miscabezas(i).head(1)
    Grh_Initialize CascoAnimData(i).head(EAST), Miscabezas(i).head(2)
    Grh_Initialize CascoAnimData(i).head(SOUTH), Miscabezas(i).head(3)
    Grh_Initialize CascoAnimData(i).head(WEST), Miscabezas(i).head(4)
Next i

Close #n

Delete_File Windows_Temp_Dir & "Cascos.ind"

Exit Sub

ErrorHandler:
    If General_File_Exists(Windows_Temp_Dir & "Cascos.ind", vbNormal) Then Delete_File Windows_Temp_Dir & "Cascos.ind"
    MsgBox "Error al cargar el recurso de índice de cascos: " & Err.Description & " (" & Err.number & ")", vbCritical, "Error al cargar"
    Call EndGame

End Sub

Private Sub Grh_Load_Body()

On Error GoTo ErrorHandler

Dim n As Integer, i As Integer
Dim NumCuerpos As Integer
Dim MisCuerpos() As tIndiceCuerpo

n = FreeFile

If Not Extract_File(Scripts, App.Path & "\Recursos", "personajes.ind", Windows_Temp_Dir, False) Then
    Err.Description = "No se ha logrado extraer el archivo de recurso."
    GoTo ErrorHandler
End If

Open Windows_Temp_Dir & "Personajes.ind" For Binary Access Read As #n

'cabecera
Get #n, , MiCabecera

'num de cabezas
Get #n, , NumCuerpos

'Resize array
ReDim BodyData(0 To NumCuerpos + 1) As BodyData
ReDim MisCuerpos(0 To NumCuerpos + 1) As tIndiceCuerpo

For i = 1 To NumCuerpos
    Get #n, , MisCuerpos(i)
    Grh_Initialize BodyData(i).Walk(NORTH), MisCuerpos(i).Body(1)
    Grh_Initialize BodyData(i).Walk(EAST), MisCuerpos(i).Body(2)
    Grh_Initialize BodyData(i).Walk(SOUTH), MisCuerpos(i).Body(3)
    Grh_Initialize BodyData(i).Walk(WEST), MisCuerpos(i).Body(4)
    BodyData(i).HeadOffset.x = MisCuerpos(i).HeadOffsetX
    BodyData(i).HeadOffset.y = MisCuerpos(i).HeadOffsetY
Next i

Close #n

Delete_File Windows_Temp_Dir & "Personajes.ind"

Exit Sub

ErrorHandler:
    If General_File_Exists(Windows_Temp_Dir & "Personajes.ind", vbNormal) Then Delete_File Windows_Temp_Dir & "Personajes.ind"
    MsgBox "Error al cargar el recurso de índice de cuerpos: " & Err.Description & " (" & Err.number & ")", vbCritical, "Error al cargar"
    Call EndGame

End Sub

Private Sub Grh_Load_Fx()

On Error GoTo ErrorHandler

Dim n As Integer, i As Integer
Dim NumFxs As Integer
Dim MisFxs() As tIndiceFx

n = FreeFile

If Not Extract_File(Scripts, App.Path & "\Recursos", "fxs.ind", Windows_Temp_Dir, False) Then
    Err.Description = "No se ha logrado extraer el archivo de recurso."
    GoTo ErrorHandler
End If

Open Windows_Temp_Dir & "Fxs.ind" For Binary Access Read As #n

'cabecera
Get #n, , MiCabecera

'num de cabezas
Get #n, , NumFxs

'Resize array
ReDim FxData(0 To NumFxs + 1) As FxData
ReDim MisFxs(0 To NumFxs + 1) As tIndiceFx

For i = 1 To NumFxs
    Get #n, , MisFxs(i)
    Call Grh_Initialize(FxData(i).fx, MisFxs(i).Animacion, 1)
    FxData(i).offsetX = MisFxs(i).offsetX
    FxData(i).OffsetY = MisFxs(i).OffsetY
Next i

Close #n

Delete_File Windows_Temp_Dir & "Fxs.ind"

Exit Sub

ErrorHandler:
    If General_File_Exists(Windows_Temp_Dir & "Fxs.ind", vbNormal) Then Delete_File Windows_Temp_Dir & "Fxs.ind"

End Sub

Private Sub Grh_Initialize(ByRef Grh As Grh, ByVal grh_index As Long, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'Sets up a grh. MUST be done before rendering
'*****************************************************************
    On Error Resume Next
    
    'Copy of parameters
    If grh_index = 0 Then Exit Sub
    
    Grh.GrhIndex = grh_index
    Grh.alpha_blend = alpha_blend
    Grh.angle = angle
    
    'Start it if it's a animated grh
    If GrhData(Grh.GrhIndex).NumFrames > 1 Then
        Grh.speed = 0.4
        Grh.started = 1
    Else
        Grh.started = 0
    End If
    
    'Set frame counters
    Grh.framecounter = 1
End Sub

Private Sub Grh_Uninitialize(ByRef Grh As Grh)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Resets a Grh
'*****************************************************************
    'Copy of parameters
    Grh.GrhIndex = 0
    Grh.alpha_blend = False
    Grh.angle = 0
    Grh.started = False
    'Set frame counters
    Grh.framecounter = 0
    Grh.speed = 0
End Sub

Private Function Device_Test_Cooperative_Level() As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/30/2004
'Handle Alt-Tab and Ctrl-Alt-Del
'**************************************************************
    'Call TestCooperativeLevel to see what state the device is in.
    Dim hr As Long
    hr = ddevice.TestCooperativeLevel
    If hr = D3DERR_DEVICELOST Then
        'Make sure all fonts react to the lsot device
        For hr = 1 To font_count
            font_list(hr).OnLostDevice
        Next hr
        'If the device is lost, exit and wait for it to come back.
        Exit Function
    ElseIf hr = D3DERR_DEVICENOTRESET Then
        'The device became lost for some reason (Alt-Tab or Ctrl-Alt-Del) and now
        'Reset() needs to be called to try and get the device back.
        Device_Reset
        'Return false so that the rendering loop will start over
        Exit Function
    End If
    
    Device_Test_Cooperative_Level = True
End Function

Private Function Grh_Check(ByVal grh_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check grh_index
    If grh_index > 0 And grh_index <= CANT_GRH_INDEX Then
        Grh_Check = GrhData(grh_index).NumFrames
    End If
End Function

Public Sub Load_Item_Data()

Dim NumItems As Long
Dim FileRoute As String
Dim Reader As clsLeerInis

Dim i As Long

FileRoute = App.Path & "\init\" & "obj.dat"

Set Reader = New clsLeerInis

Reader.Abrir FileRoute

NumItems = Val(Reader.DarValor("INIT", "NumOBJs"))

If NumItems > 0 Then
    ReDim Item_Data(1 To NumItems)
    For i = 1 To NumItems
        Item_Data(i).GrhIndex = Val(Reader.DarValor("OBJ" & i, "GrhIndex"))
        Item_Data(i).Type = Val(Reader.DarValor("OBJ" & i, "ObjType"))
        Item_Data(i).Name = Reader.DarValor("OBJ" & i, "Name")
    Next i
End If

Set Reader = Nothing

End Sub

Public Function NPC_Name_Get(ByVal npc_data_index As Integer) As String

'todo_barrin: validar

NPC_Name_Get = NPC_Data(npc_data_index).Name

End Function

Public Function NPC_Hostile_Name_Get(ByVal npc_data_index As Integer) As String

'todo_barrin: validar

NPC_Hostile_Name_Get = NPC_Hostile_Data(npc_data_index).Name

End Function

Public Sub Load_NPC_Data()

Dim NumNPCs As Long
Dim FileRoute As String
Dim Reader As clsLeerInis

Dim i As Long

FileRoute = App.Path & "\init\" & "npcs.dat"

Set Reader = New clsLeerInis

Reader.Abrir FileRoute

NumNPCs = Val(Reader.DarValor("INIT", "NumNPCs"))

If NumNPCs > 0 Then
    ReDim NPC_Data(1 To NumNPCs)
    For i = 1 To NumNPCs
        NPC_Data(i).Name = Reader.DarValor("NPC" & i, "Name")
        NPC_Data(i).Heading = Val(Reader.DarValor("NPC" & i, "Heading"))
        NPC_Data(i).BodyIndex = Val(Reader.DarValor("NPC" & i, "Body"))
        NPC_Data(i).HeadIndex = Val(Reader.DarValor("NPC" & i, "Head"))
    Next i
End If

FileRoute = App.Path & "\init\" & "npcs-hostiles.dat"

Reader.Abrir FileRoute

NumNPCs = Val(Reader.DarValor("INIT", "NumNPCs"))

If NumNPCs > 0 Then
    ReDim NPC_Hostile_Data(1 To (500 + NumNPCs))
    For i = 0 To NumNPCs
        NPC_Hostile_Data((500 + i)).Name = Reader.DarValor("NPC" & (500 + i), "Name")
        NPC_Hostile_Data((500 + i)).Heading = Val(Reader.DarValor("NPC" & (500 + i), "Heading"))
        NPC_Hostile_Data((500 + i)).BodyIndex = Val(Reader.DarValor("NPC" & (500 + i), "Body"))
        NPC_Hostile_Data((500 + i)).HeadIndex = Val(Reader.DarValor("NPC" & (500 + i), "Head"))
    Next i
End If

Set Reader = Nothing

End Sub

Public Function Item_Name_Get(s_item_data_index As Integer) As String

'todo_barrin: verificar

Item_Name_Get = Item_Data(s_item_data_index).Name

End Function

Public Function Item_Grh_Get(s_item_data_index As Integer) As Integer

'todo_barrin: verificar

Item_Grh_Get = Item_Data(s_item_data_index).GrhIndex

End Function

Public Sub Inventory_Render_Start()
'*******************************
'Erase the backbuffer so that it can be drawn on again
Device_Clear
'*******************************
    
'*******************************
'Start the scene
ddevice.BeginScene
'*******************************
End Sub

Public Sub Inventory_Render_End(ByVal desthWnd As Long)

Static re As RECT
re.left = 0
re.top = 0
re.bottom = 160
re.Right = 160

ddevice.EndScene
ddevice.Present re, ByVal 0&, desthWnd, ByVal 0&

Inventory_Render = False

End Sub

Function Engine_Text_Width(texto As String, Optional multi As Boolean = False) As Integer

Dim a, b, d, e, f As Integer
Dim graf As Grh

If multi = False Then
    For a = 1 To Len(texto)
        b = Asc(mid(texto, a, 1))
        graf.GrhIndex = Fuentes(1).Caracteres(b)
        If b <> 32 Then
            Engine_Text_Width = Engine_Text_Width + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1
        Else
            Engine_Text_Width = Engine_Text_Width + 4
        End If
    Next a
Else
    e = 0
    f = 0
    For a = 1 To Len(texto)
        b = Asc(mid(texto, a, 1))
        graf.GrhIndex = Fuentes(1).Caracteres(b)
        If b = 32 Or b = 13 Then
            If e >= 20 Then 'reemplazar por lo que os plazca
                f = f + 1
                e = 0
                d = 0
            Else
                If b = 32 Then d = d + 4
            End If
        Else
            If graf.GrhIndex > 12 Then
                d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1
                If d > Engine_Text_Width Then Engine_Text_Width = d
            End If
        End If
        e = e + 1
    Next a
End If

End Function

Public Sub Char_Dialog_Set(ByVal char_index As Integer, ByVal char_dialog As String, ByVal char_dialog_color As Long, ByVal char_dialog_life As Byte, Optional ByVal font_index As Integer = 1)
    
If Char_Check(char_index) Then
    CharList(char_index).dialog = char_dialog
    CharList(char_index).dialog_color = char_dialog_color
    CharList(char_index).dialog_life = char_dialog_life
    CharList(char_index).dialog_font_index = font_index
End If

End Sub

Public Sub Char_Dialog_Remove(ByVal char_index As Integer)
    
If Char_Check(char_index) Then
    CharList(char_index).dialog = ""
    CharList(char_index).dialog_color = 0
    CharList(char_index).dialog_life = 0
End If

End Sub

Function Engine_Text_Height(texto As String, Optional multi As Boolean = False) As Integer

Dim a, b, C, d, e, f As Integer
Dim graf As Grh
  
If multi = False Then
    Engine_Text_Height = 0
Else
    e = 0
    f = 0
    
    For a = 1 To Len(texto)
        b = Asc(mid(texto, a, 1))
        graf.GrhIndex = Fuentes(1).Caracteres(b)
        
        If b = 32 Or b = 13 Then
            If e >= 20 Then 'reemplazar por lo que os plazca
                f = f + 1
                e = 0
                d = 0
            Else
                If b = 32 Then
                    d = d + 4
                End If
            End If
        'Else
            'If graf.GrhIndex > 12 Then
        End If
        e = e + 1
    Next a
  
Engine_Text_Height = f * 14
  
End If

End Function

Sub Engine_Text_Render(texto As String, x As Integer, y As Integer, ByRef text_color() As Long, Optional ByVal font_index As Integer = 1, Optional multi_line As Boolean = False)

Dim a, b, C, d, e, f As Integer
Dim graf As Grh

Dim temp_array(3) As Long 'Si le queres dar color a la letra pasa este parametro dsp xD
temp_array(0) = text_color(0)
temp_array(1) = text_color(1)
temp_array(2) = text_color(2)
temp_array(3) = text_color(3)

If (Len(texto) = 0) Then Exit Sub

d = 0
If multi_line = False Then
    For a = 1 To Len(texto)
        b = Asc(mid(texto, a, 1))
        graf.GrhIndex = Fuentes(font_index).Caracteres(b)
        If b <> 32 Then
            If graf.GrhIndex <> 0 Then
                'mega sombra O-matica
                graf.GrhIndex = Fuentes(font_index).Caracteres(b) + 100
                Grh_Render graf, (x + d) + 1, y + 1, temp_array, False, False, False
                graf.GrhIndex = Fuentes(font_index).Caracteres(b)
                Grh_Render graf, (x + d), y, temp_array, False, False, False
                d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1
            End If
        Else
            d = d + 4
        End If
    Next a
Else
    e = 0
    f = 0
    For a = 1 To Len(texto)
        b = Asc(mid(texto, a, 1))
        graf.GrhIndex = Fuentes(font_index).Caracteres(b)
        If b = 32 Or b = 13 Then
            If e >= 20 Then 'reemplazar por lo que os plazca
                f = f + 1
                e = 0
                d = 0
            Else
                If b = 32 Then d = d + 4
            End If
        Else
            If graf.GrhIndex > 12 Then
                'mega sombra O-matica
                graf.GrhIndex = Fuentes(font_index).Caracteres(b) + 100
                Grh_Render graf, (x + d) + 1, y + 1 + f * 14, temp_array, False, False, False
                graf.GrhIndex = Fuentes(font_index).Caracteres(b)
                Grh_Render graf, (x + d), y + f * 14, temp_array, False, False, False '14 es el height de esta fuente dsp lo hacemos dinamico
                d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1
            End If
       End If
       e = e + 1
    Next a
End If

End Sub

Private Sub Engine_Font_Initialize()

Dim a As Integer

Fuentes(1).Tamanio = 9
Fuentes(1).Caracteres(48) = 21452
Fuentes(1).Caracteres(49) = 21453
Fuentes(1).Caracteres(50) = 21454
Fuentes(1).Caracteres(51) = 21455
Fuentes(1).Caracteres(52) = 21456
Fuentes(1).Caracteres(53) = 21457
Fuentes(1).Caracteres(54) = 21458
Fuentes(1).Caracteres(55) = 21459
Fuentes(1).Caracteres(56) = 21460
Fuentes(1).Caracteres(57) = 21461
For a = 0 To 25
Fuentes(1).Caracteres(a + 97) = 21400 + a
Next a

For a = 0 To 25
Fuentes(1).Caracteres(a + 65) = 21426 + a
Next a
Fuentes(1).Caracteres(33) = 21462
Fuentes(1).Caracteres(161) = 21463
Fuentes(1).Caracteres(34) = 21464
Fuentes(1).Caracteres(36) = 21465
Fuentes(1).Caracteres(191) = 21466
Fuentes(1).Caracteres(35) = 21467
Fuentes(1).Caracteres(36) = 21468
Fuentes(1).Caracteres(37) = 21469
Fuentes(1).Caracteres(38) = 21470
Fuentes(1).Caracteres(47) = 21471
Fuentes(1).Caracteres(92) = 21472
Fuentes(1).Caracteres(40) = 21473
Fuentes(1).Caracteres(41) = 21474
Fuentes(1).Caracteres(61) = 21475
Fuentes(1).Caracteres(39) = 21476
Fuentes(1).Caracteres(123) = 21477
Fuentes(1).Caracteres(125) = 21478
Fuentes(1).Caracteres(95) = 21479
Fuentes(1).Caracteres(45) = 21480
Fuentes(1).Caracteres(63) = 21465
Fuentes(1).Caracteres(64) = 21481
Fuentes(1).Caracteres(94) = 21482
Fuentes(1).Caracteres(91) = 21483
Fuentes(1).Caracteres(93) = 21484
Fuentes(1).Caracteres(60) = 21485
Fuentes(1).Caracteres(62) = 21486
Fuentes(1).Caracteres(42) = 21487
Fuentes(1).Caracteres(43) = 21488
Fuentes(1).Caracteres(46) = 21489
Fuentes(1).Caracteres(44) = 21490
Fuentes(1).Caracteres(58) = 21491
Fuentes(1).Caracteres(59) = 21492
Fuentes(1).Caracteres(124) = 21493
Fuentes(1).Caracteres(252) = 21800
Fuentes(1).Caracteres(220) = 21801
Fuentes(1).Caracteres(225) = 21802
Fuentes(1).Caracteres(233) = 21803
Fuentes(1).Caracteres(237) = 21804
Fuentes(1).Caracteres(243) = 21805
Fuentes(1).Caracteres(250) = 21806
Fuentes(1).Caracteres(253) = 21807
Fuentes(1).Caracteres(193) = 21808
Fuentes(1).Caracteres(201) = 21809
Fuentes(1).Caracteres(205) = 21810
Fuentes(1).Caracteres(211) = 21811
Fuentes(1).Caracteres(218) = 21812
Fuentes(1).Caracteres(221) = 21813
Fuentes(1).Caracteres(224) = 21814
Fuentes(1).Caracteres(232) = 21815
Fuentes(1).Caracteres(236) = 21816
Fuentes(1).Caracteres(242) = 21817
Fuentes(1).Caracteres(249) = 21818
Fuentes(1).Caracteres(192) = 21819
Fuentes(1).Caracteres(200) = 21820
Fuentes(1).Caracteres(204) = 21821
Fuentes(1).Caracteres(210) = 21822
Fuentes(1).Caracteres(217) = 21823
Fuentes(1).Caracteres(241) = 21824
Fuentes(1).Caracteres(209) = 21825
Fuentes(1).Caracteres(196) = 25238
Fuentes(1).Caracteres(194) = 25239
Fuentes(1).Caracteres(203) = 25240
Fuentes(1).Caracteres(207) = 25241
Fuentes(1).Caracteres(214) = 25242
Fuentes(1).Caracteres(212) = 25243

Fuentes(2).Tamanio = 9
Fuentes(2).Caracteres(97) = 21936
Fuentes(2).Caracteres(108) = 21937
Fuentes(2).Caracteres(115) = 21938
Fuentes(2).Caracteres(70) = 21939
Fuentes(2).Caracteres(48) = 21940
Fuentes(2).Caracteres(49) = 21941
Fuentes(2).Caracteres(50) = 21942
Fuentes(2).Caracteres(51) = 21943
Fuentes(2).Caracteres(52) = 21944
Fuentes(2).Caracteres(53) = 21945
Fuentes(2).Caracteres(54) = 21946
Fuentes(2).Caracteres(55) = 21947
Fuentes(2).Caracteres(56) = 21948
Fuentes(2).Caracteres(57) = 21949
Fuentes(2).Caracteres(33) = 21950
Fuentes(2).Caracteres(161) = 21951
Fuentes(2).Caracteres(42) = 21952

End Sub

Public Sub Grh_Inventory_Render(ByRef grh_index As Long, ByVal screen_x As Integer, ByVal screen_y As Integer, ByRef rgb_list() As Long, Optional ByVal h_centered As Boolean, Optional ByVal v_centered As Boolean, Optional ByVal alpha_blend As Boolean = False)
Dim inv_grh As Grh
Call Grh_Initialize(inv_grh, grh_index)
Call Grh_Render(inv_grh, screen_x, screen_y, rgb_list(), h_centered, v_centered, alpha_blend)
End Sub

Private Function VerifyDevice(ByVal usageflags As Long, ByVal format As CONST_D3DFORMAT) As Boolean
    
On Error GoTo Error_Handler
    
    'see if we can do blending
    If (d3dcaps.SrcBlendCaps And D3DPBLENDCAPS_ONE) = D3DPBLENDCAPS_ONE And _
            (d3dcaps.DestBlendCaps And D3DPBLENDCAPS_INVSRCALPHA) = D3DPBLENDCAPS_INVSRCALPHA Then
        VerifyDevice = True
    Else
        VerifyDevice = False
    End If
    
Exit Function

Error_Handler:
    VerifyDevice = False

End Function

Public Function Map_Load_From_File(MapRoute As String, Optional ByVal Mensajes As Boolean = True, Optional ByVal Cliente As Boolean = False) As Integer

Dim C As String, inf As MapInfo
ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As tMapBlock

Particle_Group_Remove_All
Light_Remove_All
Char_Remove_All
MapInfo = inf

If UCase(Right$(MapRoute, 4)) <> ".MAP" Then Exit Function

If Not Load_Map_Data(MapRoute, Mensajes) Then Exit Function

If Not Cliente Then
    If Not Load_Dat_Data(MapRoute, Mensajes) Then Exit Function
    If Not Load_Inf_Data(MapRoute, Mensajes) Then Exit Function
End If

C = mid$(MapRoute, 1, Len(MapRoute) - 4)
Map_Load_From_File = Val(mid(C, InStrRev(LCase(C), "mapa") + Len("mapa")))

End Function

Private Function Load_Map_Data(MapRoute As String, Optional Mensajes As Boolean = True) As Boolean

On Error Resume Next

Dim loopc As Integer
Dim y As Integer
Dim x As Integer
Dim TempInt As Integer
Dim TempLong As Long
Dim Idx As Integer
Dim LastGRH As Integer

Open MapRoute For Binary As #1
Seek #1, 1
        
Dim buffer(1 To ((YMaxMapSize - YMinMapSize + 1) * (XMaxMapSize - XMinMapSize + 1))) As tTileMap
        
'map Header
Get #1, , TempInt
Get #1, , MiCabecera
Get #1, , TempInt
Get #1, , TempInt
Get #1, , TempInt
Get #1, , TempInt

'main array
Get #1, , buffer

'offset
Idx = 1

'Load arrays
For y = YMinMapSize To YMaxMapSize
    For x = XMinMapSize To XMaxMapSize

        MapData(x, y).Blocked = buffer(Idx).bloqueado
        
        For loopc = 1 To 4
            
            If loopc = 1 Then
                MapData(x, y).Graphic(loopc).GrhIndex = buffer(Idx).graf1
            ElseIf loopc = 2 Then
                MapData(x, y).Graphic(loopc).GrhIndex = buffer(Idx).graf2
            ElseIf loopc = 3 Then
                MapData(x, y).Graphic(loopc).GrhIndex = buffer(Idx).graf3
            ElseIf loopc = 4 Then
                MapData(x, y).Graphic(loopc).GrhIndex = buffer(Idx).graf4
            End If
            
            'Set up GRH
            If MapData(x, y).Graphic(loopc).GrhIndex > 0 Then
                Map_Grh_Set x, y, MapData(x, y).Graphic(loopc).GrhIndex, loopc
                
                If LastGRH <> GrhData(MapData(x, y).Graphic(loopc).GrhIndex).FileNum Then
                    LastGRH = GrhData(MapData(x, y).Graphic(loopc).GrhIndex).FileNum
                    If Textures.CargarGrafico(LastGRH) Then DoEvents
                End If
            End If
            
        Next loopc
        
        MapData(x, y).Trigger = buffer(Idx).Trigger
        
        MapData(x, y).light_value(0) = buffer(Idx).l1
        MapData(x, y).light_value(1) = buffer(Idx).l2
        MapData(x, y).light_value(2) = buffer(Idx).l3
        MapData(x, y).light_value(3) = buffer(Idx).l4

        If buffer(Idx).p1 > 0 Then
            MapData(x, y).particle_group_index = General_Particle_Create(buffer(Idx).p1, x, y)
        End If
        
        Idx = Idx + 1
        
    Next x
Next y

Close #1

Load_Map_Data = True

End Function

Private Function Load_Inf_Data(MapRoute As String, Optional Mensajes As Boolean = True) As Boolean

On Error Resume Next

Dim loopc As Integer
Dim y As Integer
Dim x As Integer
Dim TempInt As Integer
Dim arboles As Long
Dim ItemIndex As Integer
Dim ItemAmount As Integer

Open mid$(MapRoute, 1, Len(MapRoute) - 4) & ".inf" For Binary As #2
Seek #2, 1
Get #2, , TempInt
Get #2, , TempInt
Get #2, , TempInt
Get #2, , TempInt
Get #2, , TempInt
For y = YMinMapSize To YMaxMapSize
    For x = XMinMapSize To XMaxMapSize
    
        Get #2, , MapData(x, y).TileExit.Map
        Get #2, , MapData(x, y).TileExit.x
        Get #2, , MapData(x, y).TileExit.y
    
        Get #2, , TempInt
        Map_NPC_Add x, y, TempInt
    
        Get #2, , ItemIndex
        Get #2, , ItemAmount
        Get #2, , TempInt
        Get #2, , TempInt
        
        If ItemIndex > 0 Then
            If Val(General_Var_Get(App.Path & "\init\obj.dat", "OBJ" & MapData(x, y).ObjInfo.OBJIndex, "Objtype")) = 4 Then
                If MapData(x, y).ObjGrh.GrhIndex = MapData(x, y).Graphic(3).GrhIndex Then
              
                    Map_Grh_UnSet x, y, 3
                    arboles = arboles + 1
                End If
            End If
            Map_Item_Add x, y, ItemIndex, ItemAmount
        End If
    Next x
Next y
Close #2

Load_Inf_Data = True

End Function
Private Function Load_Dat_Data(MapRoute As String, Optional Mensajes As Boolean = True) As Boolean
Dim C As String
Dim MapNum As Integer
Dim n As String

C = mid$(MapRoute, 1, Len(MapRoute) - 4)
MapNum = Val(mid(C, InStrRev(LCase(C), "mapa") + Len("mapa")))
If MapNum = 0 Then Exit Function

C = mid$(MapRoute, 1, Len(MapRoute) - 4) & ".dat"
'lee dat
MapInfo.Name = General_Var_Get(C, "Mapa" & MapNum, "Name")

MapInfo.Music = General_Var_Get(C, "Mapa" & MapNum, "MusicNum")
If Val(General_Var_Get(C, "Mapa" & MapNum, "Pk")) = 0 Then
    MapInfo.Pk = True
Else
    MapInfo.Pk = False
End If
MapInfo.Terreno = General_Var_Get(C, "Mapa" & MapNum, "Terreno")
MapInfo.Zona = General_Var_Get(C, "Mapa" & MapNum, "Zona")

MapInfo.SonidoAmbiente = General_Var_Get(C, "Mapa" & MapNum, "AmbNum")
MapInfo.Restringir = General_Var_Get(C, "Mapa" & MapNum, "Restringir")
MapInfo.BackUp = Val(General_Var_Get(C, "Mapa" & MapNum, "BackUp"))
MapInfo.base_light = Val(General_Var_Get(C, "Mapa" & MapNum, "LuzBase"))

If MapInfo.base_light = 0 Then
    map_base_light = &HFFFFFF
Else
    map_base_light = MapInfo.base_light
End If

Load_Dat_Data = True
End Function

Public Function Map_Save_To_File(MapRoute As String) As Boolean

If UCase(Right$(MapRoute, 4)) <> ".MAP" Then Exit Function
If Not Save_Map_Data(MapRoute) Then Exit Function
If Not Save_Inf_Data(MapRoute) Then Exit Function
If Not Save_Dat_Data(MapRoute) Then Exit Function
Map_Save_To_File = True
End Function

Private Function Save_Map_Data(MapRoute As String) As Boolean
Dim loopc As Integer
Dim y As Integer
Dim x As Integer
Dim TempInt As Integer

Open MapRoute For Binary As #1
Seek #1, 1
        
'map Header
Put #1, , TempInt
Put #1, , MiCabecera
Put #1, , TempInt
Put #1, , TempInt
Put #1, , TempInt
Put #1, , TempInt
'Load arrays
For y = YMinMapSize To YMaxMapSize
    For x = XMinMapSize To XMaxMapSize

        Put #1, , MapData(x, y).Blocked
        For loopc = 1 To 4
            Put #1, , CInt(MapData(x, y).Graphic(loopc).GrhIndex)
        Next loopc
        
        Put #1, , MapData(x, y).Trigger
        
        Put #1, , TempInt
        Put #1, , MapData(x, y).light_value(0)
        Put #1, , MapData(x, y).light_value(1)
        Put #1, , MapData(x, y).light_value(2)
        Put #1, , MapData(x, y).light_value(3)
        
        If MapData(x, y).particle_group_index > 0 Then
            Put #1, , CLng(particle_group_list(MapData(x, y).particle_group_index).stream_type)
        Else
            Put #1, , MapData(x, y).particle_group_index
        End If
    Next x
Next y

Close #1

Save_Map_Data = True

End Function

Private Function Save_Inf_Data(MapRoute As String) As Boolean
Dim loopc As Integer
Dim y As Integer
Dim x As Integer
Dim TempInt As Integer
Open mid$(MapRoute, 1, Len(MapRoute) - 4) & ".inf" For Binary As #2
Seek #2, 1
Put #2, , TempInt
Put #2, , TempInt
Put #2, , TempInt
Put #2, , TempInt
Put #2, , TempInt
For y = YMinMapSize To YMaxMapSize
    For x = XMinMapSize To XMaxMapSize
    
        Put #2, , MapData(x, y).TileExit.Map
        Put #2, , MapData(x, y).TileExit.x
        Put #2, , MapData(x, y).TileExit.y
    
        Put #2, , MapData(x, y).NPCIndex
    
        Put #2, , MapData(x, y).ObjInfo.OBJIndex
        Put #2, , MapData(x, y).ObjInfo.Amount
        Put #2, , TempInt
        Put #2, , TempInt
    Next x
Next y
Close #2
Save_Inf_Data = True

End Function

Private Function Save_Dat_Data(MapRoute As String) As Boolean
Dim C As String
Dim MapNum As Integer

C = mid$(MapRoute, 1, Len(MapRoute) - 4)
MapNum = Val(mid(C, InStrRev(LCase(C), "mapa") + Len("mapa")))
If MapNum = 0 Then Exit Function

C = mid$(MapRoute, 1, Len(MapRoute) - 4) & ".dat"
'escribe dat

General_Var_Write C, "Mapa" & MapNum, "Name", MapInfo.Name
General_Var_Write C, "Mapa" & MapNum, "MusicNum", MapInfo.Music
If MapInfo.Pk Then
    General_Var_Write C, "Mapa" & MapNum, "Pk", "1"
Else
    General_Var_Write C, "Mapa" & MapNum, "Pk", "0"
End If
General_Var_Write C, "Mapa" & MapNum, "Terreno", MapInfo.Terreno
General_Var_Write C, "Mapa" & MapNum, "Zona", MapInfo.Zona
General_Var_Write C, "Mapa" & MapNum, "AmbNum", MapInfo.SonidoAmbiente
General_Var_Write C, "Mapa" & MapNum, "Restringir", MapInfo.Restringir
General_Var_Write C, "Mapa" & MapNum, "BackUp", Trim(Str(MapInfo.BackUp))
General_Var_Write C, "Mapa" & MapNum, "LuzBase", Trim(Str(MapInfo.base_light))

Save_Dat_Data = True

End Function

Public Sub Fill_Grh_List(GrhList As ListBox)
'****************************
'Fill Grh List
'Author: Ryan Cain (Onezero)
'Last Modify Date: 5/14/2003
'****************************
Dim loopc As Long

For loopc = 1 To CANT_GRH_INDEX
    If GrhData(loopc).NumFrames > 0 Then
        GrhList.AddItem loopc
    End If
Next loopc

End Sub

Public Function Letter_Set(ByVal grh_index As Long, ByVal text_string As String) As Boolean
'*****************************************************************
'Author: Augusto José Rando
'*****************************************************************
    letter_text = text_string
    letter_grh.GrhIndex = grh_index
    letter_x_pos = 380
    letter_y_pos = 395
    Letter_Set = True
End Function

Public Function Letter_UnSet() As Boolean
'*****************************************************************
'Author: Augusto José Rando
'*****************************************************************
    letter_text = ""
    letter_grh.GrhIndex = 0
    letter_x_pos = 0
    letter_y_pos = 0
    Letter_UnSet = True
End Function
