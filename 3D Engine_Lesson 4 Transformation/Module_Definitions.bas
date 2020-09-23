Attribute VB_Name = "Module_Definitions"
Option Explicit
'==============================================================================================================
'
'       In this module we define all objects and types
'       we will use in the Engine like,3D Device object,iformations
'       like matricies ect....
'
'
'================================================================================================




'global objects

Public Const PI As Double = 3.14159265358979
Public Const PIdiv180 = PI / 180
Public Const RAD = PI / 180
Public Const PI_90 = PI / 2
Public Const PI_45 = PI / 4
Public Const PI_180 = PI
Public Const PI_360 = PI * 2
Public Const PI_270 = (PI / 2) * 3
Public Const QUEST3D_EPSILON = 0.0001

'for accessing to all functions provided by Directx Lib
  Global obj_DX As New DirectX8
  
'this object is an interface that provide functions and methods.
'this routines allow to check if the real 3D device has some required
'capabilities for a 3D engine
  Global obj_D3D As Direct3D8
  
  
'this engine is an interface that communicate
'directly with the 3D GFX Card
  Global obj_Device As Direct3DDevice8
  
  
  Global obj_D3DX As D3DX8



'=======================================================================
' here we define all type that will be required
'
'
'=======================================================================



Public Type QUEST3D_Frustrum
    Near As Single
    Far As Single
    FovAngle As Single
    Aspect As Single

End Type

Public Type QUEST3D_CFG
    
    'actual  screen width
    Buffer_Width As Integer
    'actual screen_height
    Buffer_Height As Integer
    'screen Rectangle (left,right,top,bottom values)
    Buffer_Rect As RECT
    'are we in windowed mode
    Is_Windowed As Boolean
    'dephtbit size
    Bpp As Integer
    'the engine is active
    Is_engineActive As Boolean
    'color for the back buffer
    BackBuff_ClearColor As Long
    
    GamaLevel As Single
    
   
    'for font
    MainFont As D3DXFont
    StFont As StdFont
    FontDesc As IFont
    
    
    'for view frustum
    ViewFrust As QUEST3D_Frustrum
    
     MatProjec As D3DMATRIX
    matView As D3DMATRIX

  
  
    'handle of the form or the windows interface
    Hwindow As Long
    
    
    'device creation parameters
    WinParam As D3DPRESENT_PARAMETERS
    
    
    'for frame counter
    Fps_CurrentTime As Single
    Fps_LastTime As Single
    Fps_FrameCounter As Single
    Fps_FramePerSecond As Single
    Fps_TimePassed As Single
   
End Type

Public Data As QUEST3D_CFG


'this is added for
'dialog based initialization

'this will received FMT format
Type tDFMT
    NumD As Long
    FMT() As Long
End Type


Type tFMT
    Wi As Long
    HI As Long
    'FMT As Long
    BK_FMT() As Long
    DP_FMT() As tDFMT

    SAMPLE() As Long
    NumBK As Long
    NumDPH As Long
    NumD() As Long

End Type

Type Tini
    RESO() As tFMT
    SelectINDEX As Long
    NumRES As Long
    CurrentWINDOWED As tFMT
    DISP_FMT() As Long
    CurrBKFMT As Long
    CurrINDEX As Long

End Type

Public TempDX8 As DirectX8          'The Root Object
Public TempD3D8 As Direct3D8      'The Direct3D Interface

Public nAdapters As Long 'How many adapters we found
Public AdapterInfo As D3DADAPTER_IDENTIFIER8 'A Structure holding information on the adapter

Public nModes As Long 'How many display modes we found
Public DEV As Tini
Public Type QUEST3D_CFG_INI

    Width As Integer
    Height As Integer
    format As Long
    USE_from_Dialog As Boolean
    MaxFramePerSec As Long
    USE_TnL As Boolean
    DeviceTyp As CONST_D3DDEVTYPE
    ForceVerSINC As Boolean
    appHandle As Long
    ChildHandle As Long
    IS_FullScreen As Boolean
    GamaLevel As Single
    Bpp As Integer
    BufferCount As Integer
    BK_FMT As Long
    DP_FMT As Long
    IS_OKAY As Long

End Type

'check if there is ERROR
Public IS_ERROR As Boolean

Public CFG As QUEST3D_CFG_INI

'some apis
'to retrieve keyboard state
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'to retrieve time
Declare Function timeGetTime Lib "winmm.dll" () As Long


Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long


'some types

'for rendering

' A structure for a simple Transformed and lighted vertex type
'
'Transformed and lighet means
'the vertice is a point on the screen
'
'===================================================
'                       |Y-
'                       |
'                       |
'                       |
'                       |
'                       |
'                       |
'                       |
'                       |
'                       |
'                       |
'X-_____________________x_________________________X+
'                       |
'                       |
'                       |
'                       |
'                       |
'                       |
'                       |Y+

'
'
'
' representing a point on the screen

Public Type QUEST3D_VERTEXCOLORED2D
    Position As D3DVECTOR
    'where
'    x As Single         'x in screen space
'    y As Single         'y in screen space
'    z  As Single        'normalized z
    rhw As Single       'normalized z rhw
    Color As Long       'vertex color
End Type

' Our custom FVF, which describes our custom vertex structure
Global Const QUEST3D_FVFVERTEXCOLORED2D = (D3DFVF_XYZRHW Or D3DFVF_DIFFUSE)


Public Type QUEST3D_VERTEXCOLORED3D
    Position As D3DVECTOR
    Color As Long       'vertex color
End Type

' Our custom FVF, which describes our custom vertex structure
Global Const QUEST3D_FVFVERTEXCOLORED3D = (D3DFVF_XYZ Or D3DFVF_DIFFUSE)

