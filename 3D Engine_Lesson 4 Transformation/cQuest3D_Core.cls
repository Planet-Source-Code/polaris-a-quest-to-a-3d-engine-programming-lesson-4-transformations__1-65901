VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cQuest3D_Core"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==============================================================================================================
'
'       THIS IS THE CORE CLASS OF THE ENGINE
'
'  - we start 3D and close 3D here
'  - We acces the GFX here
'  - we change all GFX state here
'================================================================================================


'Lesson 1
'
'We write
'  - an initialization method
'  - Begin3D to start 3d rendering
'  - End3D to close 3D drawing and flip all draw stuff to the screen
'  - Free to close engine
'  - GetKeyPressed() to get the current pressed key


'=====================================================
'Method:Init
' Create 3D device and all objects needed
' - Windowed for create a windowed device
' - FullScreenWidth and FullScreenheight screen size
' - DisableVerticalSync to disable monitor vertical refresh
'    if DisableVerticalSync=true the frames per second increase hugely

'Return: Boolean
' return true if the engine creation suceed
' return false if there was an error
'=====================================================
Function Init(ByVal FormHandle As Long, Optional ByVal winDowed As Boolean = 1, Optional ByVal FullScreenWidth As Integer = 640, Optional ByVal FullScreenHeight As Integer = 480, Optional ByVal DisableVerticalSync As Boolean = True) As Boolean




    Set obj_D3D = obj_DX.Direct3DCreate() 'we create the D3D object
    
    Set obj_D3DX = New D3DX8
    
    
     ' Get The current Display Mode format
    Dim mode As D3DDISPLAYMODE
    obj_D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, mode
    
    
    'we set the parameters for the creation of the Device
    
    
     If winDowed Then
        'for Windowed mode
        Data.WinParam.winDowed = 1
        Data.WinParam.BackBufferCount = 1 '1 backbuffer
        Data.WinParam.BackBufferFormat = mode.format 'color format
        Data.WinParam.SwapEffect = D3DSWAPEFFECT_FLIP

        If DisableVerticalSync = False Then Data.WinParam.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC

      
      Else
        'for fullscreen
        
        Data.Buffer_Rect.Left = 0
        Data.Buffer_Rect.Top = 0
        Data.Buffer_Rect.Right = FullScreenWidth
        Data.Buffer_Rect.bottom = FullScreenHeight

        Data.WinParam.BackBufferCount = 1
        Data.WinParam.BackBufferFormat = mode.format
        Data.WinParam.BackBufferWidth = FullScreenWidth
        Data.WinParam.BackBufferHeight = FullScreenHeight
        Data.WinParam.SwapEffect = D3DSWAPEFFECT_COPY

        If Not DisableVerticalSync = False Then
            Data.WinParam.FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_IMMEDIATE
          Else 'VERTSINCOFF = FALSE
            Data.WinParam.FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_DEFAULT
        End If
    End If
    
    Data.WinParam.EnableAutoDepthStencil = 1
    Data.WinParam.AutoDepthStencilFormat = D3DFMT_D16 '16 bit Z-Buffer
    Data.WinParam.hDeviceWindow = FormHandle  'target
    
    
    
    
    
    

    'create the device
      'D3DADAPTER_DEFAULT means we use default Graphic card VGA is the first
      'D3DCREATE_SOFTWARE_VERTEXPROCESSING we use directx driver software vertex processing
    
    Set obj_Device = obj_D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, FormHandle, D3DCREATE_SOFTWARE_VERTEXPROCESSING, Data.WinParam)

    'if the device is nothing we fail else Great WE DID IT nice!!
    Data.Is_engineActive = Not (obj_Device Is Nothing)
    
    'we prepare informations
    If Data.Is_engineActive Then
      DoDefault
    End If
    
    Init = Data.Is_engineActive

End Function




Function Init_Dialogue(ByVal appHandle As Long) As Boolean


 frmEnum.SHOW_DIALOG appHandle

 If IS_ERROR Then
    Init_Dialogue = False
    Exit Function
 
 End If
    
    

 
        Set obj_DX = New DirectX8
        Set obj_D3DX = New D3DX8
        Set obj_D3D = obj_DX.Direct3DCreate() ' Create D3D
        
        

      Dim DispMode As D3DDISPLAYMODE 'for current displaymode
      Dim D3DWindow As D3DPRESENT_PARAMETERS 'for present parameters
      Dim RectClient As RECT

        
       
        GetWindowRect CFG.appHandle, RectClient
        If Not CFG.IS_FullScreen Then
            'inizializza finestra
            D3DWindow.winDowed = 1
            D3DWindow.BackBufferCount = 1 '1 backbuffer
            D3DWindow.BackBufferFormat = CFG.BK_FMT 'color format from dialogue
            D3DWindow.SwapEffect = D3DSWAPEFFECT_FLIP

            If CFG.ForceVerSINC Then D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC

            GetWindowRect CFG.appHandle, RectClient
            CFG.Width = RectClient.Right - RectClient.Left
            CFG.Height = RectClient.bottom - RectClient.Top
            Data.Buffer_Rect.Right = CFG.Width
            Data.Buffer_Rect.bottom = CFG.Height
          Else
           'fullscreen
            'GetWindowRect CFG.appHandle, Data.Buffer_Rect

            Data.Buffer_Rect.Left = 0
            Data.Buffer_Rect.Top = 0
            Data.Buffer_Rect.Right = CFG.Width
            Data.Buffer_Rect.bottom = CFG.Height

            D3DWindow.BackBufferCount = CFG.BufferCount  ' backbuffer
            D3DWindow.BackBufferFormat = CFG.BK_FMT 'colore
            D3DWindow.BackBufferWidth = CFG.Width
            D3DWindow.BackBufferHeight = CFG.Height
            D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY

            If Not CFG.ForceVerSINC Then
                D3DWindow.FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_IMMEDIATE
              Else 'VERTSINCOFF = FALSE
                D3DWindow.FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_DEFAULT
            End If
        End If
        'comuni
        D3DWindow.EnableAutoDepthStencil = 1
        D3DWindow.AutoDepthStencilFormat = CFG.DP_FMT 'CFG.BK_FMT '16 bit Z-Buffer
        D3DWindow.hDeviceWindow = CFG.appHandle  'target
        'If debugMode Then D3DWindow.FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_IMMEDIATE
        'If (GetStencilBits(CFG.DP_FMT) > 0) Then _
            'D3DWindow.Flags = D3DPRESENTFLAG_DISCARD_DEPTHSTENCIL

        'crea device
        

        If CFG.DeviceTyp = D3DDEVTYPE_HAL Then
            Set obj_Device = obj_D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, CFG.appHandle, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
          ElseIf CFG.USE_TnL Then
            Set obj_Device = obj_D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, CFG.appHandle, D3DCREATE_HARDWARE_VERTEXPROCESSING, D3DWindow)
          Else
            Set obj_Device = obj_D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_REF, CFG.appHandle, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)

        End If

    

     
        Data.Buffer_Width = CFG.Width
        Data.Buffer_Height = CFG.Height
        Data.Bpp = CFG.Bpp
        Data.Is_Windowed = Not CFG.IS_FullScreen
        Data.GamaLevel = CFG.GamaLevel
       
        Data.WinParam = D3DWindow
        
     Init_Dialogue = Not (obj_Device Is Nothing)
     
     If Init_Dialogue Then
       Data.Is_engineActive = True
       DoDefault
     End If

End Function





Private Sub DoDefault()
'here we set default parameters

'we set the color to blue
 Data.BackBuff_ClearColor = D3DColorXRGB(0, 0, 255)
 
 Me.Set_Font "arial", 12
 
 Me.Set_ViewFrustum 0, 2500, PI / 4, 1

End Sub

Sub FreeEngine()
  
    'we deallocate memory and free objects
    Set obj_DX = Nothing
    Set obj_D3D = Nothing
    Set obj_Device = Nothing
    
    Data.Is_engineActive = False

End Sub



'=======================================
'We tell D3D that we Begin 3D rendering
'
'======================================
Sub Begin3D()

   ComputeFrames

   If Data.Is_engineActive = False Then Exit Sub

  'we clear previous drawn
  obj_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, Data.BackBuff_ClearColor, 1#, 0
  'We tell the device we are ready to render triangles
  obj_Device.BeginScene
  
End Sub


'==========================================================
'Here we compute frame per second
'
'==========================================================
Private Sub ComputeFrames()
    Dim Elapsed As Single
    
    'we get the current time in second
    Data.Fps_TimePassed = timeGetTime * 0.001 - Data.Fps_CurrentTime
    Data.Fps_CurrentTime = timeGetTime * 0.001
    'we count Frames
    Data.Fps_FrameCounter = Data.Fps_FrameCounter + 1
    
    'we compute time passed during two rendering scene
    Elapsed = Data.Fps_CurrentTime - Data.Fps_LastTime

'if we are over 1 sec
If Elapsed > 1 Then
  'we reset the last time to Now
  Data.Fps_LastTime = Data.Fps_CurrentTime
  
  'we check that we have 1 second,1.004=85 is not true so we do (85 x 1)/1.004
  If Elapsed <> 0 Then _
  Data.Fps_FramePerSecond = Data.Fps_FrameCounter / (Elapsed)
  'we reset the counter
  Data.Fps_FrameCounter = 0

End If


End Sub


Function Get_D3DDevice() As Direct3DDevice8
    Set Get_D3DDevice = obj_Device
End Function

Function Get_FramesPerSeconde() As Single
        Get_FramesPerSeconde = Data.Fps_FramePerSecond
End Function

Function Get_TimePassed() As Single
        Get_TimePassed = Data.Fps_TimePassed
End Function

'=========================================
'We tell D3D that we Stop 3D rendering
'=========================================
Sub End3D()

   If Data.Is_engineActive = False Then Exit Sub

  obj_Device.EndScene 'We tell the device we'are now finished
  obj_Device.Present ByVal 0, ByVal 0, 0, ByVal 0 'we flip all drawn stuffs over the monitor' screen

End Sub


Function Get_KeyPressed(ByVal Key As KeyCodeConstants) As Boolean
   Get_KeyPressed = (GetAsyncKeyState(Key) <> 0)
End Function




Private Sub Class_Initialize()
  Randomize Timer
End Sub

Sub Set_Font(ByVal FontName As String, ByVal PoliceSize As Integer)

    Set Data.StFont = New StdFont

    Data.StFont.Name = FontName
    Data.StFont.Size = PoliceSize
    Set Data.FontDesc = Data.StFont

    Set Data.MainFont = obj_D3DX.CreateFont(obj_Device, Data.FontDesc.hFont)

End Sub


Sub Draw_Text(ByVal St As String, ByVal x As Long, ByVal y As Long, Optional ByVal Color As Long = &HFFFFCC00)

    Dim TextRect As RECT

    TextRect.Top = y

    TextRect.Left = x

   
    obj_D3DX.DrawText Data.MainFont, Color, St, TextRect, DT_TOP Or DT_LEFT
   

End Sub


Sub Draw_Vertex2D(ByVal VertexArray As Variant, ByVal PrimitiveCount As Long, Optional ByVal PrimitiveType As CONST_D3DPRIMITIVETYPE = D3DPT_TRIANGLELIST)
    
    
    Dim V As QUEST3D_VERTEX2D
    
    Set_EngineVertexShader QUEST3_FVFVERTEX2D
    obj_Device.DrawPrimitiveUP PrimitiveType, PrimitiveCount, VertexArray, Len(V)
    
End Sub

'============================================================
'Engine State
'
'============================================================
Sub Set_BackbufferClearColor(ByVal ClearColor As Long)
    Data.BackBuff_ClearColor = ClearColor
End Sub

Sub Set_EngineVertexShader(ByVal FVF_format As Long)

    Call obj_Device.SetVertexShader(FVF_format)

End Sub

Sub Set_EngineLight(ByVal Enable As Boolean)

    obj_Device.SetRenderState D3DRS_LIGHTING, Enable

End Sub

Sub Set_EngineCullMode(ByVal CullMode As CONST_D3DCULL)

    obj_Device.SetRenderState D3DRS_CULLMODE, CullMode

End Sub



'===============================================================
'
'Matrix state
'===============================================================
Sub Set_WorldRotate_Scale_TranslateEx(RotRadVec As D3DVECTOR, ScalVec As D3DVECTOR, TransVec As D3DVECTOR)

  Dim MatTemp As D3DMATRIX

    'D3DXMatrixIdentity MatTemp
    Call Matrix_Ret(MatTemp, ScalVec.x, ScalVec.y, ScalVec.z, RotRadVec.x, RotRadVec.y, RotRadVec.z, TransVec.x, TransVec.y, TransVec.z)
    obj_Device.SetTransform D3DTS_WORLD, MatTemp

End Sub

'----------------------------------------
'Name: Set_WorldRotate_Scale_TranslateEX
'Object: Set
'Event: WorldRotate_Scale_TranslateEX
'----------------------------------------
'----------------------------------------
'Name: Set_WorldRotate_Scale_TranslateEX
'Object: Set
'Event: WorldRotate_Scale_TranslateEX
'Description:
'----------------------------------------
Sub Set_WorldRotate_Scale_Translate(ByVal RadX As Single, ByVal RadY As Single, ByVal RadZ As Single, ByVal ScalX As Single, ByVal ScalY As Single, ByVal ScalZ As Single, ByVal TransX As Single, ByVal TransY As Single, ByVal TransZ As Single)

    Call Set_WorldRotate_Scale_TranslateEx(Vector(RadX, RadY, RadZ), Vector(ScalX, ScalY, ScalZ), Vector(TransX, TransY, TransZ))

End Sub


Sub Set_ViewFrustum(ByVal Near As Single, ByVal Far As Single, ByVal FovAngle As Single, Optional ByVal Aspect As Single = -1)

    If Aspect = -1 Then Aspect = Data.Buffer_Width / Data.Buffer_Height
    With Data.ViewFrust
        .Aspect = Aspect 'Data.Buffer_Height / Data.Buffer_Width
        .Near = Near
        .Far = Far
        .FovAngle = FovAngle

    End With

    D3DXMatrixPerspectiveFovLH Data.MatProjec, Data.ViewFrust.FovAngle, Data.ViewFrust.Aspect, Data.ViewFrust.Near, Data.ViewFrust.Far
    
    obj_Device.SetTransform D3DTS_PROJECTION, Data.MatProjec

End Sub

'================================
