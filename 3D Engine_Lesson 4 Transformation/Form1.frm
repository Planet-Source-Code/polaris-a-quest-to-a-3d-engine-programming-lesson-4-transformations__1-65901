VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "3D Engine_Lesson 4 Transformation"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================================================
'WELCOME Engine_Lesson 2 Advanced Initialization
'_____________________________________________________________________________________
'-------------------------------------------------------------------------------------
'
'===================================================================================
'Welcome to this Step by step Quest to a 3D Engine programming
'this tutorial will show you how to design a simple 3D
'engine, Next tutorials will show how to add other engine objet
'like Camera,Mesh and Object Polygon
'
'This tutorial 4: 3D Engine_Lesson 4 Transformation
'
'It shows you how to initialize a 3D device
'  - how to render polygon to screen
'  - Apply a matrix tranformation to a polygon or an object
'
'
'
'How to read this code
'   - Form1: is the engine code in action
'   - Module_definitions will hold all engine objets definitions and types
'   - Module_Util will hold all Vector,matrix,Color math stuff
'   - cQuest3D_Core is our first object, it defines Main entry of the engine
'   - frmEnum have code for Device anumeration
'
'Good coding
'
'Vote if you want the sequel!!
'
'==================================================================================

Option Explicit





'we use the engine here
'we declare an objet
Dim QUEST As cQuest3D_Core


Private Sub Form_Load()

        'we allocate memory here
        Set QUEST = New cQuest3D_Core
        
      
        
        'we initialize the engine
        If QUEST.Init_Dialogue(Me.hwnd) = False Then
         MsgBox "Sorry there was an error"
         End
        End If
        
        Me.Refresh
        Me.Show
        
        'we call game loop
        GameLoop

End Sub

Sub GameLoop()

Dim TRIANGLE_VERT(2) As QUEST3D_VERTEXCOLORED3D

    'we define an Center oriented tringle unit 1
    TRIANGLE_VERT(0) = Make_Vertex3D(-1, -1, 0, Make_Color(25, 255, 0))
    TRIANGLE_VERT(1) = Make_Vertex3D(0, 1, 0, Make_Color(255, 55, 0))
    TRIANGLE_VERT(2) = Make_Vertex3D(1, -1, 0, Make_Color(0, 25, 255))
    



    'we tell the Engine we do not want Lighting
    'because we use a lighted polygon
    QUEST.Set_EngineLight False
    'we do not cull triangle
    'because we want to see our triangle rotating
    QUEST.Set_EngineCullMode D3DCULL_NONE


    'this code is added for camera position
    'we will used better code
    'in next tutorials
    Dim matView As D3DMATRIX
    D3DXMatrixLookAtLH matView, Vector(0#, 0#, -3#), _
                                 Vector(0#, 0#, 0#), _
                                 Vector(0#, 1#, 0#)
                                 
'    D3DXMatrixLookAtLH matView, Vector(0#, 0#, -3#)=Camera position _
'                                 Vector(0#, 0#, 0#)=Camera LookAt_
'                                 Vector(0#, 1#, 0#)=Camera Up vector

    QUEST.Get_D3DDevice.SetTransform D3DTS_VIEW, matView
    
    Dim AngleStep As Single



    Do
          'we add a share value angle
          'we used QUEST.Get_TimePassed to get Frames based animation
          AngleStep = AngleStep + (1) * QUEST.Get_TimePassed
          If AngleStep > PI * 2 Then AngleStep = 0
          
          'change the clear color randomely
          If QUEST.Get_KeyPressed(vbKeySpace) Then QUEST.Set_BackbufferClearColor (D3DColorXRGB(Rnd * 255, Rnd * 255, Rnd * 255))
          'we begin 3D
          QUEST.Begin3D
             
           
              'we tell Direct3D we want to use UnTransformed and lighted rendering
              'code provided by the driver
              
              'here we used matrix rotation provided by the Core interface
              'QUEST.Set_WorldRotate_Scale_Translate Rotate(X,Y,Z),Scale(X,Y,Z),Translate(X,Y,Z)
    
            
              QUEST.Set_WorldRotate_Scale_Translate 0, AngleStep, 0, 1, 1, 1, 0, 0, 0
           
              'we used the approppriate vertex Format
              QUEST.Get_D3DDevice.SetVertexShader QUEST3D_FVFVERTEXCOLORED3D
              'we send all vertices here
              QUEST.Get_D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 1, TRIANGLE_VERT(0), Len(TRIANGLE_VERT(0))
    
              'draw FPS
              QUEST.Draw_Text "FPS=" + CStr(QUEST.Get_FramesPerSeconde), 1, 10, &HFFFFFFFF
              QUEST.Draw_Text "Press Space to change Back color Randomly", 1, 25, &HFFFFFFFF
              QUEST.Draw_Text "Press ESC key to quit", 1, 40, &HFFFFFF00
              
          'we close 3D Drawing
          QUEST.End3D
          DoEvents
          
         If QUEST.Get_KeyPressed(vbKeyEscape) Then Call CloseGame
    Loop

End Sub

'we quit game here
Sub CloseGame()
  QUEST.FreeEngine
  End
End Sub

Private Sub Form_Unload(Cancel As Integer)
 CloseGame
End Sub
