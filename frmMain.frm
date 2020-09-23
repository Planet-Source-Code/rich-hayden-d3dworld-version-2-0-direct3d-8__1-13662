VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H0000FF00&
   Caption         =   "Richard Hayden's D3DWorld"
   ClientHeight    =   13065
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   15240
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   13065
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   14775
      Left            =   0
      ScaleHeight     =   14715
      ScaleWidth      =   14955
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      Begin VB.Timer tmrFps 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   5760
         Top             =   3120
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'THIS PROGRAM MAY NOT WORK ON MACHINES WITHOUT GOOD GRAPHICS CARDS, IF IT DOESN'T THEN YOU CAN TRY CHANGING, THE FOLLOWING LINE:
'Set g_D3DDevice = g_D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, d3dpp)
'TO
'Set g_D3DDevice = g_D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)
'THIS SHOULD FIX IT, ALTHOUGH PERFORMANCE MAY BE POOR, AND/OR BUGGY
'
'Version 2.0.0 of Richard Hayden's D3DWorld is a great improvement on version 1 (D3DScene).
'Lighting is now used in many shapes and forms to enhance this scene. I have also improved
'the program, making it more efficient and it runs slightly faster.
'I have added a church-style building to the world (complete with stained-glass windows,
'which looks pretty effective.
'Six times of day are now simulated realistically using lighting. The screenshot below
'was taken when the time of day was set to evening.
'
'Next I hope to add collision detection and transparency and billboarding, to simulate objects
'like trees etc. If anyone can help with the collision detection and transparency areas, then
'please do. (r_hayden@breathemail.net). I also hope to make a more realistic sky.
'
'So look out for any proceeding versions!
'
'Please vote and/or provide feedback in return for me making this code available to you!
'
'D3DWorld Version 2.0.0 Copyright (c) 2000 Richard Hayden. All Rights Reserved.
'If you use any of this code in your programs, please acknowledge me in your code.
'
'I must acknowledge Simon Price (http://www.vbgames.co.uk) who has introduced me to D3D in Vb, with his excellent tutorials.
'
'Cheers, Simon!
'
'If anyone can help me with Collision Detection and/or transparency in textures, ie. colour keys etc., please e-mail me on r_hayden@breathemail.net

Option Explicit

Dim g_DX As New DirectX8            ' mother of it all
Dim g_D3DX As New D3DX8
Dim g_D3D As Direct3D8              ' used to create the D3DDevice
Dim g_D3DDevice As Direct3DDevice8  ' rendering device

Dim g_VertexBuffers(0 To 41) As Direct3DVertexBuffer8 'holds all my vertexbuffers
Dim g_Textures(0 To 8) As Direct3DTexture8 'holds all my textures

Dim jumping, crouching As Boolean 'boolean to tell whether the camera is jumping or crouching
Dim jUP As Boolean 'which way is the camera going, in terms of jumping; up or down

Dim di As DirectInput8 'this is DirectInput, used to monitor the keys on the keyboard in my case
Dim diDEV As DirectInputDevice8 'this device will be the keyboard
Dim diState As DIKEYBOARDSTATE 'to check the state of keys

Dim fps As Integer 'frames/sec

Dim Angle, AngleConv As Single 'holds the angle, at which the camera is pointing
Dim pitch As Single 'holds the pitch of the camera (this is where the camera is pointing in terms of the y axis, ie. up and down etc.)

Dim camz, camx, camy As Single 'hold the position of the camera on the x, y and z axis

' a structure for custom vertex type
Private Type CUSTOMVERTEX
    position As D3DVECTOR    '3d position for vertex
    Color As Long           'color of the vertex
    tu As Single            'texture map coordinate
    tv As Single            'texture map coordinate
End Type

' custom FVF, which describes our custom vertex structure
Const D3DFVF_CUSTOMVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)

Const g_pi As Single = 3.141592653 'pi
Const g_90d As Single = g_pi / 2 '90 degrees in radians
Const g_180d As Single = g_pi '180 degrees in radians
Const g_270d As Single = (g_pi / 2) * 3 '270 degrees in radians
Const g_360d As Single = g_pi * 2 '360 degrees in radians

Const TURN_SPEED = g_90d / 18 'camera turning speed
Const MOVE_SPEED = 0.5 'camera moving speed
Const FAST_MOVE_SPEED = 0.8 'fast camera moving speed
Const JUMP_MOVE_SPEED = 1.2 'jumping camera moving speed
Const JUMP_SPEED = 1 'jumping speed
Const PITCH_SPEED = 0.2 'look up and down speed
Const CROUCH_MOVE_SPEED = 0.2 'camera speed when crouching

Private Sub Form_Resize()
    'resize the picbox to utilise full size of form
    Picture1.Width = frmMain.Width
    Picture1.Height = frmMain.Height
End Sub

Private Sub SortPix(strWhat As String)
    On Error Resume Next
    'create and delete pix from pic boxes
    If strWhat = "create" Then
        SavePicture frmTextures.picGrass.Picture, App.Path & "\grass.bmp"
        SavePicture frmTextures.picBricks.Picture, App.Path & "\bricks.bmp"
        SavePicture frmTextures.picSky.Picture, App.Path & "\sky.bmp"
        SavePicture frmTextures.picRoof.Picture, App.Path & "\roof.bmp"
        SavePicture frmTextures.picTile.Picture, App.Path & "\tile.bmp"
        SavePicture frmTextures.picGorilla.Picture, App.Path & "\gorilla.bmp"
        SavePicture frmTextures.picAsphalt.Picture, App.Path & "\asphalt.bmp"
        SavePicture frmTextures.picChurchBricks.Picture, App.Path & "\churchbricks.bmp"
        SavePicture frmTextures.picChurchWindow.Picture, App.Path & "\churchwindow.bmp"
    ElseIf strWhat = "delete" Then
        Kill App.Path & "\grass.bmp"
        Kill App.Path & "\bricks.bmp"
        Kill App.Path & "\sky.bmp"
        Kill App.Path & "\roof.bmp"
        Kill App.Path & "\tile.bmp"
        Kill App.Path & "\gorilla.bmp"
        Kill App.Path & "\asphalt.bmp"
        Kill App.Path & "\churchbricks.bmp"
        Kill App.Path & "\churchwindow.bmp"
    End If
    Err.Number = 0
End Sub

Private Sub Form_Load()
    Dim b As Boolean
    frmMain.Caption = "Richard Hayden's D3DWorld " & App.Major & "." & App.Minor & "." & App.Revision
    ' Allow the form to become visible
    DoEvents
    'make the pix
    SortPix "create"
    'starting position + angle
    camx = 0
    camy = 10
    camz = -1
    Angle = g_360d
    pitch = 0
    'set time of day and fill mode to defaults
    lngLightType = MIDDAY_LIGHT
    lngFillMode = D3DFILL_SOLID
    'maximising form
    frmMain.Width = Screen.Width
    frmMain.Height = Screen.Height
    frmMain.Top = 0
    frmMain.Left = 0
    Picture1.Width = frmMain.Width
    Picture1.Height = frmMain.Height
    Picture1.Top = 0
    Picture1.Left = 0
    Me.Show
    frmDetails.Show
    frmDetails.SetFocus
    'create directinput object
    Set di = g_DX.DirectInputCreate()
        
    If Err.Number <> 0 Then
        MsgBox "Error starting Direct Input, please make sure you have DirectX installed", vbApplicationModal
        End
    End If
        
    'create keyboard device
    Set diDEV = di.CreateDevice("GUID_SysKeyboard")
    'set common data format to keyboard
    diDEV.SetCommonDataFormat DIFORMAT_KEYBOARD
    diDEV.SetCooperativeLevel Me.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    
    diDEV.Acquire
        
    
    
    ' Initialize D3D and D3DDevice
    b = InitD3D(Picture1.hWnd)
    If Not b Then
        MsgBox "Unable to CreateDevice (see InitD3D() source for comments)"
        End
    End If
    
    
    ' Initialize vertex buffer with geometry and load texture
    b = InitGeometry()
    If Not b Then
        MsgBox "Unable to Create VertexBuffer"
        End
    End If
    
    
    'enabled fps timer to get the frames/second
    tmrFps.Enabled = True
    Do While 1
        DoEvents
        Render
    Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'let's cleanup
    Cleanup
    End
End Sub

Function InitD3D(hWnd As Long) As Boolean
    On Local Error Resume Next
    
    ' Create the D3D object
    Set g_D3D = g_DX.Direct3DCreate()
    If g_D3D Is Nothing Then Exit Function
    
    ' Get The current Display Mode format
    Dim Mode As D3DDISPLAYMODE
    g_D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Mode
         
    ' Set up the structure used to create the D3DDevice. Since we are now
    ' using more complex geometry, we will create a device with a zbuffer.
    ' the D3DFMT_D16 indicates we want a 16 bit z buffer.
    Dim d3dpp As D3DPRESENT_PARAMETERS
    d3dpp.Windowed = 1
    d3dpp.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
    d3dpp.BackBufferFormat = Mode.Format
    d3dpp.BackBufferCount = 1
    d3dpp.EnableAutoDepthStencil = 1
    d3dpp.AutoDepthStencilFormat = D3DFMT_D16

    ' Create the D3DDevice
    ' If you do not have hardware 3d acceleration. Enable the reference rasterizer
    ' using the DirectX control panel and change D3DDEVTYPE_HAL to D3DDEVTYPE_REF
    
    Set g_D3DDevice = g_D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, d3dpp)
    If g_D3DDevice Is Nothing Then Exit Function
    
    ' Device state is set here
    ' Turn off culling, so we see the front and back
    g_D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE

    ' Turn on the zbuffer
    g_D3DDevice.SetRenderState D3DRS_ZENABLE, 1
    
    ' Turn off lighting
    'g_D3DDevice.SetRenderState D3DRS_LIGHTING, 0
    
    g_D3DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID

    InitD3D = True
End Function

Sub SetupMatrices()
    Dim matView As D3DMATRIX
    Dim matRotation As D3DMATRIX
    Dim matPitch As D3DMATRIX
    Dim matLook As D3DMATRIX
    Dim matPos As D3DMATRIX
    Dim matWorld As D3DMATRIX
    Dim matProj As D3DMATRIX
    
    'setup world matrix
    D3DXMatrixIdentity matWorld
    g_D3DDevice.SetTransform D3DTS_WORLD, matWorld
    
    'get the state of the keyboard to distate
    diDEV.GetDeviceStateKeyboard diState
    
    'if turning left or right, then change angle accordingly
    If diState.Key(205) <> 0 Then
        Angle = Angle - TURN_SPEED
        If Angle < 0 Then
            Angle = g_360d - (-Angle)
        End If
    End If
    If diState.Key(203) <> 0 Then
        Angle = Angle + TURN_SPEED
        If Angle > g_360d Then
            Angle = 0 + (Angle - g_360d)
        End If
    End If
    'convert to correct angle system
    AngleConv = g_360d - Angle
    'move forward or backward if up or down keys are active
    If diState.Key(200) <> 0 Then
        'check whats happening (ie. jumping or crouching) and move forwards at the correct speed
        If jumping Then
            camx = camx + (Sin(AngleConv) * JUMP_MOVE_SPEED)
            camz = camz + (Cos(AngleConv) * JUMP_MOVE_SPEED)
        ElseIf crouching Then
            camx = camx + (Sin(AngleConv) * CROUCH_MOVE_SPEED)
            camz = camz + (Cos(AngleConv) * CROUCH_MOVE_SPEED)
        Else
            If (diState.Key(42) <> 0) Or (diState.Key(54) <> 0) Then
                camx = camx + (Sin(AngleConv) * FAST_MOVE_SPEED)
                camz = camz + (Cos(AngleConv) * FAST_MOVE_SPEED)
            Else
                camx = camx + (Sin(AngleConv) * MOVE_SPEED)
                camz = camz + (Cos(AngleConv) * MOVE_SPEED)
            End If
        End If
    End If
    If diState.Key(208) <> 0 Then
        'check whats happening (ie. jumping or crouching) and move backwards at the correct speed
        If jumping Then
            camx = camx - (Sin(AngleConv) * JUMP_MOVE_SPEED)
            camz = camz - (Cos(AngleConv) * JUMP_MOVE_SPEED)
        ElseIf crouching Then
            camx = camx - (Sin(AngleConv) * CROUCH_MOVE_SPEED)
            camz = camz - (Cos(AngleConv) * CROUCH_MOVE_SPEED)
        Else
            If (diState.Key(42) <> 0) Or (diState.Key(54) <> 0) Then
                camx = camx - (Sin(AngleConv) * FAST_MOVE_SPEED)
                camz = camz - (Cos(AngleConv) * FAST_MOVE_SPEED)
            Else
                camx = camx - (Sin(AngleConv) * MOVE_SPEED)
                camz = camz - (Cos(AngleConv) * MOVE_SPEED)
            End If
        End If
    End If
    'if pressing page up or down then look up or down, respectively
    If diState.Key(201) <> 0 Then
        pitch = pitch + PITCH_SPEED
    End If
    If diState.Key(209) <> 0 Then
        pitch = pitch - PITCH_SPEED
    End If
    'if press space then let's jump!
    If diState.Key(57) <> 0 Then
        If Not jumping Then
            jumping = True
            jUP = True
        End If
    End If
    
    'crouching stuff v. simple...
    If (diState.Key(29) <> 0) Or (diState.Key(157) <> 0) Then
        camy = 10 - 4
        crouching = True
        jumping = False
    Else
        If Not jumping Then
            camy = 10
            crouching = False
        End If
    End If
    
    'this all does the jumping stuff, quite simple......
    If jumping Then
        If jUP = False Then
            camy = camy - JUMP_SPEED
            If camy <= 10 Then
                jumping = False
                camy = 10
            End If
        Else
            camy = camy + JUMP_SPEED
            If camy >= 20 Then
                jUP = False
            End If
        End If
    End If

    'make them identity matrices
    D3DXMatrixIdentity matView
    D3DXMatrixIdentity matPos
    D3DXMatrixIdentity matRotation
    D3DXMatrixIdentity matLook
    'rotate around x and y, for angle and pitch
    D3DXMatrixRotationY matRotation, Angle
    D3DXMatrixRotationX matPitch, pitch
    'multiply angle and pitch matrices together to create one 'look' matrix
    D3DXMatrixMultiply matLook, matRotation, matPitch
    'put the position of the camera into the translation matrix, matPos
    D3DXMatrixTranslation matPos, -camx, -camy, -camz
    'multiply that with the look matrix to make the complete view matrix
    D3DXMatrixMultiply matView, matPos, matLook
    'which we can then set as the view matrix:
    g_D3DDevice.SetTransform D3DTS_VIEW, matView
    'update details form
    frmDetails.Label1.Caption = "camx: " & camx & Chr(13) & "camy: " & camy & Chr(13) & "camz: " & camz & Chr(13) & "angleconv: " & AngleConv & Chr(13) & "cos(AngleConv) = " & Cos(AngleConv) & Chr(13) & "sin(AngleConv) = " & Sin(AngleConv) & Chr(13) & "pitch = " & pitch

    'setup the projection matrix
    D3DXMatrixPerspectiveFovLH matProj, g_pi / 3, 1, 1, 10000
    g_D3DDevice.SetTransform D3DTS_PROJECTION, matProj
End Sub

Sub SetupLights()
     
    Dim col As D3DCOLORVALUE
    
    
    ' Set up a material. The material here just has the diffuse and ambient
    ' colors set to yellow. Note that only one material can be used at a time.
    Dim mtrl As D3DMATERIAL8
    With col:    .r = 1: .g = 1: .b = 1: .a = 1:   End With
    mtrl.diffuse = col
    mtrl.Ambient = col
    g_D3DDevice.SetMaterial mtrl
    
    'set up the sun/moon light. Should be quite easy to understand.
    Dim light As D3DLIGHT8
    Dim lightPosVec As D3DVECTOR
    With lightPosVec
        .X = 999: .Y = 999: .z = -999
    End With

    
    If lngLightType = NIGHT_LIGHT Or lngLightType = EARLYMORNING_LIGHT Then
        light.Type = D3DLIGHT_POINT
        light.diffuse.r = 0.5
        light.diffuse.g = 0.5
        light.diffuse.b = 0.5
        light.position = lightPosVec
        light.Range = 50000#
    Else
        light.Type = D3DLIGHT_POINT
        light.diffuse.r = 1000#
        light.diffuse.g = 1000#
        light.diffuse.b = 0#
        light.position = lightPosVec
        light.Range = 50000#
    End If
    
    If frmDetails.Check1.Value = 1 Then
        g_D3DDevice.SetLight 0, light                   'let d3d know about the light
        g_D3DDevice.LightEnable 0, 1                    'turn it on
    Else
        g_D3DDevice.SetLight 0, light                   'let d3d know about the light
        g_D3DDevice.LightEnable 0, 0                    'turn it on
    End If
    
    'set up the light in the house. Should also be quite easy to understand.
    Dim houseLight As D3DLIGHT8
    Dim hlightPosVec As D3DVECTOR
    With hlightPosVec
        .X = 25: .Y = 29: .z = 25
    End With

    houseLight.Type = D3DLIGHT_POINT
    houseLight.diffuse.r = 1#
    houseLight.diffuse.g = 1#
    houseLight.diffuse.b = 1#
    houseLight.position = hlightPosVec
    houseLight.Range = 50#
    
    If frmDetails.Check2.Value = 1 Then
        g_D3DDevice.SetLight 1, houseLight                   'let d3d know about the light
        g_D3DDevice.LightEnable 1, 1
    Else
        g_D3DDevice.SetLight 1, houseLight                   'let d3d know about the light
        g_D3DDevice.LightEnable 1, 0
    End If
    
    'set up weird search light thing.
    Dim searchLight As D3DLIGHT8
    Dim slightDirVec As D3DVECTOR
    With slightDirVec
        .X = Cos(Timer * 2): .Y = 1: .z = Sin(Timer * 2)
    End With
    
    searchLight.Type = D3DLIGHT_DIRECTIONAL
    searchLight.diffuse.r = searchLightIntensity
    searchLight.diffuse.g = searchLightIntensity
    searchLight.diffuse.b = searchLightIntensity
    searchLight.Direction = slightDirVec
    searchLight.Range = 10#
    
    If frmDetails.Check3.Value = 1 Then
        g_D3DDevice.SetLight 2, searchLight                   'let d3d know about the light
        g_D3DDevice.LightEnable 2, 1
    Else
        g_D3DDevice.SetLight 2, searchLight                   'let d3d know about the light
        g_D3DDevice.LightEnable 2, 0
    End If
    
'this light doesn't work for some reason (well not the way I want it to; it should be like a torch, pointing in front of the camera), if anyone can fix it, your welcome............
'    Dim torchLight As D3DLIGHT8
'    Dim tlightDirVec As D3DVECTOR
'    With tlightDirVec
'        .X = Sin(camx): .Y = camy: .z = Cos(camz)
'    End With
'    Dim tlightPosVec As D3DVECTOR
'    With tlightPosVec
'        .X = camx: .Y = camy: .z = camz
'    End With
'
'    torchLight.Type = D3DLIGHT_SPOT
'    torchLight.diffuse.r = torchLightIntensity
'    torchLight.diffuse.g = torchLightIntensity
'    torchLight.diffuse.b = torchLightIntensity
'    torchLight.Direction = tlightDirVec
'    torchLight.position = tlightPosVec
'    torchLight.Falloff = 1
'    torchLight.Attenuation0 = 1
'    torchLight.Attenuation1 = 0
'    torchLight.Attenuation2 = 0
'    torchLight.Phi = 3.141592653
'    torchLight.Theta = 1.5
'    torchLight.Range = 10#
'
'
'    'If frmDetails.Check4.Value = 1 Then
'     '   g_D3DDevice.SetLight 3, torchLight                   'let d3d know about the light
'      '  g_D3DDevice.LightEnable 3, 1
'   ' Else
'    '    g_D3DDevice.SetLight 3, torchLight                  'let d3d know about the light
'     '   g_D3DDevice.LightEnable 3, 0
'    'End If
    
    'eerie green church aura
    Dim churchLight As D3DLIGHT8
    Dim clightPosVec As D3DVECTOR
    With clightPosVec
        .X = -190: .Y = 19: .z = 25
    End With

    churchLight.Type = D3DLIGHT_POINT
    churchLight.diffuse.r = 0
    churchLight.diffuse.g = 100
    churchLight.diffuse.b = 0
    churchLight.position = clightPosVec
    churchLight.Range = 50#
    
    If frmDetails.Check5.Value = 1 Then
        g_D3DDevice.SetLight 4, churchLight                   'let d3d know about the light
        g_D3DDevice.LightEnable 4, 1
    Else
        g_D3DDevice.SetLight 4, churchLight                   'let d3d know about the light
        g_D3DDevice.LightEnable 4, 0
    End If
    
    g_D3DDevice.SetRenderState D3DRS_LIGHTING, 1 'turn on lighting
    
    ' Finally, turn on some ambient light.
    ' Ambient light is light that scatters and lights all objects evenly
    g_D3DDevice.SetRenderState D3DRS_AMBIENT, lngLightType
    
End Sub

Function InitGeometry() As Boolean
    Dim i As Long
    
    'setup textures
    Set g_Textures(0) = g_D3DX.CreateTextureFromFile(g_D3DDevice, App.Path & "\bricks.bmp")
    If g_Textures(0) Is Nothing Then Exit Function
    Set g_Textures(1) = g_D3DX.CreateTextureFromFile(g_D3DDevice, App.Path & "\grass.bmp")
    If g_Textures(1) Is Nothing Then Exit Function
    Set g_Textures(2) = g_D3DX.CreateTextureFromFile(g_D3DDevice, App.Path & "\sky.bmp")
    If g_Textures(2) Is Nothing Then Exit Function
    Set g_Textures(3) = g_D3DX.CreateTextureFromFile(g_D3DDevice, App.Path & "\roof.bmp")
    If g_Textures(3) Is Nothing Then Exit Function
    Set g_Textures(4) = g_D3DX.CreateTextureFromFile(g_D3DDevice, App.Path & "\tile.bmp")
    If g_Textures(4) Is Nothing Then Exit Function
    Set g_Textures(5) = g_D3DX.CreateTextureFromFile(g_D3DDevice, App.Path & "\gorilla.bmp")
    If g_Textures(5) Is Nothing Then Exit Function
    Set g_Textures(6) = g_D3DX.CreateTextureFromFile(g_D3DDevice, App.Path & "\asphalt.bmp")
    If g_Textures(6) Is Nothing Then Exit Function
    Set g_Textures(7) = g_D3DX.CreateTextureFromFile(g_D3DDevice, App.Path & "\churchbricks.bmp")
    If g_Textures(7) Is Nothing Then Exit Function
    Set g_Textures(8) = g_D3DX.CreateTextureFromFile(g_D3DDevice, App.Path & "\churchwindow.bmp")
    If g_Textures(8) Is Nothing Then Exit Function
    
    'create an array to hold the vertex values temporarily, until added to buffer
    Dim Vertices(0 To 3) As CUSTOMVERTEX
    Dim VertexSizeInBytes As Long
    'get the size of a vertex
    VertexSizeInBytes = Len(Vertices(0))

    'create the grass or floor vertex buffer
    
    Vertices(0).position = vec3(-1000, 2, 1000)
    Vertices(1).position = vec3(1000, 2, 1000)
    Vertices(2).position = vec3(-1000, 2, -1000)
    Vertices(3).position = vec3(1000, 2, -1000)
    Vertices(0).Color = &H8080FF
    Vertices(1).Color = &HFF6FFFFF
    Vertices(2).Color = &HFFFF80
    Vertices(3).Color = &HFFC0FF
    Vertices(0).tu = 0
    Vertices(1).tu = 100
    Vertices(2).tu = 0
    Vertices(3).tu = 100
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 100
    Vertices(3).tv = 100

    ' Create the vertex buffer.
    Set g_VertexBuffers(0) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(0) Is Nothing Then Exit Function

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData g_VertexBuffers(0), 0, VertexSizeInBytes * 4, 0, Vertices(0)

    'create the sky vertex buffers
    
    Vertices(0).position = vec3(-1000, 1000, 1000)
    Vertices(1).position = vec3(1000, 1000, 1000)
    Vertices(2).position = vec3(-1000, 2, 1000)
    Vertices(3).position = vec3(1000, 2, 1000)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 1
    Vertices(2).tu = 0
    Vertices(3).tu = 1
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 1
    Vertices(3).tv = 1

    ' Create the vertex buffer.
    Set g_VertexBuffers(1) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(1) Is Nothing Then Exit Function

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData g_VertexBuffers(1), 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(-1000, 1000, -1000)
    Vertices(1).position = vec3(1000, 1000, -1000)
    Vertices(2).position = vec3(-1000, 2, -1000)
    Vertices(3).position = vec3(1000, 2, -1000)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 1
    Vertices(2).tu = 0
    Vertices(3).tu = 1
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 1
    Vertices(3).tv = 1

    ' Create the vertex buffer.
    Set g_VertexBuffers(2) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(2) Is Nothing Then Exit Function

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData g_VertexBuffers(2), 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(1000, 1000, -1000)
    Vertices(1).position = vec3(1000, 1000, 1000)
    Vertices(2).position = vec3(1000, 2, -1000)
    Vertices(3).position = vec3(1000, 2, 1000)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 1
    Vertices(2).tu = 0
    Vertices(3).tu = 1
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 1
    Vertices(3).tv = 1

    ' Create the vertex buffer.
    Set g_VertexBuffers(3) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(3) Is Nothing Then Exit Function

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData g_VertexBuffers(3), 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(-1000, 1000, -1000)
    Vertices(1).position = vec3(-1000, 1000, 1000)
    Vertices(2).position = vec3(-1000, 2, -1000)
    Vertices(3).position = vec3(-1000, 2, 1000)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 1
    Vertices(2).tu = 0
    Vertices(3).tu = 1
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 1
    Vertices(3).tv = 1

    ' Create the vertex buffer.
    Set g_VertexBuffers(4) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(4) Is Nothing Then Exit Function

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData g_VertexBuffers(4), 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(1000, 1000, -1000)
    Vertices(1).position = vec3(-1000, 1000, -1000)
    Vertices(2).position = vec3(1000, 1000, 1000)
    Vertices(3).position = vec3(-1000, 1000, 1000)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 1
    Vertices(2).tu = 0
    Vertices(3).tu = 1
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 1
    Vertices(3).tv = 1

    ' Create the vertex buffer.
    Set g_VertexBuffers(5) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(5) Is Nothing Then Exit Function

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData g_VertexBuffers(5), 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    'create the house vertex buffers
    
    Vertices(0).position = vec3(10, 20, 10)
    Vertices(1).position = vec3(40, 20, 10)
    Vertices(2).position = vec3(10, 2, 10)
    Vertices(3).position = vec3(40, 2, 10)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 10
    Vertices(2).tu = 0
    Vertices(3).tu = 10
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 10
    Vertices(3).tv = 10

    ' Create the vertex buffer.
    Set g_VertexBuffers(6) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(6) Is Nothing Then Exit Function

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData g_VertexBuffers(6), 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(10, 20, 40)
    Vertices(1).position = vec3(40, 20, 40)
    Vertices(2).position = vec3(10, 2, 40)
    Vertices(3).position = vec3(40, 2, 40)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 10
    Vertices(2).tu = 0
    Vertices(3).tu = 10
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 10
    Vertices(3).tv = 10

    ' Create the vertex buffer.
    Set g_VertexBuffers(7) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(7) Is Nothing Then Exit Function

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData g_VertexBuffers(7), 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(10, 20, 10)
    Vertices(1).position = vec3(10, 20, 40)
    Vertices(2).position = vec3(10, 2, 10)
    Vertices(3).position = vec3(10, 2, 40)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 10
    Vertices(2).tu = 0
    Vertices(3).tu = 10
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 10
    Vertices(3).tv = 10

    ' Create the vertex buffer.
    Set g_VertexBuffers(8) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(8) Is Nothing Then Exit Function

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData g_VertexBuffers(8), 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(40, 20, 10)
    Vertices(1).position = vec3(40, 20, 40)
    Vertices(2).position = vec3(40, 2, 10)
    Vertices(3).position = vec3(40, 2, 40)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 10
    Vertices(2).tu = 0
    Vertices(3).tu = 10
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 10
    Vertices(3).tv = 10

    ' Create the vertex buffer.
    Set g_VertexBuffers(9) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(9) Is Nothing Then Exit Function

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData g_VertexBuffers(9), 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    'create the roof vertex buffers
    
    Vertices(0).position = vec3(10, 30, 25)
    Vertices(1).position = vec3(40, 30, 25)
    Vertices(2).position = vec3(10, 20, 10)
    Vertices(3).position = vec3(40, 20, 10)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 10
    Vertices(2).tu = 0
    Vertices(3).tu = 10
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 10
    Vertices(3).tv = 10

    ' Create the vertex buffer.
    Set g_VertexBuffers(10) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(10) Is Nothing Then Exit Function

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData g_VertexBuffers(10), 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(10, 30, 25)
    Vertices(1).position = vec3(40, 30, 25)
    Vertices(2).position = vec3(10, 20, 40)
    Vertices(3).position = vec3(40, 20, 40)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 10
    Vertices(2).tu = 0
    Vertices(3).tu = 10
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 10
    Vertices(3).tv = 10

    ' Create the vertex buffer.
    Set g_VertexBuffers(11) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(11) Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_VertexBuffers(11), 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(10, 30, 25)
    Vertices(1).position = vec3(10, 20, 10)
    Vertices(2).position = vec3(10, 20, 40)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(0).tu = 2
    Vertices(1).tu = 4
    Vertices(2).tu = 6
    Vertices(0).tv = 2
    Vertices(1).tv = 0
    Vertices(2).tv = 0

    ' Create the vertex buffer.
    Set g_VertexBuffers(12) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 3, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(12) Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_VertexBuffers(12), 0, VertexSizeInBytes * 3, 0, Vertices(0)
    
    Vertices(0).position = vec3(40, 30, 25)
    Vertices(1).position = vec3(40, 20, 10)
    Vertices(2).position = vec3(40, 20, 40)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(0).tu = 2
    Vertices(1).tu = 4
    Vertices(2).tu = 6
    Vertices(0).tv = 2
    Vertices(1).tv = 0
    Vertices(2).tv = 0

    ' Create the vertex buffer.
    Set g_VertexBuffers(13) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 3, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(13) Is Nothing Then Exit Function

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData g_VertexBuffers(13), 0, VertexSizeInBytes * 3, 0, Vertices(0)
    
    'create the house floor vertex buffer
    
    Vertices(0).position = vec3(40, 2.1, 10)
    Vertices(1).position = vec3(10, 2.1, 10)
    Vertices(2).position = vec3(40, 2.1, 40)
    Vertices(3).position = vec3(10, 2.1, 40)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 10
    Vertices(2).tu = 0
    Vertices(3).tu = 10
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 10
    Vertices(3).tv = 10

    ' Create the vertex buffer.
    Set g_VertexBuffers(14) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(14) Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_VertexBuffers(14), 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    'create the gorilla pic vertex buffer
    
    Vertices(0).position = vec3(39.9, 15, 22.5)
    Vertices(1).position = vec3(39.9, 15, 27.5)
    Vertices(2).position = vec3(39.9, 9, 22.5)
    Vertices(3).position = vec3(39.9, 9, 27.5)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 1
    Vertices(2).tu = 0
    Vertices(3).tu = 1
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 1
    Vertices(3).tv = 1

    ' Create the vertex buffer.
    Set g_VertexBuffers(15) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(15) Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_VertexBuffers(15), 0, VertexSizeInBytes * 4, 0, Vertices(0)

    'create the road vertex buffer

    Vertices(0).position = vec3(-20, 2.1, -1000)
    Vertices(1).position = vec3(0, 2.1, -1000)
    Vertices(2).position = vec3(-20, 2.1, 1000)
    Vertices(3).position = vec3(0, 2.1, 1000)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 5
    Vertices(2).tu = 0
    Vertices(3).tu = 5
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 1000
    Vertices(3).tv = 1000

    ' Create the vertex buffer.
    Set g_VertexBuffers(16) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(16) Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_VertexBuffers(16), 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    'create the road sides vertex buffers
    
    Vertices(0).position = vec3(-20, 2.2, -1000)
    Vertices(1).position = vec3(-19, 2.2, -1000)
    Vertices(2).position = vec3(-20, 2.2, 1000)
    Vertices(3).position = vec3(-19, 2.2, 1000)
    Vertices(0).Color = &H0&
    Vertices(1).Color = &H0&
    Vertices(2).Color = &H0&
    Vertices(3).Color = &H0&

    ' Create the vertex buffer.
    Set g_VertexBuffers(17) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(17) Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_VertexBuffers(17), 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(-1, 2.2, -1000)
    Vertices(1).position = vec3(0, 2.2, -1000)
    Vertices(2).position = vec3(-1, 2.2, 1000)
    Vertices(3).position = vec3(0, 2.2, 1000)
    Vertices(0).Color = &H0&
    Vertices(1).Color = &H0&
    Vertices(2).Color = &H0&
    Vertices(3).Color = &H0&

    ' Create the vertex buffer.
    Set g_VertexBuffers(18) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(18) Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_VertexBuffers(18), 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    'create the road middle vertex buffer
    
    Vertices(0).position = vec3(-9.5, 2.2, -1000)
    Vertices(1).position = vec3(-10.5, 2.2, -1000)
    Vertices(2).position = vec3(-9.5, 2.2, 1000)
    Vertices(3).position = vec3(-10.5, 2.2, 1000)
    Vertices(0).Color = &HC0C0C0
    Vertices(1).Color = &HC0C0C0
    Vertices(2).Color = &HC0C0C0
    Vertices(3).Color = &HC0C0C0

    ' Create the vertex buffer.
    Set g_VertexBuffers(19) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(19) Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_VertexBuffers(19), 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    'create the lower church walls vertex buffers

    Vertices(0).position = vec3(-250, 20, 10)
    Vertices(1).position = vec3(-130, 20, 10)
    Vertices(2).position = vec3(-250, 2, 10)
    Vertices(3).position = vec3(-130, 2, 10)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 10
    Vertices(2).tu = 0
    Vertices(3).tu = 10
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 10
    Vertices(3).tv = 10

    ' Create the vertex buffer.
    Set g_VertexBuffers(20) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(20) Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_VertexBuffers(20), 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(-250, 20, 40)
    Vertices(1).position = vec3(-130, 20, 40)
    Vertices(2).position = vec3(-250, 2, 40)
    Vertices(3).position = vec3(-130, 2, 40)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 10
    Vertices(2).tu = 0
    Vertices(3).tu = 10
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 10
    Vertices(3).tv = 10

    ' Create the vertex buffer.
    Set g_VertexBuffers(21) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(21) Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_VertexBuffers(21), 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(-250, 20, 10)
    Vertices(1).position = vec3(-250, 20, 40)
    Vertices(2).position = vec3(-250, 2, 10)
    Vertices(3).position = vec3(-250, 2, 40)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 10
    Vertices(2).tu = 0
    Vertices(3).tu = 10
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 10
    Vertices(3).tv = 10

    ' Create the vertex buffer.
    Set g_VertexBuffers(22) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(22) Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_VertexBuffers(22), 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(-130, 20, 10)
    Vertices(1).position = vec3(-130, 20, 40)
    Vertices(2).position = vec3(-130, 2, 10)
    Vertices(3).position = vec3(-130, 2, 40)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 10
    Vertices(2).tu = 0
    Vertices(3).tu = 10
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 10
    Vertices(3).tv = 10

    ' Create the vertex buffer.
    Set g_VertexBuffers(23) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(23) Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_VertexBuffers(23), 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    'create the lower church's roof vertex buffers
    
    Vertices(0).position = vec3(-200, 30, 25)
    Vertices(1).position = vec3(-130, 30, 25)
    Vertices(2).position = vec3(-200, 20, 10)
    Vertices(3).position = vec3(-130, 20, 10)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 10
    Vertices(2).tu = 0
    Vertices(3).tu = 10
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 10
    Vertices(3).tv = 10

    ' Create the vertex buffer.
    Set g_VertexBuffers(24) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(24) Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_VertexBuffers(24), 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(-130, 30, 25)
    Vertices(1).position = vec3(-200, 30, 25)
    Vertices(2).position = vec3(-130, 20, 40)
    Vertices(3).position = vec3(-200, 20, 40)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 10
    Vertices(2).tu = 0
    Vertices(3).tu = 10
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 10
    Vertices(3).tv = 10

    ' Create the vertex buffer.
    Set g_VertexBuffers(25) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(25) Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_VertexBuffers(25), 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(-130, 20, 10)
    Vertices(1).position = vec3(-130, 30, 25)
    Vertices(2).position = vec3(-130, 20, 40)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(0).tu = 2
    Vertices(1).tu = 4
    Vertices(2).tu = 6
    Vertices(0).tv = 2
    Vertices(1).tv = 0
    Vertices(2).tv = 0

    ' Create the vertex buffer.
    Set g_VertexBuffers(26) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 3, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(26) Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_VertexBuffers(26), 0, VertexSizeInBytes * 3, 0, Vertices(0)
    
    Vertices(0).position = vec3(-200, 20, 10)
    Vertices(1).position = vec3(-200, 30, 25)
    Vertices(2).position = vec3(-200, 20, 40)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(0).tu = 2
    Vertices(1).tu = 4
    Vertices(2).tu = 6
    Vertices(0).tv = 2
    Vertices(1).tv = 0
    Vertices(2).tv = 0

    ' Create the vertex buffer.
    Set g_VertexBuffers(27) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 3, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(27) Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_VertexBuffers(27), 0, VertexSizeInBytes * 3, 0, Vertices(0)
    
    'create the vertex buffer for the bit of roof below the church tower
    
    Vertices(0).position = vec3(-250, 20, 40)
    Vertices(1).position = vec3(-200, 20, 40)
    Vertices(2).position = vec3(-250, 20, 10)
    Vertices(3).position = vec3(-200, 20, 10)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 10
    Vertices(2).tu = 0
    Vertices(3).tu = 10
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 10
    Vertices(3).tv = 10

    ' Create the vertex buffer.
    Set g_VertexBuffers(28) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(28) Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_VertexBuffers(28), 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    'create the vertex buffers for the four side faces of the tower
    
    Vertices(0).position = vec3(-250, 150, 10)
    Vertices(1).position = vec3(-200, 150, 10)
    Vertices(2).position = vec3(-250, 20, 10)
    Vertices(3).position = vec3(-200, 20, 10)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 10
    Vertices(2).tu = 0
    Vertices(3).tu = 10
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 10
    Vertices(3).tv = 10

    ' Create the vertex buffer.
    Set g_VertexBuffers(29) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(29) Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_VertexBuffers(29), 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(-250, 150, 40)
    Vertices(1).position = vec3(-200, 150, 40)
    Vertices(2).position = vec3(-250, 20, 40)
    Vertices(3).position = vec3(-200, 20, 40)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 10
    Vertices(2).tu = 0
    Vertices(3).tu = 10
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 10
    Vertices(3).tv = 10

    ' Create the vertex buffer.
    Set g_VertexBuffers(30) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(30) Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_VertexBuffers(30), 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(-250, 150, 40)
    Vertices(1).position = vec3(-250, 150, 10)
    Vertices(2).position = vec3(-250, 20, 40)
    Vertices(3).position = vec3(-250, 20, 10)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 10
    Vertices(2).tu = 0
    Vertices(3).tu = 10
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 10
    Vertices(3).tv = 10

    ' Create the vertex buffer.
    Set g_VertexBuffers(31) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(31) Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_VertexBuffers(31), 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(-200.01, 150, 40)
    Vertices(1).position = vec3(-200.01, 150, 10)
    Vertices(2).position = vec3(-200.01, 20, 40)
    Vertices(3).position = vec3(-200.01, 20, 10)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 10
    Vertices(2).tu = 0
    Vertices(3).tu = 10
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 10
    Vertices(3).tv = 10

    ' Create the vertex buffer.
    Set g_VertexBuffers(32) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(32) Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_VertexBuffers(32), 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    'create the vertex buffers for the roof of the church tower
    
    Vertices(0).position = vec3(-250, 170, 25)
    Vertices(1).position = vec3(-200, 170, 25)
    Vertices(2).position = vec3(-250, 150, 10)
    Vertices(3).position = vec3(-200, 150, 10)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 10
    Vertices(2).tu = 0
    Vertices(3).tu = 10
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 10
    Vertices(3).tv = 10

    ' Create the vertex buffer.
    Set g_VertexBuffers(33) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(33) Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_VertexBuffers(33), 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(-200, 170, 25)
    Vertices(1).position = vec3(-250, 170, 25)
    Vertices(2).position = vec3(-200, 150, 40)
    Vertices(3).position = vec3(-250, 150, 40)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 10
    Vertices(2).tu = 0
    Vertices(3).tu = 10
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 10
    Vertices(3).tv = 10

    ' Create the vertex buffer.
    Set g_VertexBuffers(34) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(34) Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_VertexBuffers(34), 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(-200, 150, 10)
    Vertices(1).position = vec3(-200, 170, 25)
    Vertices(2).position = vec3(-200, 150, 40)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(0).tu = 2
    Vertices(1).tu = 4
    Vertices(2).tu = 6
    Vertices(0).tv = 2
    Vertices(1).tv = 0
    Vertices(2).tv = 0

    ' Create the vertex buffer.
    Set g_VertexBuffers(35) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 3, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(35) Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_VertexBuffers(35), 0, VertexSizeInBytes * 3, 0, Vertices(0)
    
    Vertices(0).position = vec3(-250, 150, 10)
    Vertices(1).position = vec3(-250, 170, 25)
    Vertices(2).position = vec3(-250, 150, 40)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(0).tu = 2
    Vertices(1).tu = 4
    Vertices(2).tu = 6
    Vertices(0).tv = 2
    Vertices(1).tv = 0
    Vertices(2).tv = 0

    ' Create the vertex buffer.
    Set g_VertexBuffers(36) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 3, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(36) Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_VertexBuffers(36), 0, VertexSizeInBytes * 3, 0, Vertices(0)
    
    'create the stained glass windows' vertex buffers
    
    Vertices(0).position = vec3(-199.9, 140, 20)
    Vertices(1).position = vec3(-199.9, 140, 30)
    Vertices(2).position = vec3(-199.9, 115, 20)
    Vertices(3).position = vec3(-199.9, 115, 30)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 1
    Vertices(2).tu = 0
    Vertices(3).tu = 1
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 1
    Vertices(3).tv = 1

    ' Create the vertex buffer.
    Set g_VertexBuffers(37) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(37) Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_VertexBuffers(37), 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(-230, 140, 9.9)
    Vertices(1).position = vec3(-220, 140, 9.9)
    Vertices(2).position = vec3(-230, 115, 9.9)
    Vertices(3).position = vec3(-220, 115, 9.9)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 1
    Vertices(2).tu = 0
    Vertices(3).tu = 1
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 1
    Vertices(3).tv = 1

    ' Create the vertex buffer.
    Set g_VertexBuffers(38) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(38) Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_VertexBuffers(38), 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(-250.1, 140, 20)
    Vertices(1).position = vec3(-250.1, 140, 30)
    Vertices(2).position = vec3(-250.1, 115, 20)
    Vertices(3).position = vec3(-250.1, 115, 30)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 1
    Vertices(2).tu = 0
    Vertices(3).tu = 1
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 1
    Vertices(3).tv = 1

    ' Create the vertex buffer.
    Set g_VertexBuffers(39) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(39) Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_VertexBuffers(39), 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    Vertices(0).position = vec3(-230, 140, 40.1)
    Vertices(1).position = vec3(-220, 140, 40.1)
    Vertices(2).position = vec3(-230, 115, 40.1)
    Vertices(3).position = vec3(-220, 115, 40.1)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 1
    Vertices(2).tu = 0
    Vertices(3).tu = 1
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 1
    Vertices(3).tv = 1

    ' Create the vertex buffer.
    Set g_VertexBuffers(40) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(40) Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_VertexBuffers(40), 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    'create the church floor vertex buffer
    
    Vertices(0).position = vec3(-250, 2.1, 40)
    Vertices(1).position = vec3(-130, 2.1, 40)
    Vertices(2).position = vec3(-250, 2.1, 10)
    Vertices(3).position = vec3(-130, 2.1, 10)
    Vertices(0).Color = &HFFFFFF
    Vertices(1).Color = &HFFFFFF
    Vertices(2).Color = &HFFFFFF
    Vertices(3).Color = &HFFFFFF
    Vertices(0).tu = 0
    Vertices(1).tu = 10
    Vertices(2).tu = 0
    Vertices(3).tu = 10
    Vertices(0).tv = 0
    Vertices(1).tv = 0
    Vertices(2).tv = 10
    Vertices(3).tv = 10

    ' Create the vertex buffer.
    Set g_VertexBuffers(41) = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, _
         0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If g_VertexBuffers(41) Is Nothing Then Exit Function
    
    D3DVertexBuffer8SetData g_VertexBuffers(41), 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    InitGeometry = True
End Function

Sub Cleanup()
    'release all components
    Erase g_Textures
    Erase g_VertexBuffers
    Set g_D3DDevice = Nothing
    Set g_D3D = Nothing
    diDEV.Unacquire
    SortPix "delete"
End Sub

Sub Render()

    Dim v As CUSTOMVERTEX
    Dim sizeOfVertex As Long
    
    If g_D3DDevice Is Nothing Then Exit Sub
    g_D3DDevice.SetRenderState D3DRS_FILLMODE, lngFillMode

    ' Clear the backbuffer to a blue color (ARGB = 000000ff)
    ' Clear the z buffer to 1
    g_D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, &HFF&, 1#, 0
    
     
    ' Begin the scene
    g_D3DDevice.BeginScene
    
    g_D3DDevice.SetTextureStageState 0, D3DTSS_ADDRESSU, D3DTADDRESS_WRAP
    g_D3DDevice.SetTextureStageState 0, D3DTSS_ADDRESSV, D3DTADDRESS_WRAP
    
    ' Setup the world, view, and projection matrices
    SetupMatrices
    'setup the lights
    SetupLights
    
    g_D3DDevice.SetVertexShader D3DFVF_CUSTOMVERTEX
    ' Draw the triangles in the vertex buffer
    ' Note we are now using a triangle strip of vertices
    ' instead of a triangle list
    sizeOfVertex = Len(v)

    'draw the contents of our vertex buffers, remembering to change to the correct textures.
    g_D3DDevice.SetTexture 0, g_Textures(1)
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(0), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetTexture 0, g_Textures(2)
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(1), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(2), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(3), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(4), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(5), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2

    g_D3DDevice.SetTexture 0, g_Textures(0)
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(6), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(7), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(8), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(9), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetTexture 0, g_Textures(3)
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(10), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(11), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(12), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 1
    
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(13), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 1
    
    g_D3DDevice.SetTexture 0, g_Textures(4)
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(14), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetTexture 0, g_Textures(5)
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(15), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetTexture 0, g_Textures(6)
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(16), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetTexture 0, Nothing
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(17), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(18), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(19), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetTexture 0, g_Textures(7)
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(20), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(21), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(22), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(23), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetTexture 0, g_Textures(3)
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(24), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(25), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(26), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 1
    
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(27), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 1
    
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(28), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetTexture 0, g_Textures(7)
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(29), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(30), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(31), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(32), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetTexture 0, g_Textures(3)
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(33), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(34), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(35), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 1
    
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(36), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 1
    
    g_D3DDevice.SetTexture 0, g_Textures(8)
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(37), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(38), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(39), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(40), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.SetTexture 0, g_Textures(4)
    g_D3DDevice.SetStreamSource 0, g_VertexBuffers(41), sizeOfVertex
    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
    g_D3DDevice.EndScene
    
    
     
    ' Present the backbuffer contents to the front buffer (screen)
    g_D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
    'update fps
    fps = fps + 1
End Sub

Function vec3(ByVal X As Single, ByVal Y As Single, ByVal z As Single) As D3DVECTOR
    'vector creation helper function
    vec3.X = X
    vec3.Y = Y
    vec3.z = z
End Function

Private Sub Picture1_Click()
    On Error Resume Next
    'bring back the details form
    frmDetails.SetFocus
End Sub

Private Sub tmrFps_Timer()
    'display fps
    frmDetails.lblFps.Caption = fps & " frames per second."
    'reset fps
    fps = 0
End Sub
