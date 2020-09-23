VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32" () As Long

Dim DX As DirectX8
Dim D3DX As D3DX8
Dim D3D As Direct3D8
Dim D3DD As Direct3DDevice8
Dim DM As D3DDISPLAYMODE
Dim DPP As D3DPRESENT_PARAMETERS
Dim DI As DirectInput8
Dim DID As DirectInputDevice8
Dim DIS As DIKEYBOARDSTATE

Const PI = 3.14159265358979
Const RAD = PI / 180

Const RED = &HFF0000
Const GREEN = &HFFFF00
Const BLUE = &HFF
Const WHITE = &HFFFFFF

Const FVF_3DUNLIT = D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1

Private Type UdtWorld
    Angle As Single
    Pitch As Single
    CamX As Single
    CamY As Single
    CamZ As Single
    Distance As Single
    Speed As Single
    Jump As Single
    BegY As Single
End Type

Private Type TLV
    POS As D3DVECTOR
    COLOR As Long
    TU As Single
    TV As Single
End Type

Private Type UdtFrame
    LastCheck As Single
    Drawn As Integer
    Rate As Byte
    MF As D3DXFont
    MFD As IFont
    vRECT As RECT
End Type

Private Type UdtBuffer
    VB(99) As Direct3DVertexBuffer8
    WaterVB(39) As Direct3DVertexBuffer8
End Type

Private Type UdtTexture
    Tex(39) As Direct3DTexture8
    WaterTex(9, 9) As Direct3DTexture8
End Type

Private Type UdtWater
    Frame As Integer
    Prims(3) As TLV
    Delay As Integer
    MaxDelay As Integer 'The Higher The Slower
    Set As Integer
End Type

Dim World As UdtWorld
Dim Prims(0 To 3) As TLV
Dim Frame As UdtFrame
Dim Buffer As UdtBuffer
Dim Texture As UdtTexture
Dim Water(9) As UdtWater

Dim MatWorld As D3DMATRIX
Dim MatTemp As D3DMATRIX
Dim MatProj As D3DMATRIX
Dim MatView As D3DMATRIX
Dim MatPitch As D3DMATRIX
Dim MatPos As D3DMATRIX
Dim MatRotation As D3DMATRIX
Dim MatLook As D3DMATRIX

Dim Path As String
Dim RotX, RotY, RotZ As Single
Dim ScrWidth, ScrHeight As Integer
Dim MaxPrims As Integer
Dim MaxWaterPrims As Integer
Dim UseTex(99) As Byte

Dim Moving As Boolean
Dim MovDir As Integer
Dim MovAmount As Single
Dim JUMPING As Boolean
Dim TURNING As Boolean
Dim TurnAmount As Integer

Private Sub Form_Load()
    If Len(App.Path) = 3 Then Path = App.Path Else Path = App.Path & "\"
    StartDX
    Init
    Start
End Sub

Sub StartDX()
    Set DX = New DirectX8
    Set D3DX = New D3DX8
    Set D3D = DX.Direct3DCreate
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DM
    With DPP
        .BackBufferCount = 1
        .BackBufferFormat = DM.Format
        .BackBufferWidth = 640 'DM.Width
        .BackBufferHeight = 480 'DM.Height
        .hDeviceWindow = Me.hWnd
        .AutoDepthStencilFormat = D3DFMT_D16
        .EnableAutoDepthStencil = True
        .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
        .Windowed = 0
    End With
    Set D3DD = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Me.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, DPP)
        
    D3DD.SetRenderState D3DRS_LIGHTING, 1
    D3DD.SetRenderState D3DRS_ZENABLE, 1
    D3DD.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    D3DD.SetRenderState D3DRS_AMBIENT, WHITE
    
    D3DD.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DD.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, True
    
    Set DI = DX.DirectInputCreate
    Set DID = DI.CreateDevice("GUID_SysKeyboard")
    DID.SetCommonDataFormat DIFORMAT_KEYBOARD
    DID.SetCooperativeLevel Me.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    DID.Acquire
    
    DoRotation
    DoLights
    
End Sub

Sub EndDX()
    DID.Unacquire
    Set D3DX = Nothing
    Set D3DD = Nothing
    Set D3D = Nothing
    Set DX = Nothing
End Sub

Sub Init()
    ScrWidth = DPP.BackBufferWidth
    ScrHeight = DPP.BackBufferHeight
    MovDir = 1
    MovAmount = 0
    With Frame
        .LastCheck = GetTickCount
        .Drawn = 0
        .Rate = 0
        With .vRECT
            .Left = 1
            .Top = 1
            .Right = ScrWidth / 4
            .Top = ScrHeight / 8
        End With
        Font.Size = 8
        Set .MFD = Font
        Set .MF = D3DX.CreateFont(D3DD, .MFD.hFont)
    End With
    
    World.Angle = 0
    World.Speed = 2
    World.Distance = -5
    World.CamX = 0
    World.CamY = 15
    World.CamZ = -1
    World.Pitch = 0
    MaxPrims = -1
    MaxWaterPrims = -1
    
    World.BegY = World.CamY
    World.Jump = 0
    
    '0 = Ground Texture: Ground 1
    '1 = Sky Texture: Sky 1
    '2 = Building Texture 1: Wall 1
    '3 = Building Roof: Roof 1
    '4 = Building Floor: Floor 1
    '5 = Painting On Wall: Metroid 'S' Symbol: Painting 1
    '6 = Pyramid Texture: Pic 1
    '7 = Pond Housing Wall: Wall 2
    '8 = Pond Roofing: Roof 2
    For a = 0 To 8
        Set Texture.Tex(a) = D3DX.CreateTextureFromFileEx(D3DD, Path & "Objects\" & a + 1 & ".bmp", D3DX_DEFAULT, D3DX_DEFAULT, 1, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, D3DColorRGBA(255, 255, 255, 255), ByVal 0, ByVal 0)
    Next
    
    Open Path & "World.dat" For Input As #1
        Input #1, temp
    Close #1
    
    For b = 1 To temp
        For a = 0 To 9
            Set Texture.WaterTex(a, 0) = D3DX.CreateTextureFromFileEx(D3DD, Path & "Water\Set " & b & "\" & a + 1 & ".bmp", D3DX_DEFAULT, D3DX_DEFAULT, 1, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
        Next
    Next
    
    'The Ground Primitives
    UseTex(0) = 0
    Prims(0).POS = MV(-1000, 0, 1000)
    Prims(1).POS = MV(1000, 0, 1000)
    Prims(2).POS = MV(-1000, 0, -1000)
    Prims(3).POS = MV(1000, 0, -1000)
    
    MakeTUV 0, 0, 10, 0, 0, 10, 10, 10
    MakeVB 0
    
    'The Sky Primitives
    'Front
    UseTex(1) = 1
    Prims(0).POS = MV(-1000, 1000, 1000)
    Prims(1).POS = MV(1000, 1000, 1000)
    Prims(2).POS = MV(-1000, 0, 1000)
    Prims(3).POS = MV(1000, 0, 1000)
    
    MakeTUV 0, 0, 1, 0, 0, 1, 1, 1
    MakeVB 1
    
    'Back
    UseTex(2) = 1
    Prims(0).POS = MV(-1000, 1000, -1000)
    Prims(1).POS = MV(1000, 1000, -1000)
    Prims(2).POS = MV(-1000, 0, -1000)
    Prims(3).POS = MV(1000, 0, -1000)
    
    MakeVB 2
    
    'Left
    UseTex(3) = 1
    Prims(0).POS = MV(1000, 1000, -1000)
    Prims(1).POS = MV(1000, 1000, 1000)
    Prims(2).POS = MV(1000, 0, -1000)
    Prims(3).POS = MV(1000, 0, 1000)
    
    MakeVB 3
    
    'Right
    UseTex(4) = 1
    Prims(0).POS = MV(-1000, 1000, -1000)
    Prims(1).POS = MV(-1000, 1000, 1000)
    Prims(2).POS = MV(-1000, 0, -1000)
    Prims(3).POS = MV(-1000, 0, 1000)
    
    MakeVB 4
    
    'Top
    UseTex(5) = 1
    Prims(0).POS = MV(-1000, 1000, 1000)
    Prims(1).POS = MV(1000, 1000, 1000)
    Prims(2).POS = MV(-1000, 1000, -1000)
    Prims(3).POS = MV(1000, 1000, -1000)
    
    MakeVB 5
    
    'House Front
    UseTex(6) = 2
    Prims(0).POS = MV(-30, 40, 100)
    Prims(1).POS = MV(30, 40, 100)
    Prims(2).POS = MV(-30, 1, 100)
    Prims(3).POS = MV(30, 1, 100)
    
    MakeTUV 0, 0, 1, 0, 0, 1, 1, 1
    MakeVB 6
    
    'House Back
    UseTex(7) = 2
    Prims(0).POS = MV(-30, 40, 150)
    Prims(1).POS = MV(30, 40, 150)
    Prims(2).POS = MV(-30, 1, 150)
    Prims(3).POS = MV(30, 1, 150)
    
    MakeVB 7
    
    'House Left
    UseTex(8) = 2
    Prims(0).POS = MV(-30, 40, 100)
    Prims(1).POS = MV(-30, 40, 150)
    Prims(2).POS = MV(-30, 1, 100)
    Prims(3).POS = MV(-30, 1, 150)
    
    MakeVB 8
    
    'House Right
    UseTex(9) = 2
    Prims(0).POS = MV(30, 40, 100)
    Prims(1).POS = MV(30, 40, 150)
    Prims(2).POS = MV(30, 1, 100)
    Prims(3).POS = MV(30, 1, 150)
    
    MakeVB 9
    
    'House Roof Left
    UseTex(10) = 3
    Prims(0).POS = MV(-30, 40, 100)
    Prims(1).POS = MV(-30, 40, 150)
    Prims(2).POS = MV(0, 60, 100)
    Prims(3).POS = MV(0, 60, 150)
    
    MakeVB 10
    
    'House Roof Right
    UseTex(11) = 3
    Prims(0).POS = MV(30, 40, 100)
    Prims(1).POS = MV(30, 40, 150)
    Prims(2).POS = MV(0, 60, 100)
    Prims(3).POS = MV(0, 60, 150)
    
    MakeVB 11
    
    'House Floor
    UseTex(12) = 4
    Prims(0).POS = MV(-30, 1, 100)
    Prims(1).POS = MV(30, 1, 100)
    Prims(2).POS = MV(-30, 1, 150)
    Prims(3).POS = MV(30, 1, 150)
    
    MakeVB 12
    
    'House Portrait
    UseTex(13) = 5
    Prims(0).POS = MV(-5, 15, 149)
    Prims(1).POS = MV(-5, 30, 149)
    Prims(2).POS = MV(5, 15, 149)
    Prims(3).POS = MV(5, 30, 149)
    
    MakeTUV 0, 1, 0, 0, 1, 1, 1, 0
    MakeVB 13
    
    'Pyramid Back
    UseTex(14) = 6
    Prims(0).POS = MV(-100, 0, -300)
    Prims(1).POS = MV(0, 200, -200)
    Prims(2).POS = MV(100, 0, -300)
    Prims(3).POS = MV(100, 0, -300)
    
    'MakeTUV 0, 1, 0.5, 0, 1, 1, 1, 1
    MakeTUV 0, 5, 2.5, 0, 5, 5, 5, 5
    MakeVB 14
    
    'Pyramid Front
    UseTex(15) = 6
    Prims(0).POS = MV(-100, 0, -100)
    Prims(1).POS = MV(0, 200, -200)
    Prims(2).POS = MV(100, 0, -100)
    Prims(3).POS = MV(100, 0, -100)
    
    MakeVB 15
    
    'Pyramid Left
    UseTex(16) = 6
    
    Prims(0).POS = MV(-100, 0, -100)
    Prims(1).POS = MV(0, 200, -200)
    Prims(2).POS = MV(-100, 0, -300)
    Prims(3).POS = MV(-100, 0, -300)
    
    MakeVB 16
    
    'Pyramid Right
    UseTex(17) = 6
    
    Prims(0).POS = MV(100, 0, -100)
    Prims(1).POS = MV(0, 200, -200)
    Prims(2).POS = MV(100, 0, -300)
    Prims(3).POS = MV(100, 0, -300)
    
    MakeVB 17
    
    'Pyramid Floor
    UseTex(18) = 6
    
    Prims(0).POS = MV(-100, 1, -100)
    Prims(1).POS = MV(100, 1, -100)
    Prims(2).POS = MV(-100, 1, -300)
    Prims(3).POS = MV(100, 1, -300)
    
    MakeTUV 0, 0, 5, 0, 0, 5, 5, 5
    MakeVB 18
    
    'Pond
    UseTex(19) = 6
    
    Prims(0).POS = MV(-300, 0, -75)
    Prims(1).POS = MV(-300, 5, -75)
    Prims(2).POS = MV(-150, 0, -75)
    Prims(3).POS = MV(-150, 5, -75)
    
    MakeTUV 0, 1, 0, 0, 1, 1, 1, 0
    MakeVB 19
    
    UseTex(20) = 6
    
    Prims(0).POS = MV(-290, 0, -65)
    Prims(1).POS = MV(-290, 5, -65)
    Prims(2).POS = MV(-160, 0, -65)
    Prims(3).POS = MV(-160, 5, -65)
    
    MakeTUV 0, 1, 0, 0, 1, 1, 1, 0
    MakeVB 20
    
    UseTex(21) = 6
    
    Prims(0).POS = MV(-300, 0, 75)
    Prims(1).POS = MV(-300, 5, 75)
    Prims(2).POS = MV(-150, 0, 75)
    Prims(3).POS = MV(-150, 5, 75)

    MakeTUV 0, 1, 0, 0, 1, 1, 1, 0
    MakeVB 21
    
    UseTex(22) = 6
    
    Prims(0).POS = MV(-290, 0, 65)
    Prims(1).POS = MV(-290, 5, 65)
    Prims(2).POS = MV(-160, 0, 65)
    Prims(3).POS = MV(-160, 5, 65)
    
    MakeTUV 0, 1, 0, 0, 1, 1, 1, 0
    MakeVB 22
    
    UseTex(23) = 6
    
    Prims(0).POS = MV(-300, 0, -75)
    Prims(1).POS = MV(-300, 5, -75)
    Prims(2).POS = MV(-300, 0, 75)
    Prims(3).POS = MV(-300, 5, 75)
    
    MakeTUV 0, 1, 0, 0, 1, 1, 1, 0
    MakeVB 23
    
    UseTex(24) = 6
    
    Prims(0).POS = MV(-290, 0, -65)
    Prims(1).POS = MV(-290, 5, -65)
    Prims(2).POS = MV(-290, 0, 65)
    Prims(3).POS = MV(-290, 5, 65)
    
    MakeTUV 0, 1, 0, 0, 1, 1, 1, 0
    MakeVB 24
    
    UseTex(25) = 6

    Prims(0).POS = MV(-150, 0, -75)
    Prims(1).POS = MV(-150, 5, -75)
    Prims(2).POS = MV(-150, 0, 75)
    Prims(3).POS = MV(-150, 5, 75)
    
    MakeTUV 0, 1, 0, 0, 1, 1, 1, 0
    MakeVB 25
    
    UseTex(26) = 6
    
    Prims(0).POS = MV(-160, 0, -65)
    Prims(1).POS = MV(-160, 5, -65)
    Prims(2).POS = MV(-160, 0, 65)
    Prims(3).POS = MV(-160, 5, 65)
    
    MakeTUV 0, 1, 0, 0, 1, 1, 1, 0
    MakeVB 26
    
    UseTex(27) = 6
    
    Prims(0).POS = MV(-300, 5, -75)
    Prims(1).POS = MV(-300, 5, -65)
    Prims(2).POS = MV(-150, 5, -75)
    Prims(3).POS = MV(-150, 5, -65)
    
    MakeTUV 0, 1, 0, 0, 1, 1, 1, 0
    MakeVB 27
    
    UseTex(28) = 6
    
    Prims(0).POS = MV(-300, 5, 75)
    Prims(1).POS = MV(-300, 5, 65)
    Prims(2).POS = MV(-150, 5, 75)
    Prims(3).POS = MV(-150, 5, 65)
    
    MakeTUV 0, 1, 0, 0, 1, 1, 1, 0
    MakeVB 28
    
    UseTex(29) = 6
    
    Prims(0).POS = MV(-300, 5, -65)
    Prims(1).POS = MV(-290, 5, -65)
    Prims(2).POS = MV(-300, 5, 65)
    Prims(3).POS = MV(-290, 5, 65)
    
    MakeTUV 0, 1, 0, 0, 1, 1, 1, 0
    MakeVB 29
    
    UseTex(30) = 6
    
    Prims(0).POS = MV(-160, 5, -65)
    Prims(1).POS = MV(-150, 5, -65)
    Prims(2).POS = MV(-160, 5, 65)
    Prims(3).POS = MV(-150, 5, 65)
    
    MakeTUV 0, 1, 0, 0, 1, 1, 1, 0
    MakeVB 30
       
    With Water(MaxWaterPrims + 1)
        .MaxDelay = 3
        .Delay = .MaxDelay
        .Frame = 0
        .Set = 0
        
        .Prims(0).POS = MV(-290, 2, -65)
        .Prims(1).POS = MV(-290, 2, 65)
        .Prims(2).POS = MV(-160, 2, -65)
        .Prims(3).POS = MV(-160, 2, 65)
        
        MakeWaterTUV MaxWaterPrims + 1, 0, 1, 0, 0, 1, 1, 1, 0
        MakeWaterVB MaxWaterPrims + 1
    End With
    
    'Pond Housing
    UseTex(31) = 7
    
    Prims(0).POS = MV(-310, 0, -85)
    Prims(1).POS = MV(-310, 50, -85)
    Prims(2).POS = MV(-140, 0, -85)
    Prims(3).POS = MV(-140, 50, -85)
    
    MakeTUV 0, 2, 0, 0, 2, 2, 2, 0
    MakeVB 31
    
    UseTex(32) = 7
    
    Prims(0).POS = MV(-310, 0, 85)
    Prims(1).POS = MV(-310, 50, 85)
    Prims(2).POS = MV(-140, 0, 85)
    Prims(3).POS = MV(-140, 50, 85)
    
    MakeTUV 0, 2, 0, 0, 2, 2, 2, 0
    MakeVB 32
    
    UseTex(33) = 7
    
    Prims(0).POS = MV(-310, 0, -85)
    Prims(1).POS = MV(-310, 50, -85)
    Prims(2).POS = MV(-310, 0, 0)
    Prims(3).POS = MV(-310, 60, 0)
    
    MakeTUV 0, 2, 0, 0, 2, 2, 2, 0
    MakeVB 33
    
    UseTex(34) = 7
    
    Prims(0).POS = MV(-310, 0, 0)
    Prims(1).POS = MV(-310, 60, 0)
    Prims(2).POS = MV(-310, 0, 85)
    Prims(3).POS = MV(-310, 50, 85)
    
    MakeTUV 0, 2, 0, 0, 2, 2, 2, 0
    MakeVB 34
    
    UseTex(35) = 7
    
    Prims(0).POS = MV(-140, 0, -85)
    Prims(1).POS = MV(-140, 50, -85)
    Prims(2).POS = MV(-140, 0, -20)
    Prims(3).POS = MV(-140, 57.7, -20)
    
    MakeTUV 0, 2, 0, 0, 2, 2, 2, 0
    MakeVB 35
    
    UseTex(36) = 7
    
    Prims(0).POS = MV(-140, 0, 20)
    Prims(1).POS = MV(-140, 57.7, 20)
    Prims(2).POS = MV(-140, 0, 85)
    Prims(3).POS = MV(-140, 50, 85)
    
    MakeTUV 0, 2, 0, 0, 2, 2, 2, 0
    MakeVB 36
    
    UseTex(37) = 8
    
    Prims(0).POS = MV(-310, 50, -85)
    Prims(1).POS = MV(-310, 60, 0)
    Prims(2).POS = MV(-140, 50, -85)
    Prims(3).POS = MV(-140, 60, 0)
    
    MakeTUV 0, 2, 0, 0, 2, 2, 2, 0
    MakeVB 37
    
    UseTex(38) = 8
    
    Prims(0).POS = MV(-310, 60, 0)
    Prims(1).POS = MV(-140, 60, 0)
    Prims(2).POS = MV(-310, 50, 85)
    Prims(3).POS = MV(-140, 50, 85)
    
    MakeTUV 0, 2, 0, 0, 2, 2, 2, 0
    MakeVB 38
    
End Sub

Sub Start()
    Do
        CheckKeys
        Render
        DoEvents
    Loop
End Sub

Sub CheckKeys()
    DID.GetDeviceStateKeyboard DIS
    
    If DIS.Key(DIK_ESCAPE) Then
        EndDX
        End
    End If
    
    If DIS.Key(DIK_LALT) Or DIS.Key(DIK_RALT) Then
        If DIS.Key(DIK_LEFT) Then
            World.CamX = World.CamX - Cos((World.Angle) * RAD) * World.Speed
            World.CamZ = World.CamZ - Sin((World.Angle) * RAD) * World.Speed
        ElseIf DIS.Key(DIK_RIGHT) Then
            World.CamX = World.CamX + Cos((World.Angle) * RAD) * World.Speed
            World.CamZ = World.CamZ + Sin((World.Angle) * RAD) * World.Speed
        End If
    Else
        If DIS.Key(DIK_LEFT) Then
            If Not World.Angle >= 360 Then World.Angle = World.Angle + World.Speed Else World.Angle = World.Speed
        ElseIf DIS.Key(DIK_RIGHT) Then
            If Not World.Angle <= 0 Then World.Angle = World.Angle - World.Speed Else World.Angle = 360 - World.Speed
        End If
    End If
    
    If DIS.Key(DIK_UP) Then
        Moving = True
        If MovDir = 1 Then If Not MovAmount >= 1.5 Then MovAmount = MovAmount + 0.15 Else MovDir = 2
        If MovDir = 2 Then If Not MovAmount <= 0 Then MovAmount = MovAmount - 0.15 Else MovDir = 1
        World.CamX = World.CamX - Sin(World.Angle * RAD) * World.Speed
        World.CamZ = World.CamZ + Cos(World.Angle * RAD) * World.Speed
    ElseIf DIS.Key(DIK_DOWN) Then
        If MovDir = 1 Then If Not MovAmount >= 1.5 Then MovAmount = MovAmount + 0.15 Else MovDir = 2
        If MovDir = 2 Then If Not MovAmount <= 0 Then MovAmount = MovAmount - 0.15 Else MovDir = 1
        World.CamX = World.CamX + Sin(World.Angle * RAD) * World.Speed
        World.CamZ = World.CamZ - Cos(World.Angle * RAD) * World.Speed
    Else
        MovDir = 1
        If Not MovAmount <= 0 Then MovAmount = MovAmount - 0.15
    End If
    
    If DIS.Key(DIK_A) Then
        If Not World.Pitch >= 1 Then World.Pitch = World.Pitch + 0.05
    ElseIf DIS.Key(DIK_Z) Then
        If Not World.Pitch <= -1 Then World.Pitch = World.Pitch - 0.05
    End If
    
    If DIS.Key(DIK_SPACE) Then
        If Not JUMPING = True And World.Jump <= 0 Then JUMPING = True
    End If
    
    If DIS.Key(DIK_BACKSPACE) Then
        If Not TURNING = True Then TURNING = True
    End If
    
    If JUMPING = True Then If Not World.Jump >= 30 Then World.Jump = World.Jump + World.Speed Else JUMPING = False
    If JUMPING = False Then If Not World.Jump <= 0 Then World.Jump = World.Jump - World.Speed Else World.Jump = 0
    
    If TURNING = True Then
        If Not TurnAmount >= 180 Then
            TurnAmount = TurnAmount + World.Speed * 4
            If Not World.Angle >= 360 Then World.Angle = World.Angle + World.Speed * 4 Else World.Angle = 0
        Else
            TurnAmount = 0
            TURNING = False
        End If
    End If
End Sub

Sub Render()
    D3DD.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, D3DColorRGBA(0, 0, 0, 255), 1#, 0
    D3DD.BeginScene
    
    D3DD.SetVertexShader FVF_3DUNLIT
    
    'DoLights
    DoRotation
    
    For a = 0 To MaxPrims
        D3DD.SetTexture 0, Texture.Tex(UseTex(a))
        D3DD.SetStreamSource 0, Buffer.VB(a), Len(Prims(0))
        D3DD.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    Next
    
    For a = 0 To MaxWaterPrims
        D3DD.SetTexture 0, Texture.WaterTex(Water(a).Frame, Water(a).Set)
        D3DD.SetStreamSource 0, Buffer.WaterVB(a), Len(Water(a).Prims(0))
        D3DD.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
        
        If Not Water(a).Delay <= 0 Then
            Water(a).Delay = Water(a).Delay - 1
        Else
            Water(a).Delay = Water(a).MaxDelay
            If Not Water(a).Frame >= 9 Then Water(a).Frame = Water(a).Frame + 1 Else Water(a).Frame = 0
        End If
    Next
    
    With Frame
        D3DX.DrawText .MF, D3DColorRGBA(2550, 0, 0, 255), "FPS: " & GetFrameRate, .vRECT, DT_LEFT Or DT_TOP
    End With

    D3DD.EndScene
    D3DD.Present ByVal 0, ByVal 0, 0, ByVal 0
End Sub

Private Function MV(ByVal X As Single, ByVal Y As Single, ByVal Z As Single) As D3DVECTOR
    With MV
        .X = X
        .Y = Y
        .Z = Z
    End With
End Function

Sub Rotate(ByVal X As Single, ByVal Y As Single, ByVal Z As Single)
    RotZ = Z * RAD
    RotX = X * Cos(RotZ) - Y * Sin(RotZ)
    RotY = Y * Cos(RotZ) - X * Sin(RotZ)
End Sub

Private Function GetFrameRate() As Long
    With Frame
        If GetTickCount - .LastCheck >= 1000 Then
            .Rate = .Drawn
            .Drawn = 0
            .LastCheck = GetTickCount
        End If
        .Drawn = .Drawn + 1
        GetFrameRate = .Rate
    End With
End Function

Sub DoLights()
    Dim MTRL As D3DMATERIAL8
    Dim COL As D3DCOLORVALUE
    Dim LIGHT As D3DLIGHT8
    
    With COL
        .a = 1
        .b = 1
        .g = 1
        .r = 1
    End With
    
    MTRL.diffuse = COL
    MTRL.Ambient = COL
    D3DD.SetMaterial MTRL
    
    With LIGHT
        .Type = D3DLIGHT_POINT
        .diffuse.r = 255
        .diffuse.g = 255
        .diffuse.b = 255
        .Position = MV(0, -1, 0)
        .Direction = MV(0, 1, 0)
        .Range = 1000
        .Attenuation0 = 0.5
        .Attenuation1 = 0.5
        .Attenuation2 = 0.5
    End With
    
    D3DD.SetLight 0, LIGHT
    D3DD.LightEnable 0, True
End Sub

Sub DoRotation()
    D3DXMatrixIdentity MatView
    D3DXMatrixIdentity MatPos
    D3DXMatrixIdentity MatRotation
    D3DXMatrixIdentity MatLook
    D3DXMatrixIdentity MatWorld
        
    World.CamY = World.BegY + MovAmount + World.Jump
        
    D3DXMatrixRotationY MatRotation, World.Angle * RAD
    D3DXMatrixRotationX MatPitch, World.Pitch
    D3DXMatrixMultiply MatLook, MatRotation, MatPitch
    
    D3DXMatrixTranslation MatPos, -World.CamX, -World.CamY, -World.CamZ
    D3DXMatrixMultiply MatView, MatPos, MatLook
    D3DD.SetTransform D3DTS_VIEW, MatView
    
    D3DD.SetTransform D3DTS_WORLD, MatWorld
    D3DXMatrixPerspectiveFovLH MatProj, PI / 3, 1, 1, 10000
    D3DD.SetTransform D3DTS_PROJECTION, MatProj
End Sub

Sub MakeTUV(ByVal TU1 As Single, ByVal TV1 As Single, ByVal TU2 As Single, ByVal TV2 As Single, ByVal TU3 As Single, ByVal TV3 As Single, ByVal TU4 As Single, ByVal TV4 As Single)
    Prims(0).TU = TU1: Prims(0).TV = TV1
    Prims(1).TU = TU2: Prims(1).TV = TV2
    Prims(2).TU = TU3: Prims(2).TV = TV3
    Prims(3).TU = TU4: Prims(3).TV = TV4
End Sub

Sub MakeVB(ByVal Index As Integer)
    Set Buffer.VB(Index) = D3DD.CreateVertexBuffer(Len(Prims(0)) * 4, 0, FVF_3DUNLIT, D3DPOOL_DEFAULT)
    D3DVertexBuffer8SetData Buffer.VB(Index), 0, Len(Prims(0)) * 4, 0, Prims(0)
    MaxPrims = MaxPrims + 1
End Sub

Sub MakeWaterTUV(ByVal Index As Integer, ByVal TU1 As Single, ByVal TV1 As Single, ByVal TU2 As Single, ByVal TV2 As Single, ByVal TU3 As Single, ByVal TV3 As Single, ByVal TU4 As Single, ByVal TV4 As Single)
    With Water(Index)
        .Prims(0).TU = TU1: .Prims(0).TV = TV1
        .Prims(1).TU = TU2: .Prims(1).TV = TV2
        .Prims(2).TU = TU3: .Prims(2).TV = TV3
        .Prims(3).TU = TU4: .Prims(3).TV = TV4
    End With
End Sub

Sub MakeWaterVB(ByVal Index As Integer)
    Set Buffer.WaterVB(Index) = D3DD.CreateVertexBuffer(Len(Water(Index).Prims(0)) * 4, 0, FVF_3DUNLIT, D3DPOOL_DEFAULT)
    D3DVertexBuffer8SetData Buffer.WaterVB(Index), 0, Len(Water(Index).Prims(0)) * 4, 0, Water(Index).Prims(0)
    MaxWaterPrims = MaxWaterPrims + 1
End Sub
