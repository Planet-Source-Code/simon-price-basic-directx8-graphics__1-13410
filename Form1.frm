VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simon's First DirectX8 Program!"
   ClientHeight    =   5760
   ClientLeft      =   36
   ClientTop       =   348
   ClientWidth     =   7680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' FIRST DX8 BY SIMON PRICE 6/12/00

' This is my first DirectX8 program, it's not very good but
' I've only had DirectX8 for 2 hours so far. Most people
' will be beginners at DX8 anyway so this is for the guys like
' me who are still learning. Visit my website www.VBgames.co.uk!

' the main directx object
Private DX As New DirectX8
' the main direct3d object
Private D3D As Direct3D8
' the rendering device
Private D3Ddevice As Direct3DDevice8
' the vertex buffer
Private D3DVB As Direct3DVertexBuffer8

' tells the main loop when to stop
Private EndNow As Boolean

' my vertex type
Private Type COLORVERTEX
    x As Single
    y As Single
    z As Single
    Color As Long
End Type
Private Const D3DFVF_COLORVERTEX = D3DFVF_XYZ Or D3DFVF_DIFFUSE

' color data type
Private Type RGBcolor
    r As Byte
    g As Byte
    b As Byte
End Type
' colors for each corner of the pyrimid
Private VertexRGBColor(0 To 3) As RGBcolor

' the vertices needed to draw the pyrimid
Private Vertex(0 To 11) As COLORVERTEX
' the number of vertices used
Private Const NUMVERTICES = 12
' the size of the pyrimid
Private Const SIZE = 1
' the speed of rotation
Private Const SPEED = 0.05
' the number pi
Private Const PI = 3.1415

Private Sub Form_Load()
On Error Resume Next
    ' show form
    Show
    ' call init function and continue or quit depending on result
    If Init Then
        ' call main program loop
        MainLoop
    Else
        'display error message
        MsgBox "ERROR: occured during ""Init"" function! Closing down...", vbCritical, "FATAL ERROR!"
    End If
    ' end program
    Unload Me
End Sub

Function Init() As Boolean
On Error Resume Next
Dim D3Dpp As D3DPRESENT_PARAMETERS
Dim DisplayMode As D3DDISPLAYMODE
Dim matView As D3DMATRIX
Dim matProj As D3DMATRIX
Dim n As Byte
    ' get a reference to a direct3d object
    Set D3D = DX.Direct3DCreate
    ' get current display mode
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DisplayMode
    ' set d3d parameters
    With D3Dpp
        .BackBufferFormat = DisplayMode.Format
        .Windowed = 1
        .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
    End With
    ' create d3d device
    Set D3Ddevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3Dpp)
    ' turn off lighting
    D3Ddevice.SetRenderState D3DRS_LIGHTING, 0
    ' initialize colors
    For n = 0 To 3
        With VertexRGBColor(n)
            .r = Int(Rnd * 255)
            .g = Int(Rnd * 255)
            .b = Int(Rnd * 255)
        End With
    Next
    CreatePyrimid
    ' create vertex buffer
    Set D3DVB = D3Ddevice.CreateVertexBuffer(Len(Vertex(0)) * NUMVERTICES, 0, D3DFVF_COLORVERTEX, D3DPOOL_DEFAULT)
    ' fill vertex buffer with vertex data
    D3DVertexBuffer8SetData D3DVB, 0, Len(Vertex(0)) * NUMVERTICES, 0, Vertex(0)
    ' set vertex stream source
    D3Ddevice.SetStreamSource 0, D3DVB, Len(Vertex(0))
    ' set vertex shader format
    D3Ddevice.SetVertexShader D3DFVF_COLORVERTEX
    ' set camera position
    D3DXMatrixLookAtLH matView, CreateVector(0, 3, -5), CreateVector(0, 0, 0), CreateVector(0, 1, 0)
    D3Ddevice.SetTransform D3DTS_VIEW, matView
    ' set camera lens
    D3DXMatrixPerspectiveFovLH matProj, PI / 4, 1, 1, 1000
    D3Ddevice.SetTransform D3DTS_PROJECTION, matProj
    ' report success/failure
    If Not Err.Number Then Init = True
End Function

Sub CreatePyrimid()
On Error Resume Next
' temporary vertices
Dim tmpVertex(0 To 3) As COLORVERTEX
' temporary color values
Dim VertexColor(0 To 3) As Long
Dim n As Byte
    ' create vertices in a pyrimid shape
    For n = 0 To 3
        With VertexRGBColor(n)
            .r = .r + Int(Rnd * 11) - 5
            .g = .g + Int(Rnd * 11) - 5
            .b = .b + Int(Rnd * 11) - 5
            VertexColor(n) = RGB(.r, .g, .b)
        End With
    Next
    With tmpVertex(0): .x = 0: .y = SIZE: .z = 0: .Color = VertexColor(0): End With
    With tmpVertex(1): .x = SIZE: .y = -SIZE: .z = -SIZE: .Color = VertexColor(1): End With
    With tmpVertex(2): .x = -SIZE: .y = -SIZE: .z = -SIZE: .Color = VertexColor(2): End With
    With tmpVertex(3): .x = 0: .y = -SIZE: .z = SIZE: .Color = VertexColor(3): End With
    Vertex(0) = tmpVertex(0)
    Vertex(1) = tmpVertex(1)
    Vertex(2) = tmpVertex(2)
    Vertex(3) = tmpVertex(0)
    Vertex(4) = tmpVertex(2)
    Vertex(5) = tmpVertex(3)
    Vertex(6) = tmpVertex(0)
    Vertex(7) = tmpVertex(3)
    Vertex(8) = tmpVertex(1)
    Vertex(9) = tmpVertex(2)
    Vertex(10) = tmpVertex(1)
    Vertex(11) = tmpVertex(3)
End Sub

Function CreateVector(x As Single, y As Single, z As Single) As D3DVECTOR
On Error Resume Next
    ' return a vector
    With CreateVector: .x = x: .y = y: .z = z: End With
End Function

Sub MainLoop()
On Error Resume Next
Dim Angle As Single
Dim matWorld As D3DMATRIX
Dim n As Byte
    Do
        ' clear backbuffer
        D3Ddevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, vbBlack, 0, 0
        ' create a new colorful pyrimid
        CreatePyrimid
        ' fill vertex buffer with vertex data
        D3DVertexBuffer8SetData D3DVB, 0, Len(Vertex(0)) * NUMVERTICES, 0, Vertex(0)
        ' rotate the world
        Angle = Angle + SPEED
        If Angle > 2 * PI Then Angle = Angle - 2 * PI
        D3DXMatrixRotationY matWorld, Angle
        D3Ddevice.SetTransform D3DTS_WORLD, matWorld
        ' begin scene rendering
        D3Ddevice.BeginScene
        ' render the pyrimid shape
        D3Ddevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 4
        ' end scene rendering
        D3Ddevice.EndScene
        ' display scene
        D3Ddevice.Present ByVal 0, ByVal 0, 0, ByVal 0
        ' allow some time for other events
        DoEvents
        ' exit if endnow is true
        If EndNow Then Exit Do
    Loop
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    ' quit is escape key is pressed
    If KeyCode = vbKeyEscape Then EndNow = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    ' tell mainloop to stop
    EndNow = True
    ' set objects to nothing, in reverse order of their creation
    Set D3DVB = Nothing
    Set D3Ddevice = Nothing
    Set D3D = Nothing
    Set DX = Nothing
End Sub
