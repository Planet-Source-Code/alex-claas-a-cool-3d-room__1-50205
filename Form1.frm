VERSION 5.00
Object = "{08216199-47EA-11D3-9479-00AA006C473C}#2.1#0"; "RMCONTROL.OCX"
Begin VB.Form Form1 
   Caption         =   "3D Room"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7800
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin RMControl7.RMCanvas RMCanvas1 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   10610
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3000
      Top             =   1800
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FloorMesh As Direct3DRMMeshBuilder3 'Floor Mesh
Dim WallsMeshes As Direct3DRMMeshBuilder3 ' Walls Mesh
Dim TableMesh As Direct3DRMMeshBuilder3 ' Table 1 mesh
Dim ChairMesh As Direct3DRMMeshBuilder3 'Chair mesh
Dim CompMesh As Direct3DRMMeshBuilder3 ' Computer mesh
Dim T2mesh As Direct3DRMMeshBuilder3 ' Table 2 mesh

Dim Floor As Direct3DRMFrame3 ' Floor Frame
Dim Walls As Direct3DRMFrame3 ' Walls Frame
Dim Table As Direct3DRMFrame3 'Table 1 Frame
Dim Chair As Direct3DRMFrame3 ' Chair Frame
Dim Comp As Direct3DRMFrame3 ' Computer Frame
Dim T2 As Direct3DRMFrame3 ' Table 2 Frame

Dim Movspeed ' Moving Speed
Const RotSpeed = 2 ' Rotation Speed
Const PI = 3.14324545324366 'PI Constant
Private Sub Form_Load()
ChDir App.Path ' Makes the code direct all paths to the path of the application ( what folder it is saved in)
With RMCanvas1 'Makes all following statements do with RMCanvas1
    .StartWindowed 'The start windowed object (I'm not really sure what this does but I just use it)
    .SceneFrame.SetSceneBackgroundRGB 0, 0, 0 'Set the color of the background
    .Viewport.SetBack (9000) 'Sets how far you can see
    .CameraFrame.SetPosition Nothing, 0, 3, 0 'The position of the camera
    
    Set FloorMesh = .D3DRM.CreateMeshBuilder() 'Creates the mesh builder for the floor
    Set WallsMeshes = .D3DRM.CreateMeshBuilder() 'Create the mesh builder for the walls
    Set TableMesh = .D3DRM.CreateMeshBuilder() ' Creates the mesh builder for the table
    Set ChairMesh = .D3DRM.CreateMeshBuilder() ' Creates the mesh builder for the chair
    Set CompMesh = .D3DRM.CreateMeshBuilder() ' Creates the mesh builder for the computer
    Set T2mesh = .D3DRM.CreateMeshBuilder() ' Creates the mesh builder for the 2nd table
    
    Set Floor = .D3DRM.CreateFrame(.SceneFrame) ' Creates the frame for the floor
    Set Walls = .D3DRM.CreateFrame(.SceneFrame) ' Creates the frame for the walls
    Set Table = .D3DRM.CreateFrame(.SceneFrame) ' Creates the frame for the table
    Set Chair = .D3DRM.CreateFrame(.SceneFrame) ' Creates the frame for the chair
    Set Comp = .D3DRM.CreateFrame(.SceneFrame) ' Creates the frame for the computer
    Set T2 = .D3DRM.CreateFrame(.SceneFrame) ' Creates the frame for the second table
End With

FloorMesh.LoadFromFile "Floor.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing 'Loads the floor .x file
Floor.AddVisual FloorMesh 'Makes the scene add the floor to it
Floor.SetPosition Nothing, 2, 0, 0 ' Sets the position of the floor
Floor.SetMaterialMode D3DRMMATERIAL_FROMFRAME 'Makes it so you choose the color
Floor.SetColorRGB 1, 0, 0 'Makes the color RGB (Red, Blue, Green) red

WallsMeshes.LoadFromFile "Walls.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
Walls.AddVisual WallsMeshes
Walls.SetPosition Nothing, 2, 0, 14.2
Walls.SetMaterialMode D3DRMMATERIAL_FROMFRAME
Walls.SetColorRGB 0, 0, 1

TableMesh.LoadFromFile "Table.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
Table.AddVisual TableMesh
Table.SetPosition Nothing, 2, 0, 11
Table.SetMaterialMode D3DRMMATERIAL_FROMFRAME
Table.SetColorRGB 0, 1, 0

ChairMesh.LoadFromFile "Chair.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
Chair.AddVisual ChairMesh
Chair.SetPosition Nothing, -1, 0, 10
Chair.SetMaterialMode D3DRMMATERIAL_FROMMESH ' Makes the program get the texture from the mesh

CompMesh.LoadFromFile "Computer.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
Comp.AddVisual CompMesh
Comp.SetPosition Nothing, 3, 2.5, 10
Comp.SetMaterialMode D3DRMMATERIAL_FROMMESH

T2mesh.LoadFromFile "Table.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
T2.AddVisual T2mesh
T2.SetPosition Nothing, 4, 0, 11
T2.SetMaterialMode D3DRMMATERIAL_FROMFRAME
T2.SetColorRGB 0, 1, 0

Movspeed = 1 'Sets the movespeed to 1
RMCanvas1.Update 'Updates the RMcanvas
End Sub
 

Private Sub RMCanvas1_KeyPress(KeyAscii As Integer)
Dim CV As D3DVECTOR 'camera vector (position)
RMCanvas1.CameraFrame.GetPosition Nothing, CV 'get the camera's position
   
'axal movement
If LCase(Chr(KeyAscii)) = "a" Then RMCanvas1.CameraFrame.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, -RotSpeed * PI / 180 'Rotates camera frame
If LCase(Chr(KeyAscii)) = "d" Then RMCanvas1.CameraFrame.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, RotSpeed * PI / 180 ' Rotates camera frame
If LCase(Chr(KeyAscii)) = "w" Then Call RMCanvas1.CameraFrame.SetPosition(RMCanvas1.CameraFrame, 0, 0, Movspeed) ' Moves camera frame forward
If LCase(Chr(KeyAscii)) = "s" Then Call RMCanvas1.CameraFrame.SetPosition(RMCanvas1.CameraFrame, 0, 0, -Movspeed) ' Moves camera frame backward
End Sub

Private Sub Timer1_Timer()
RMCanvas1.Update ' Keeps the RMcanvas updated
End Sub
Private Sub FOrm_Resize()
'Makes it so the form1 width & height is the same as the RMcanvas
RMCanvas1.Width = Form1.ScaleWidth - RMCanvas1.Left
RMCanvas1.Height = Form1.ScaleHeight - RMCanvas1.Top
End Sub
