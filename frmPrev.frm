VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPrev 
   BackColor       =   &H00000000&
   Caption         =   "Image After Effect - Preview"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6555
   Icon            =   "frmPrev.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   292
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   437
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox LogoBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   960
      Picture         =   "frmPrev.frx":038A
      ScaleHeight     =   465
      ScaleWidth      =   2625
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   2655
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   0
         Top             =   0
      End
      Begin MSComDlg.CommonDialog SaveLoad 
         Left            =   840
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
   End
   Begin VB.PictureBox ToolBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   720
      ScaleHeight     =   465
      ScaleWidth      =   4785
      TabIndex        =   0
      Top             =   480
      Width           =   4815
      Begin VB.CommandButton CmdSave 
         BackColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   4320
         Picture         =   "frmPrev.frx":0B4C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Save as .."
         Top             =   0
         Width           =   450
      End
      Begin VB.CommandButton CmdCopy 
         BackColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   3840
         Picture         =   "frmPrev.frx":1122
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Copy"
         Top             =   0
         Width           =   450
      End
      Begin VB.CommandButton CmdLighting 
         BackColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   3360
         Picture         =   "frmPrev.frx":16F8
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Lighting"
         Top             =   0
         Width           =   450
      End
      Begin VB.CommandButton CmdColor 
         BackColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   2880
         Picture         =   "frmPrev.frx":1BEA
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Colors"
         Top             =   0
         Width           =   450
      End
      Begin VB.CommandButton CmdZoomC 
         BackColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   1320
         Picture         =   "frmPrev.frx":1F2C
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Stretch"
         Top             =   0
         Width           =   450
      End
      Begin VB.CommandButton CmdZoomD 
         BackColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   2280
         Picture         =   "frmPrev.frx":242E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Stretch"
         Top             =   0
         Width           =   450
      End
      Begin VB.CommandButton CmdZoomB 
         BackColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   1800
         Picture         =   "frmPrev.frx":2830
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Stretch"
         Top             =   0
         Width           =   450
      End
      Begin VB.CommandButton CmdZoomA 
         BackColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   840
         Picture         =   "frmPrev.frx":2C32
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Stretch"
         Top             =   0
         Width           =   450
      End
      Begin VB.Image imgZoomIn 
         Height          =   240
         Left            =   120
         Picture         =   "frmPrev.frx":3034
         Tag             =   "Up"
         ToolTipText     =   "Zoom In"
         Top             =   120
         Width           =   240
      End
      Begin VB.Image imgZoomOut 
         Height          =   240
         Left            =   480
         Picture         =   "frmPrev.frx":3285
         Tag             =   "Up"
         ToolTipText     =   "Zoom Out"
         Top             =   120
         Width           =   240
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      FillColor       =   &H0000FF00&
      Height          =   855
      Left            =   2280
      Top             =   2400
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frmPrev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MeshScaleX, MeshScaleZ As Single
Dim MeshPosX, MeshPosZ As Single

Private Sub CmdColor_Click()
EfcColor.Show
End Sub

Private Sub CmdCopy_Click()
Form_Resize
ToolBox.Visible = False
LogoBox.Visible = True
DoEvents
Sleep 250
Timer1_Timer
DoEvents
Sleep 250
Set Me.Picture = CaptureWindow(Me.HWnd, True, 0, 0, Me.ScaleWidth, Me.ScaleHeight)
DoEvents
ToolBox.Visible = True
LogoBox.Visible = False
Clipboard.Clear
Clipboard.SetData Me.Picture
End Sub

Private Sub CmdLighting_Click()
EfcLighting.Show
End Sub

Private Sub CmdSave_Click()
On Error GoTo Err
Form_Resize
Mesh.ScaleMesh 1, 1, 1
MeshScaleX = 1
MeshScaleZ = 1
Mesh.SetPosition -50, 0, -50
MeshPosX = -50
MeshPosZ = -50

If Me.ScaleWidth > Me.ScaleHeight Then
Shape1.Width = Me.ScaleWidth
Shape1.Height = Me.ScaleWidth
Do While Shape1.Height > Me.ScaleHeight
Shape1.Height = Shape1.Height - 1
Shape1.Width = Shape1.Width - 1
Loop
Else
Shape1.Width = Me.ScaleHeight
Shape1.Height = Me.ScaleHeight
End If

Shape1.left = Me.ScaleWidth / 2 - Shape1.Width / 2
Shape1.top = Me.ScaleHeight / 2 - Shape1.Height / 2

ToolBox.Visible = False
LogoBox.Visible = True
LogoBox.left = Shape1.left + Shape1.Width - LogoBox.Width
LogoBox.Width = Me.Width
DoEvents
Sleep 250
Timer1_Timer
DoEvents
Sleep 250
Set Me.Picture = CaptureWindow(Me.HWnd, True, Shape1.left, Shape1.top, Shape1.Width, Shape1.Height)
DoEvents
ToolBox.Visible = True
LogoBox.Visible = False
Clipboard.Clear
Clipboard.SetData Me.Picture

LogoBox.Width = 177

SaveLoad.ShowSave
SavePicture Me.Picture, SaveLoad.FileName + ".dll"
Err:
End Sub

Private Sub CmdZoomA_Click()
Mesh.ScaleMesh MeshScaleX + 0.1, 1, MeshScaleZ
Mesh.SetPosition MeshPosX - 5, 0, MeshPosZ
MeshScaleX = MeshScaleX + 0.1
MeshPosX = MeshPosX - 5
End Sub

Private Sub CmdZoomB_Click()
Mesh.ScaleMesh MeshScaleX, 1, MeshScaleZ + 0.1
Mesh.SetPosition MeshPosX, 0, MeshPosZ - 5
MeshScaleZ = MeshScaleZ + 0.1
MeshPosZ = MeshPosZ - 5
End Sub

Private Sub CmdZoomC_Click()
Mesh.ScaleMesh MeshScaleX - 0.1, 1, MeshScaleZ
Mesh.SetPosition MeshPosX + 5, 0, MeshPosZ
MeshScaleX = MeshScaleX - 0.1
MeshPosX = MeshPosX + 5
End Sub

Private Sub CmdZoomD_Click()
Mesh.ScaleMesh MeshScaleX, 1, MeshScaleZ - 0.1
Mesh.SetPosition MeshPosX, 0, MeshPosZ + 5
MeshScaleZ = MeshScaleZ - 0.1
MeshPosZ = MeshPosZ + 5
End Sub

Private Sub Form_Load()
Tv3D.Init3DWindowedMode Me.HWnd

Set Mesh = Scene.CreateMeshBuilder
Mesh.LoadXFile App.Path + "\Temp\Face.x", False, False

'TextureFactory.LoadTexture App.Path + "\Temp\FaceF.JPG", "Face", , , TV_COLORKEY_WHITE
Scene.LoadTexture App.Path + "\Temp\FaceF.JPG", , , "Face"
Mesh.SetTexture GetTex("Face")

Mesh.SetColor RGBA(1, 1, 1, 1)

Mesh.ScaleMesh 1, 1, 1
MeshScaleX = 1
MeshScaleZ = 1
Mesh.SetPosition -50, 0, -50
MeshPosX = -50
MeshPosZ = -50

Scene.SetCamera 0, -50, 0, 0, 0, 0.01

CmdZoomA_Click

LogoBox.Picture = LoadPicture(App.Path + "\Temp\Logo.dll")
End Sub

Private Sub Form_Resize()
If Me.WindowState = 1 Then Exit Sub

ToolBox.top = Me.ScaleHeight - ToolBox.Height - 7
ToolBox.left = Me.ScaleWidth - ToolBox.Width - 7

LogoBox.top = Me.ScaleHeight - LogoBox.Height - 7
LogoBox.left = Me.ScaleWidth - LogoBox.Width - 7

Tv3D.ResizeDevice
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Mesh = Nothing
Set Scene = Nothing
Set Tv3D = Nothing
Timer1.Enabled = False
End Sub

Private Sub imgZoomIn_Click()
Mesh.ScaleMesh MeshScaleX + 0.1, 0, MeshScaleZ + 0.1
Mesh.SetPosition MeshPosX - 5, 0, MeshPosZ - 5
MeshScaleX = MeshScaleX + 0.1
MeshScaleZ = MeshScaleZ + 0.1
MeshPosX = MeshPosX - 5
MeshPosZ = MeshPosZ - 5
End Sub

Private Sub imgZoomOut_Click()
Mesh.ScaleMesh MeshScaleX - 0.1, 0, MeshScaleZ - 0.1
Mesh.SetPosition MeshPosX + 5, 0, MeshPosZ + 5
MeshScaleX = MeshScaleX - 0.1
MeshScaleZ = MeshScaleZ - 0.1
MeshPosX = MeshPosX + 5
MeshPosZ = MeshPosZ + 5
End Sub

Private Sub Timer1_Timer()
Tv3D.Clear
Mesh.Render

Screen2DImmediate.ACTION_Begin2D
Screen2DImmediate.DRAW_FilledBox 0, 0, Me.ScaleWidth, Me.ScaleHeight, RGBA(AddColorLightingR, AddColorLightingG, AddColorLightingB, AddColorLighting1), RGBA(AddColorLightingR, AddColorLightingG, AddColorLightingB, AddColorLighting2), RGBA(AddColorLightingR, AddColorLightingG, AddColorLightingB, AddColorLighting3), RGBA(AddColorLightingR, AddColorLightingG, AddColorLightingB, AddColorLighting4)
If AddColorLightingA > 0 Then
Screen2DImmediate.DRAW_FilledBox 0, 0, Me.ScaleWidth, Me.ScaleHeight, RGBA(1, 1, 1, AddColorLightingA)
Else
Screen2DImmediate.DRAW_FilledBox 0, 0, Me.ScaleWidth, Me.ScaleHeight, RGBA(0, 0, 0, -AddColorLightingA)
End If
Screen2DImmediate.ACTION_End2D

Tv3D.RenderToScreen
End Sub
