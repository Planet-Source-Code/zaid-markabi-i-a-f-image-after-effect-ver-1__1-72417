VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image After Effect"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   13575
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":038A
   ScaleHeight     =   5355
   ScaleWidth      =   13575
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   8040
      Picture         =   "frmMain.frx":03D0
      ScaleHeight     =   375
      ScaleWidth      =   1215
      TabIndex        =   21
      Top             =   120
      Width           =   1215
      Begin VB.Label CdmRestore 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Restore"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   0
         TabIndex        =   22
         Top             =   60
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   12240
      Picture         =   "frmMain.frx":17AA
      ScaleHeight     =   375
      ScaleWidth      =   1215
      TabIndex        =   19
      Top             =   120
      Width           =   1215
      Begin VB.Label CmdAbout 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   0
         TabIndex        =   20
         Top             =   60
         Width           =   1095
      End
   End
   Begin VB.PictureBox PictureA 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      DrawStyle       =   2  'Dot
      ForeColor       =   &H0000FF00&
      Height          =   4635
      Left            =   2280
      Picture         =   "frmMain.frx":2B84
      ScaleHeight     =   1
      ScaleMode       =   0  'User
      ScaleWidth      =   1
      TabIndex        =   12
      Top             =   600
      Width           =   5535
      Begin VB.Timer ScrollPic 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   600
         Top             =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         Caption         =   "Before"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   555
      End
      Begin VB.Shape DotPos 
         BorderColor     =   &H000000FF&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape DotPos 
         BorderColor     =   &H000000FF&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   5400
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape DotPos 
         BorderColor     =   &H000000FF&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   0
         Shape           =   3  'Circle
         Top             =   4440
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape DotPos 
         BorderColor     =   &H000000FF&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   3
         Left            =   5280
         Shape           =   3  'Circle
         Top             =   4320
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape DotPos 
         BorderColor     =   &H000000FF&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   4
         Left            =   2160
         Shape           =   3  'Circle
         Top             =   2625
         Width           =   135
      End
      Begin VB.Shape DotPos 
         BorderColor     =   &H000000FF&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   5
         Left            =   2918
         Shape           =   3  'Circle
         Top             =   2625
         Width           =   135
      End
      Begin VB.Shape DotPos 
         BorderColor     =   &H000000FF&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   6
         Left            =   2560
         Shape           =   3  'Circle
         Top             =   2460
         Width           =   135
      End
      Begin VB.Shape DotPos 
         BorderColor     =   &H000000FF&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   7
         Left            =   2560
         Shape           =   3  'Circle
         Top             =   2809
         Width           =   135
      End
   End
   Begin VB.CheckBox ChkLines 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Lines"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6960
      TabIndex        =   18
      Top             =   360
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.CheckBox ChkBoth 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Both"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6960
      TabIndex        =   17
      Top             =   120
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3600
      Picture         =   "frmMain.frx":2BF6
      ScaleHeight     =   375
      ScaleWidth      =   2175
      TabIndex        =   15
      Top             =   120
      Width           =   2175
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Load New Image"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Top             =   60
         Width           =   2055
      End
   End
   Begin VB.PictureBox PictureB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4635
      Left            =   7920
      Picture         =   "frmMain.frx":5050
      ScaleHeight     =   1
      ScaleMode       =   0  'User
      ScaleWidth      =   1
      TabIndex        =   11
      Top             =   600
      Width           =   5535
      Begin VB.Timer ScrollPicB 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   240
         Top             =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         Caption         =   "After"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   405
      End
      Begin VB.Shape DotPosN 
         BorderColor     =   &H000000FF&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   7
         Left            =   2587
         Shape           =   3  'Circle
         Top             =   2640
         Width           =   135
      End
      Begin VB.Shape DotPosN 
         BorderColor     =   &H000000FF&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   6
         Left            =   2580
         Shape           =   3  'Circle
         Top             =   2520
         Width           =   135
      End
      Begin VB.Shape DotPosN 
         BorderColor     =   &H000000FF&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   5
         Left            =   2880
         Shape           =   3  'Circle
         Top             =   2520
         Width           =   135
      End
      Begin VB.Shape DotPosN 
         BorderColor     =   &H000000FF&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   4
         Left            =   2160
         Shape           =   3  'Circle
         Top             =   2520
         Width           =   135
      End
      Begin VB.Shape DotPosN 
         BorderColor     =   &H000000FF&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   3
         Left            =   5400
         Shape           =   3  'Circle
         Top             =   4440
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape DotPosN 
         BorderColor     =   &H000000FF&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   0
         Shape           =   3  'Circle
         Top             =   4440
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape DotPosN 
         BorderColor     =   &H000000FF&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   5400
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape DotPosN 
         BorderColor     =   &H000000FF&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   135
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   2280
      Picture         =   "frmMain.frx":50F6
      ScaleHeight     =   375
      ScaleWidth      =   1215
      TabIndex        =   9
      Top             =   120
      Width           =   1215
      Begin VB.Label PreviewCmd 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Preview"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   60
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mesh"
      Enabled         =   0   'False
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      Begin MSComDlg.CommonDialog SaveLoad 
         Left            =   1320
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.ListBox lstMat 
         Appearance      =   0  'Flat
         Height          =   810
         Left            =   1440
         TabIndex        =   5
         Top             =   4080
         Width           =   495
      End
      Begin VB.ListBox lstTexs 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   3360
         Width           =   1815
      End
      Begin VB.ListBox lstTex 
         Appearance      =   0  'Flat
         Height          =   810
         Left            =   120
         TabIndex        =   3
         Top             =   4080
         Width           =   1335
      End
      Begin VB.ListBox lstFaces 
         Appearance      =   0  'Flat
         Height          =   1200
         Left            =   120
         TabIndex        =   2
         Top             =   1800
         Width           =   1815
      End
      Begin VB.ListBox lstPoints 
         Appearance      =   0  'Flat
         Height          =   1005
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Texture Coords"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   3120
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Faces"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   510
      End
      Begin VB.Label Label0 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Points"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   525
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuLoadImage 
         Caption         =   "Load New Image"
      End
      Begin VB.Menu mnuLoadProject 
         Caption         =   "Load Project"
      End
      Begin VB.Menu mnuSaveProject 
         Caption         =   "Save Project"
      End
      Begin VB.Menu mnuSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuProject 
      Caption         =   "Project"
      Begin VB.Menu mnuPreview 
         Caption         =   "Preview"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Written By [ Zaid Markabi ]
' Em@il   : ZaidMarkabi@yahoo.com
' Website : http://www.yazanmarkabi.webs.com/
'
' _Caution__________________________________
'| THIS APPLICATION IS FREEWARE. ANY USE IN |
'| COMMERCIAL APPLICATIONS WITHOUT WRITTEN  |
'| PERMISSION BY THE AUTHOR IS PROHIBITED.  |
'|__________________________________________|
'
' Image After Effect (IAE) is free trial software,
' let you to add effects and animations on your photos,
' you can change mouth, eyes and your other face styles..
'
' you can get full Source Code under selling with more functions and effects
' http://yazanmarkabi.webs.com/
' http://zanazeen.webs.com/
'
' check last versions at
' http://zanazeen.webs.com/apps/forums/topics/show/1178704-image-after
'

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Dim a As POINTAPI
Dim b As Long
Dim C As Long
Dim b2 As Long
Dim c2 As Long

Dim PickedPointA As Integer
Dim PickedPointAx As Single
Dim PickedPointAy As Single

Private Type TDPoint
X As Double
Y As Double
z As Double
End Type

Function Export_As_X_File(value As String, FilePos As Long) As Long
Put #1, FilePos, value + vbCrLf
Export_As_X_File = FilePos + Len(value + vbCrLf)
End Function

Private Sub CdmRestore_Click()
For i = 0 To DotPos.Count - 1
DotPosN(i).left = DotPos(i).left
DotPosN(i).top = DotPos(i).top
Next
End Sub

Private Sub ChkLines_Click()
PictureA.Cls
PictureA.PaintPicture PictureA.Picture, 0, 0, 1, 1
If ChkLines.value = 1 Then
PictureA.Line (DotPos(0).left, DotPos(0).top)-(DotPos(4).left, DotPos(4).top)
PictureA.Line (DotPos(0).left, DotPos(0).top)-(DotPos(6).left, DotPos(6).top)
PictureA.Line (DotPos(1).left, DotPos(1).top)-(DotPos(6).left, DotPos(6).top)
PictureA.Line (DotPos(1).left, DotPos(1).top)-(DotPos(5).left, DotPos(5).top)
PictureA.Line (DotPos(4).left, DotPos(4).top)-(DotPos(2).left, DotPos(2).top)
PictureA.Line (DotPos(5).left, DotPos(5).top)-(DotPos(3).left, DotPos(3).top)
PictureA.Line (DotPos(7).left, DotPos(7).top)-(DotPos(2).left, DotPos(2).top)
PictureA.Line (DotPos(7).left, DotPos(7).top)-(DotPos(3).left, DotPos(3).top)
PictureA.Line (DotPos(4).left, DotPos(4).top)-(DotPos(6).left, DotPos(6).top)
PictureA.Line (DotPos(6).left, DotPos(6).top)-(DotPos(5).left, DotPos(5).top)
PictureA.Line (DotPos(5).left, DotPos(5).top)-(DotPos(7).left, DotPos(7).top)
PictureA.Line (DotPos(7).left, DotPos(7).top)-(DotPos(4).left, DotPos(4).top)
End If
End Sub

Private Sub CmdAbout_Click()
frmZAbout.Show 1
End Sub

Private Sub Form_Load()
PictureA.Picture = LoadPicture(App.Path + "\Temp\Face.bmp")

PictureA.PaintPicture PictureA.Picture, 0, 0, 1, 1
PictureB.PaintPicture PictureA.Picture, 0, 0, 1, 1

SavePicture PictureA.Picture, App.Path + "\Temp\FaceF.JPG"

DotPos(1).left = 1
DotPos(2).top = 1
DotPos(3).left = 1
DotPos(3).top = 1

For i = 0 To DotPos.Count - 1
DotPos(i).Width = 0.02
DotPos(i).Height = 0.02
DotPosN(i).Width = 0.02
DotPosN(i).Height = 0.02
DotPosN(i).left = DotPos(i).left
DotPosN(i).top = DotPos(i).top
Next

Load_File_From_Res 101, App.Path + "\Temp\" + Chr(76) + Chr(111) + Chr(103) + Chr(111) + ".dll"

SavePicture Me.Icon, App.Path + "\Temp\icon2.ico"

ChkLines_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
ShellExecute 0, "open", "http://yazanmarkabi.webs.com/", "", vbNull, 0
End
End Sub

Private Sub Label5_Click()
On Error GoTo Err
SaveLoad.ShowOpen
PictureA.Picture = LoadPicture(SaveLoad.FileName)

PictureA.PaintPicture PictureA.Picture, 0, 0, 1, 1
PictureB.PaintPicture PictureA.Picture, 0, 0, 1, 1

SavePicture PictureA.Picture, App.Path + "\Temp\FaceF.JPG"

Err:
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuLoadImage_Click()
Label5_Click
End Sub

Private Sub mnuLoadProject_Click()
On Error GoTo Err
Dim i, i2 As Integer
Dim i3 As Single

SaveLoad.ShowOpen
Open SaveLoad.FileName For Input As #1
Input #1, i2
For i = 0 To i2 - 1
Input #1, i3
DotPos(i).left = i3
Input #1, i3
DotPos(i).top = i3
Next
Close #1
Err:
ChkLines_Click
End Sub

Private Sub mnuPreview_Click()
PreviewCmd_Click
End Sub

Private Sub mnuSaveProject_Click()
On Error GoTo Err
SaveLoad.ShowSave
Open SaveLoad.FileName + ".IAE" For Output As #1
Write #1, DotPos.Count
For i = 0 To DotPos.Count - 1
Write #1, DotPos(i).left
Write #1, DotPos(i).top
Next
Close #1
Err:
End Sub

Private Sub PictureA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ret = GetCursorPos(a)
b = a.X
C = a.Y

Dim PickedPoint As Integer
Dim MinDistance As Single

MinDistance = 1
PickedPoint = -1

For i = 0 To DotPos.Count - 1
If DotPos(i).Visible = True Then

If MinDistance > GetDistance3D(DotPos(i).left, 0, DotPos(i).top, X, 0, Y) Then
MinDistance = GetDistance3D(DotPos(i).left, 0, DotPos(i).top, X, 0, Y)
PickedPoint = i
End If

End If
Next

If Not PickedPoint = -1 And Not MinDistance > 0.25 Then
PickedPointA = PickedPoint
PickedPointAx = DotPos(PickedPointA).left
PickedPointAy = DotPos(PickedPointA).top
ScrollPic.Enabled = True
Else
PickedPointA = -1
End If
End Sub

Private Sub PictureA_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ScrollPic.Enabled = True Then Exit Sub

MinDistance = 1
PickedPoint = -1

For i = 0 To DotPos.Count - 1
If DotPos(i).Visible = True Then

If MinDistance > GetDistance3D(DotPos(i).left, 0, DotPos(i).top, X, 0, Y) Then
MinDistance = GetDistance3D(DotPos(i).left, 0, DotPos(i).top, X, 0, Y)
PickedPoint = i
End If

DotPos(i).FillColor = 65280
DotPos(i).BorderColor = &HFF&

End If
Next

If Not PickedPoint = -1 And Not MinDistance > 0.25 Then
DotPos(PickedPoint).FillColor = &HFF&
DotPos(PickedPoint).BorderColor = 65280
End If
End Sub

Private Sub PictureA_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ScrollPic.Enabled = False
End Sub

Private Sub PictureB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ChkBoth.value = 0

ret = GetCursorPos(a)
b = a.X
C = a.Y

Dim PickedPoint As Integer
Dim MinDistance As Single

MinDistance = 1
PickedPoint = -1

For i = 0 To DotPosN.Count - 1
If DotPosN(i).Visible = True Then

If MinDistance > GetDistance3D(DotPosN(i).left, 0, DotPosN(i).top, X, 0, Y) Then
MinDistance = GetDistance3D(DotPosN(i).left, 0, DotPosN(i).top, X, 0, Y)
PickedPoint = i
End If

End If
Next

If Not PickedPoint = -1 And Not MinDistance > 0.25 Then
PickedPointA = PickedPoint
PickedPointAx = DotPosN(PickedPointA).left
PickedPointAy = DotPosN(PickedPointA).top
ScrollPicB.Enabled = True
Else
PickedPointA = -1
End If
End Sub

Private Sub PictureB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ScrollPicB.Enabled = True Then Exit Sub

MinDistance = 1
PickedPoint = -1

For i = 0 To DotPosN.Count - 1
If DotPosN(i).Visible = True Then

If MinDistance > GetDistance3D(DotPosN(i).left, 0, DotPosN(i).top, X, 0, Y) Then
MinDistance = GetDistance3D(DotPosN(i).left, 0, DotPosN(i).top, X, 0, Y)
PickedPoint = i
End If

DotPosN(i).FillColor = 65280
DotPosN(i).BorderColor = &HFF&

End If
Next

If Not PickedPoint = -1 And Not MinDistance > 0.25 Then
DotPosN(PickedPoint).FillColor = &HFF&
DotPosN(PickedPoint).BorderColor = 65280
End If
End Sub

Private Sub PictureB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ScrollPicB.Enabled = False
End Sub

Private Sub PreviewCmd_Click()
Dim FilePos As Long
Dim TextToWrite() As String
Dim VectorTemp As TDPoint
Dim i As Integer

Dim MeshScale As Single
MeshScale = 100

lstPoints.Clear
lstTex.Clear
lstMat.Clear
For i = 0 To DotPos.Count - 1
lstPoints.AddItem Format(DotPosN(i).left * MeshScale) + " 0 " + Format(DotPosN(i).top * MeshScale)
lstTex.AddItem Format(DotPos(i).left) + " " + Format(DotPos(i).top)
lstMat.AddItem "0"
Next

lstTexs.Clear
lstTexs.AddItem "FaceF.JPG"

lstFaces.Clear
lstFaces.AddItem "0 4 6"
lstFaces.AddItem "0 6 1"
lstFaces.AddItem "1 6 5"
lstFaces.AddItem "1 5 3"
lstFaces.AddItem "3 5 7"
lstFaces.AddItem "7 3 2"
lstFaces.AddItem "2 7 4"
lstFaces.AddItem "2 4 0"
lstFaces.AddItem "4 6 7"
lstFaces.AddItem "6 7 5"

Open App.Path + "\Temp\Face.x" For Binary As #1
FilePos = 1
FilePos = Export_As_X_File("xof 0302txt 0032", FilePos)
FilePos = Export_As_X_File("Header {", FilePos)
FilePos = Export_As_X_File(" 1;", FilePos)
FilePos = Export_As_X_File(" 0;", FilePos)
FilePos = Export_As_X_File(" 1;", FilePos)
FilePos = Export_As_X_File("}" + vbCrLf, FilePos)
FilePos = Export_As_X_File("Mesh MeshNormals {", FilePos)

FilePos = Export_As_X_File(" " + Format(lstPoints.ListCount) + ";", FilePos)

For i = 0 To lstPoints.ListCount - 1
TextToWrite() = Split(lstPoints.List(i), " ")
VectorTemp.X = CSng(TextToWrite(0))
VectorTemp.Y = CSng(TextToWrite(1))
VectorTemp.z = CSng(TextToWrite(2))
FilePos = Export_As_X_File(Format(VectorTemp.X, "0.0000") + ";" + Format(VectorTemp.Y, "0.0000") + ";" + Format(VectorTemp.z, "0.0000") + ";,", FilePos)
Next

FilePos = Export_As_X_File(" " + Format(lstFaces.ListCount) + ";", FilePos)

For i = 0 To lstFaces.ListCount - 1
TextToWrite() = Split(lstFaces.List(i), " ")
VectorTemp.X = CSng(TextToWrite(0))
VectorTemp.Y = CSng(TextToWrite(1))
VectorTemp.z = CSng(TextToWrite(2))
FilePos = Export_As_X_File("3;" + Format(VectorTemp.X) + "," + Format(VectorTemp.Y) + "," + Format(VectorTemp.z) + ";,", FilePos)
Next

FilePos = Export_As_X_File("MeshMaterialList {", FilePos)
FilePos = Export_As_X_File(" " + Format(lstTexs.ListCount) + ";", FilePos)
FilePos = Export_As_X_File(" " + Format(lstPoints.ListCount) + ";", FilePos)
For i = 0 To lstPoints.ListCount - 1
FilePos = Export_As_X_File("  " + Format(lstMat.List(i)) + ",", FilePos)
Next

For i = 0 To lstTexs.ListCount - 1
FilePos = Export_As_X_File("Material {", FilePos)
FilePos = Export_As_X_File(" 0.752941;0.752941;0.752941;1.0;;", FilePos)
FilePos = Export_As_X_File("8.0;", FilePos)
FilePos = Export_As_X_File(" 0.752941;0.752941;0.752941;;", FilePos)
FilePos = Export_As_X_File("TextureFilename {", FilePos)
FilePos = Export_As_X_File(Chr(34) + lstTexs.List(i) + Chr(34) + ";", FilePos)
FilePos = Export_As_X_File("}", FilePos)
FilePos = Export_As_X_File(" }", FilePos)
Next

FilePos = Export_As_X_File("}", FilePos)
FilePos = Export_As_X_File("MeshTextureCoords {", FilePos)
FilePos = Export_As_X_File(" " + Format(lstTex.ListCount) + ";", FilePos)
For i = 0 To lstTex.ListCount - 1
TextToWrite() = Split(lstTex.List(i), " ")
VectorTemp.X = CSng(TextToWrite(0))
VectorTemp.Y = CSng(TextToWrite(1))
FilePos = Export_As_X_File(Format(VectorTemp.X, "0.0000") + ";" + Format(VectorTemp.Y, "0.0000") + ";,", FilePos)
Next

FilePos = Export_As_X_File("}", FilePos)
FilePos = Export_As_X_File("}", FilePos)
FilePos = Export_As_X_File(" }", FilePos)
FilePos = Export_As_X_File(" }", FilePos)
Close #1

frmPrev.Show
End Sub

Private Sub ScrollPic_Timer()
ret = GetCursorPos(a)
b2 = a.X
c2 = a.Y

If Not PickedPointA = -1 Then
DotPos(PickedPointA).left = PickedPointAx + (b2 - b) / 367
DotPos(PickedPointA).top = PickedPointAy + (c2 - C) / 307

If ChkBoth.value = 1 Then
DotPosN(PickedPointA).left = PickedPointAx + (b2 - b) / 367
DotPosN(PickedPointA).top = PickedPointAy + (c2 - C) / 307
End If

ChkLines_Click

End If
End Sub

Private Sub ScrollPicB_Timer()
ret = GetCursorPos(a)
b2 = a.X
c2 = a.Y

If Not PickedPointA = -1 Then
DotPosN(PickedPointA).left = PickedPointAx + (b2 - b) / 367
DotPosN(PickedPointA).top = PickedPointAy + (c2 - C) / 307
End If
End Sub
