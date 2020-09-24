VERSION 5.00
Begin VB.Form EfcLighting 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Add Lighting"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1935
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   1935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar Rcolor 
      Height          =   255
      Left            =   360
      Max             =   100
      TabIndex        =   6
      Top             =   1440
      Value           =   50
      Width           =   1455
   End
   Begin VB.HScrollBar Gcolor 
      Height          =   255
      Left            =   360
      Max             =   100
      TabIndex        =   5
      Top             =   1680
      Value           =   50
      Width           =   1455
   End
   Begin VB.HScrollBar Bcolor 
      Height          =   255
      Left            =   360
      Max             =   100
      TabIndex        =   4
      Top             =   1920
      Value           =   50
      Width           =   1455
   End
   Begin VB.CheckBox CheckUL 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.CheckBox CheckUR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
   Begin VB.CheckBox CheckDL 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox CheckDR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   255
   End
   Begin VB.Image SmallerImage 
      Height          =   1185
      Left            =   360
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "EfcLighting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Apply_Changes()
AddColorLighting1 = CheckUL.value
AddColorLighting2 = CheckUR.value
AddColorLighting3 = CheckDL.value
AddColorLighting4 = CheckDR.value
AddColorLightingR = Rcolor.value / 100
AddColorLightingG = Gcolor.value / 100
AddColorLightingB = Bcolor.value / 100
End Sub

Private Sub Bcolor_Change()
Apply_Changes
End Sub

Private Sub Bcolor_Scroll()
Apply_Changes
End Sub

Private Sub CheckDL_Click()
Apply_Changes
End Sub

Private Sub CheckDR_Click()
Apply_Changes
End Sub

Private Sub CheckUL_Click()
Apply_Changes
End Sub

Private Sub CheckUR_Click()
Apply_Changes
End Sub

Private Sub Form_Load()
SmallerImage.Picture = frmMain.PictureA.Picture
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
Me.Hide
End Sub

Private Sub Gcolor_Change()
Apply_Changes
End Sub

Private Sub Gcolor_Scroll()
Apply_Changes
End Sub

Private Sub Rcolor_Change()
Apply_Changes
End Sub

Private Sub Rcolor_Scroll()
Apply_Changes
End Sub
