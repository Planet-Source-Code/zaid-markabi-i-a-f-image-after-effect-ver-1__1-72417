VERSION 5.00
Begin VB.Form EfcColor 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Color Balance"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2175
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   2175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar HcolG 
      Height          =   255
      Left            =   360
      Max             =   100
      Min             =   -100
      TabIndex        =   8
      Top             =   1560
      Width           =   1695
   End
   Begin VB.HScrollBar Hcol 
      Height          =   255
      Index           =   3
      Left            =   360
      Max             =   100
      TabIndex        =   7
      Top             =   1200
      Value           =   100
      Width           =   1695
   End
   Begin VB.HScrollBar Hcol 
      Height          =   255
      Index           =   2
      Left            =   360
      Max             =   100
      TabIndex        =   5
      Top             =   840
      Value           =   100
      Width           =   1695
   End
   Begin VB.HScrollBar Hcol 
      Height          =   255
      Index           =   1
      Left            =   360
      Max             =   100
      TabIndex        =   3
      Top             =   480
      Value           =   100
      Width           =   1695
   End
   Begin VB.HScrollBar Hcol 
      Height          =   255
      Index           =   0
      Left            =   360
      Max             =   100
      TabIndex        =   1
      Top             =   120
      Value           =   100
      Width           =   1695
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   120
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   135
   End
End
Attribute VB_Name = "EfcColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
Me.Hide
End Sub

Private Sub Hcol_Change(Index As Integer)
Mesh.SetColor RGBA(Hcol(0).value / 100, Hcol(1).value / 100, Hcol(2).value / 100, Hcol(3).value / 100)
End Sub

Private Sub Hcol_Scroll(Index As Integer)
Hcol_Change (Index)
End Sub

Private Sub HcolG_Change()
AddColorLightingA = HcolG.value / 100
End Sub

Private Sub HcolG_Scroll()
HcolG_Change
End Sub
