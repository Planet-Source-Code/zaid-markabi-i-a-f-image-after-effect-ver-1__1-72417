VERSION 5.00
Begin VB.Form frmZAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " About  I.A.E"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8340
   Icon            =   "frmZAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   8340
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright Registered 2009-2010 (c) ZANAZEEN co."
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   4
      Left            =   1200
      TabIndex        =   8
      Top             =   3360
      Width           =   3690
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Zaid Markabi ,  Arabic Syrian Programmer"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   3
      Left            =   5280
      TabIndex        =   7
      Top             =   3360
      Width           =   2925
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "This is trial version, you can get Full version from http://yazanmarkabi.webs.com/"
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
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   3000
      Width           =   6975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   6720
      TabIndex        =   4
      Top             =   960
      Width           =   1395
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderStyle     =   3  'Dot
      Index           =   0
      X1              =   2640
      X2              =   4800
      Y1              =   960
      Y2              =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Image After Effect"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   675
      Left            =   2520
      TabIndex        =   5
      Top             =   360
      Width           =   5040
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   8
      Index           =   1
      X1              =   0
      X2              =   8400
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   8
      Index           =   0
      X1              =   0
      X2              =   8400
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "by  Zaid Markabi"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   330
      Left            =   6120
      TabIndex        =   3
      Top             =   2280
      Width           =   2040
   End
   Begin VB.Shape DotPos 
      BorderColor     =   &H000000FF&
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   5
      Left            =   4680
      Shape           =   3  'Circle
      Top             =   240
      Width           =   135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderStyle     =   3  'Dot
      Index           =   1
      X1              =   5640
      X2              =   4800
      Y1              =   1200
      Y2              =   240
   End
   Begin VB.Shape DotPos 
      BorderColor     =   &H000000FF&
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   1200
      Width           =   135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderStyle     =   3  'Dot
      Index           =   2
      X1              =   8040
      X2              =   5760
      Y1              =   720
      Y2              =   1320
   End
   Begin VB.Shape DotPos 
      BorderColor     =   &H000000FF&
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   7920
      Shape           =   3  'Circle
      Top             =   720
      Width           =   135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderStyle     =   3  'Dot
      Index           =   3
      X1              =   8280
      X2              =   8040
      Y1              =   1320
      Y2              =   720
   End
   Begin VB.Shape DotPos 
      BorderColor     =   &H000000FF&
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   1680
      Width           =   135
   End
   Begin VB.Shape DotPos 
      BorderColor     =   &H000000FF&
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   3
      Left            =   8280
      Shape           =   3  'Circle
      Top             =   1320
      Width           =   135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderStyle     =   3  'Dot
      Index           =   4
      X1              =   1800
      X2              =   2640
      Y1              =   1680
      Y2              =   960
   End
   Begin VB.Shape DotPos 
      BorderColor     =   &H000000FF&
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   4
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.yazanmarkabi.webs.com/"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   1560
      TabIndex        =   2
      Top             =   2400
      Width           =   2685
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email : ZaidMarkabi@yahoo.com"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   0
      Left            =   2520
      TabIndex        =   1
      Top             =   1440
      Width           =   2340
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QQ : 1147529632"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   2
      Left            =   3000
      TabIndex        =   0
      Top             =   1680
      Width           =   1275
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmZAbout.frx":038A
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image BackImg 
      Height          =   330
      Left            =   0
      Picture         =   "frmZAbout.frx":2054
      Top             =   0
      Width           =   420
   End
End
Attribute VB_Name = "frmZAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
BackImg.Picture = LoadPicture(App.Path + "\Temp\" + Chr(76) + Chr(111) + Chr(103) + Chr(111) + "2.dll")

Label4(3).Caption = "Zaid Markabi ,  Arabic Syrian Programmer"
Label4(4).Caption = "Copyright Registered 2009-2010 (c) ZANAZEEN co."
Label4(0).Caption = "Email : ZaidMarkabi@yahoo.com"
Label4(2).Caption = "QQ : 1147529632"
Label4(1).Caption = "http://www.yazanmarkabi.webs.com/"
Label1.Caption = "Image After Effect"
Label3.Caption = "by  Zaid Markabi"

End Sub

Private Sub Label5_Click()
ShellExecute 0, "open", "http://www.zanazeen.webs.com/apps/forums/topics/show/1178704-image-after", "", vbNull, 0
End Sub
