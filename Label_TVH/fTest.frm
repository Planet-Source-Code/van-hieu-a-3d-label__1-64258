VERSION 5.00
Object = "{B30AF7F6-8E0C-4AC7-951B-3C1A5DF856FD}#1.0#0"; "Label_TVH.ocx"
Begin VB.Form fTest 
   BackColor       =   &H00000000&
   Caption         =   "Use Font Unicode - Don't use Form 2.0"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   ScaleHeight     =   414
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   356
   StartUpPosition =   3  'Windows Default
   Begin TVH.Label_TVH T 
      Height          =   900
      Index           =   11
      Left            =   720
      TabIndex        =   11
      Top             =   4485
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   1588
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Transparent"
      AutoSize        =   -1  'True
      BackColor       =   255
      ForeColor       =   16776960
      BorderColor     =   255
      BorderSize      =   3
      BorderStyle     =   5
      OutlineColor    =   65280
      Shadow          =   -1  'True
      ShadowDepth     =   10
      ShadowStyle     =   0
      ShadowColorStart=   255
      Alignment       =   2
      BackColorStyle  =   1
      GradientBackColorStyle=   4
      GradientBackColorStart=   16777215
      GradientBackColorEnd=   0
   End
   Begin TVH.Label_TVH T 
      Height          =   465
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   820
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Tie61ng Vie65t"
      AutoSize        =   -1  'True
      BackColor       =   16777215
      ForeColor       =   65535
      BorderColor     =   16776960
      Transparent     =   0   'False
      OutlineColor    =   16776960
      Shadow          =   -1  'True
      ShadowDepth     =   2
      ShadowStyle     =   0
      Alignment       =   2
      BackColorStyle  =   1
      GradientBackColorStart=   16777215
      GradientBackColorEnd=   0
   End
   Begin TVH.Label_TVH T 
      Height          =   930
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Top             =   105
      Width           =   3180
      _ExtentX        =   5609
      _ExtentY        =   1640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Tie61ng Vie65t du2ng WordWrap"
      AutoSize        =   -1  'True
      WordWrap        =   -1  'True
      BackColor       =   16777215
      ForeColor       =   65535
      BorderColor     =   16776960
      BorderSize      =   3
      BorderStyle     =   5
      Transparent     =   0   'False
      OutlineColor    =   16776960
      Shadow          =   -1  'True
      ShadowDepth     =   4
      ShadowStyle     =   0
      ShadowColorStart=   65280
      Alignment       =   2
      BackColorStyle  =   1
      GradientBackColorStyle=   3
      GradientBackColorStart=   16777215
      GradientBackColorEnd=   0
   End
   Begin TVH.Label_TVH T 
      Height          =   465
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   820
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Tie61ng Vie65t"
      AutoSize        =   -1  'True
      BackColor       =   16777215
      ForeColor       =   65535
      BorderColor     =   16776960
      BorderStyle     =   1
      Transparent     =   0   'False
      OutlineColor    =   16776960
      Shadow          =   -1  'True
      ShadowDepth     =   2
      ShadowStyle     =   0
      Alignment       =   2
      BackColorStyle  =   1
      GradientBackColorStyle=   1
      GradientBackColorStart=   16777215
      GradientBackColorEnd=   0
   End
   Begin TVH.Label_TVH T 
      Height          =   465
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1095
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   820
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Font Unicode"
      BackColor       =   16777215
      ForeColor       =   255
      BorderColor     =   255
      BorderStyle     =   2
      Transparent     =   0   'False
      OutlineColor    =   65280
      ShadowDepth     =   2
      ShadowStyle     =   0
      Alignment       =   2
      BackColorStyle  =   1
      GradientBackColorStyle=   2
      GradientBackColorStart=   16777215
      GradientBackColorEnd=   0
   End
   Begin TVH.Label_TVH T 
      Height          =   645
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   1138
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "3D Text"
      AutoSize        =   -1  'True
      BackColor       =   16777215
      ForeColor       =   65280
      BorderColor     =   16776960
      Transparent     =   0   'False
      OutlineColor    =   16776960
      Shadow          =   -1  'True
      ShadowDepth     =   4
      ShadowStyle     =   0
      ShadowColorStart=   65535
      Alignment       =   2
      BackColorStyle  =   1
      GradientBackColorStyle=   4
      GradientBackColorStart=   16777215
      GradientBackColorEnd=   0
   End
   Begin TVH.Label_TVH T 
      Height          =   645
      Index           =   5
      Left            =   105
      TabIndex        =   5
      Top             =   2400
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   1138
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "3D Text"
      AutoSize        =   -1  'True
      BackColor       =   16777215
      ForeColor       =   65280
      BorderColor     =   16776960
      Transparent     =   0   'False
      OutlineColor    =   16776960
      Shadow          =   -1  'True
      ShadowDepth     =   4
      ShadowStyle     =   1
      ShadowColorStart=   65535
      Alignment       =   2
      BackColorStyle  =   1
      GradientBackColorStyle=   4
      GradientBackColorStart=   16777215
      GradientBackColorEnd=   0
   End
   Begin TVH.Label_TVH T 
      Height          =   645
      Index           =   6
      Left            =   2265
      TabIndex        =   6
      Top             =   1680
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   1138
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "3D Text"
      AutoSize        =   -1  'True
      BackColor       =   16777215
      ForeColor       =   65280
      BorderColor     =   16776960
      Transparent     =   0   'False
      OutlineColor    =   16776960
      Shadow          =   -1  'True
      ShadowDepth     =   4
      ShadowStyle     =   2
      ShadowColorStart=   65535
      Alignment       =   2
      BackColorStyle  =   1
      GradientBackColorStyle=   4
      GradientBackColorStart=   16777215
      GradientBackColorEnd=   0
   End
   Begin TVH.Label_TVH T 
      Height          =   645
      Index           =   7
      Left            =   2265
      TabIndex        =   7
      Top             =   2400
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   1138
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "3D Text"
      AutoSize        =   -1  'True
      BackColor       =   16777215
      ForeColor       =   65280
      BorderColor     =   16776960
      Transparent     =   0   'False
      OutlineColor    =   16776960
      Shadow          =   -1  'True
      ShadowDepth     =   4
      ShadowColorStart=   65535
      Alignment       =   2
      BackColorStyle  =   1
      GradientBackColorStyle=   4
      GradientBackColorStart=   16777215
      GradientBackColorEnd=   0
   End
   Begin TVH.Label_TVH T 
      Height          =   615
      Index           =   8
      Left            =   105
      TabIndex        =   8
      Top             =   3120
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Outline Text"
      AutoSize        =   -1  'True
      BackColor       =   255
      ForeColor       =   16711680
      BorderColor     =   16776960
      BorderSize      =   2
      BorderStyle     =   2
      Transparent     =   0   'False
      OutlineColor    =   65280
      ShadowDepth     =   4
      ShadowStyle     =   0
      ShadowColorStart=   65535
      Alignment       =   2
      BackColorStyle  =   1
      GradientBackColorStyle=   6
      GradientBackColorStart=   16777215
      GradientBackColorEnd=   16776960
   End
   Begin TVH.Label_TVH T 
      Height          =   645
      Index           =   9
      Left            =   105
      TabIndex        =   9
      Top             =   3765
      Width           =   4845
      _ExtentX        =   8652
      _ExtentY        =   1244
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "3D && Outline Text"
      AutoSize        =   -1  'True
      BackColor       =   255
      ForeColor       =   16776960
      BorderColor     =   16776960
      BorderSize      =   3
      BorderStyle     =   5
      Transparent     =   0   'False
      OutlineColor    =   255
      Shadow          =   -1  'True
      ShadowDepth     =   4
      ShadowStyle     =   0
      ShadowColorStart=   255
      ShadowColorEnd  =   16744576
      Alignment       =   2
      BackColorStyle  =   1
      GradientBackColorStyle=   7
      GradientBackColorStart=   16777215
      GradientBackColorEnd=   255
   End
   Begin TVH.Label_TVH T 
      Height          =   585
      Index           =   10
      Left            =   105
      TabIndex        =   10
      Top             =   4470
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   1032
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Flat Text"
      BackColor       =   255
      ForeColor       =   255
      BorderColor     =   16776960
      BorderStyle     =   1
      Transparent     =   0   'False
      OutlineColor    =   65280
      ShadowDepth     =   4
      ShadowStyle     =   1
      ShadowColorStart=   65535
      Alignment       =   1
      BackColorStyle  =   1
      GradientBackColorStyle=   8
      GradientBackColorStart=   16777215
      GradientBackColorEnd=   16711935
   End
   Begin TVH.Label_TVH T 
      Height          =   585
      Index           =   12
      Left            =   -15
      TabIndex        =   12
      Top             =   5625
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   1032
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "> Move Here <"
      BackColor       =   255
      ForeColor       =   0
      BorderColor     =   16776960
      BorderStyle     =   2
      Transparent     =   0   'False
      OutlineColor    =   65280
      ShadowDepth     =   4
      ShadowStyle     =   0
      ShadowColorStart=   65535
      Alignment       =   2
      BackColorStyle  =   1
      GradientBackColorStyle=   7
      GradientBackColorStart=   16777215
      GradientBackColorEnd=   16744576
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub T_MouseLeave(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index = 12 Then
        T(12).Text = "> Move Here <"
    End If
End Sub

Private Sub T_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index = 12 Then
        T(12).Text = "Close this!"
    End If
End Sub
