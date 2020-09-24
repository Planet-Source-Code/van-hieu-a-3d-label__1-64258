VERSION 5.00
Begin VB.UserControl Label_TVH 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2475
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   63
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   165
   Begin VB.PictureBox tPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   480
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "Label_TVH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'///////////////////////////////////////// Truong Van Hieu ////////////////////////////////////////
'////////////////////////////////// tvhhh2003@yahoo.com /////////////////////////////////////
'//////////////////////////////////// Special for Vietnamese /////////////////////////////////////

Option Explicit

Const DT_TOP = &H0
Const DT_LEFT = &H0
Const DT_CENTER = &H1
Const DT_RIGHT = &H2
Const DT_VCENTER = &H4                        ' Canh giua Text theo chieu dung (Chi 1 do`ng)
Const DT_BOTTOM = &H8
Const DT_WORDBREAK = &H10                 ' Tu dong xuong hang
Const DT_SINGLELINE = &H20                    ' Hien thi Text 1 dong, khong xuong hang
Const DT_EXPANDTABS = &H40
Const DT_TABSTOP = &H80
Const DT_NOCLIP = &H100
Const DT_EXTERNALLEADING = &H200
Const DT_CALCRECT = &H400
Const DT_NOPREFIX = &H800
Const DT_INTERNAL = &H1000
Private Declare Function DrawTextA Lib "user32" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Enum E_ShadowStyle
    sLeftTop = 0
    sRightTop = 1
    sLeftBottom = 2
    sRightBottom = 3
End Enum

Public Enum E_BorderStyle
    None = 0
    Flat = 1
    Outline = 2
    [3D] = 3
    Frame1 = 4
    Frame2 = 5
End Enum

Enum E_Alignment
    aLeft = 0
    aRight = 1
    aCenter = 2
End Enum

Enum E_BackColorStyle
    aSingleColor = 0
    aGradientColor = 1
End Enum

Enum E_GradientBackColorStyle
    aLeftToRight = 0
    aRightToLeft = 1
    aTopToBottom = 2
    aBottomToTop = 3
    aLeftTopToRightBottom = 4
    aLeftBottomToRightTop = 5
    aCenterToLeftRight = 6
    aCenterToTopBottom = 7
    aCenterToLeftTopNRightBottom = 8
    aCenterToLeftBottomNRightTop = 9
End Enum

Const TransColor = &H8000000F

'--------------------Private
'_____Attribute

Private m_AutoSize As Boolean
Private m_BackColor As OLE_COLOR
Private m_ForeColor As OLE_COLOR
Private m_BorderColor As OLE_COLOR
Private m_BorderStyle As E_BorderStyle
Private m_BorderSize As Long
Private m_Text As String
Private m_Font As StdFont
Private m_WordWrap As Boolean
Private m_TiengViet As Boolean
Private m_Transparent As Boolean
Private m_EdgeSpace As Long
Private m_OutlineColor As OLE_COLOR
Private m_Shadow As Boolean
Private m_ShadowDepth As Integer
Private m_ShadowStyle As E_ShadowStyle
Private m_ShadowColorStart As OLE_COLOR
Private m_ShadowColorEnd As OLE_COLOR
Private m_Alignment As E_Alignment
Private m_LineCount As Long
Private m_BackColorStyle As E_BackColorStyle
Private m_GradientBackColorStyle As E_GradientBackColorStyle
Private m_GradientBackColorStart As OLE_COLOR
Private m_GradientBackColorEnd As OLE_COLOR

'______Other
Private MeChanged As Boolean
'----------------------End Private

'--------------------------------------------------
'--------Events--------------------------------
'--------------------------------------------------
Event Change()
Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseLeave(Button As Integer, Shift As Integer, x As Single, y As Single)


Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

'--------------------------------------------------
'--------UserControl-------------------------
'--------------------------------------------------
Private Sub UserControl_Initialize()
    m_Text = "Nha4n Tie61ng Vie65t" ' "Extender.Name"
    m_TiengViet = True
    m_BorderSize = 1
    m_EdgeSpace = 1
    m_BackColor = &H8000000F
    m_BorderColor = 0
    m_ForeColor = 0
    m_OutlineColor = &HFFFFFF
    Set m_Font = UserControl.Font
    tPic.BackColor = TransColor
    mUnicode.KhoiTao
    Fresh
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
With UserControl
    If x < 0 Or y < 0 Or x > .ScaleWidth Or y > .ScaleHeight Then
        ReleaseCapture
        RaiseEvent MouseLeave(Button, Shift, x, y)
    Else
        SetCapture .hwnd
        RaiseEvent MouseMove(Button, Shift, x, y)
    End If
End With
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
    Set m_Font = .ReadProperty("Font", Font)
    Set UserControl.Font = m_Font
    m_Text = .ReadProperty("Text", "Tie61ng Viet")
    m_AutoSize = .ReadProperty("AutoSize", False)
    m_WordWrap = .ReadProperty("WordWrap", False)
    m_TiengViet = .ReadProperty("TiengViet", True)
    m_BackColor = .ReadProperty("BackColor", &H8000000F)
    m_ForeColor = .ReadProperty("ForeColor", &H80000008)
    m_BorderColor = .ReadProperty("BorderColor", &H80000008)
    m_BorderSize = .ReadProperty("BorderSize", 1)
    m_BorderStyle = .ReadProperty("BorderStyle", 0)
    m_Transparent = .ReadProperty("Transparent", True)
    m_EdgeSpace = .ReadProperty("EdgeSpace", 1)
    m_OutlineColor = .ReadProperty("OutlineColor", &HFFFFFF)
    m_Shadow = .ReadProperty("Shadow", False)
    m_ShadowDepth = .ReadProperty("ShadowDepth", 3)
    m_ShadowStyle = .ReadProperty("ShadowStyle", sRightBottom)
    m_ShadowColorStart = .ReadProperty("ShadowColorStart", m_ForeColor)
    m_ShadowColorEnd = .ReadProperty("ShadowColorEnd", 0)
    m_Alignment = .ReadProperty("Alignment", 0)
    m_BackColorStyle = .ReadProperty("BackColorStyle", 0)
    m_GradientBackColorStyle = .ReadProperty("GradientBackColorStyle", 0)
    m_GradientBackColorStart = .ReadProperty("GradientBackColorStart", 0)
    m_GradientBackColorEnd = .ReadProperty("GradientBackColorEnd", RGB(255, 255, 255))
    Fresh
    
End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    Call .WriteProperty("Font", m_Font, "Arial")
    Call .WriteProperty("Text", m_Text, "Tie61ng Viet")
    Call .WriteProperty("AutoSize", m_AutoSize, False)
    Call .WriteProperty("WordWrap", m_WordWrap, False)
    Call .WriteProperty("TiengViet", m_TiengViet, True)
    Call .WriteProperty("BackColor", m_BackColor, &H8000000F)
    Call .WriteProperty("ForeColor", m_ForeColor, &H80000008)
    Call .WriteProperty("BorderColor", m_BorderColor, &H80000008)
    Call .WriteProperty("BorderSize", m_BorderSize, 1)
    Call .WriteProperty("BorderStyle", m_BorderStyle, 0)
    Call .WriteProperty("Transparent", m_Transparent, True)
    Call .WriteProperty("EdgeSpace", m_EdgeSpace, 1)
    Call .WriteProperty("OutlineColor", m_OutlineColor, &HFFFFFF)
    Call .WriteProperty("Shadow", m_Shadow, False)
    Call .WriteProperty("ShadowDepth", m_ShadowDepth, 3)
    Call .WriteProperty("ShadowStyle", m_ShadowStyle, sRightBottom)
    Call .WriteProperty("ShadowColorStart", m_ShadowColorStart, m_ForeColor)
    Call .WriteProperty("ShadowColorEnd", m_ShadowColorEnd, 0)
    Call .WriteProperty("Alignment", m_Alignment, 0)
    Call .WriteProperty("BackColorStyle", m_BackColorStyle, 0)
    Call .WriteProperty("GradientBackColorStyle", m_GradientBackColorStyle, 0)
    Call .WriteProperty("GradientBackColorStart", m_GradientBackColorStart, 0)
    Call .WriteProperty("GradientBackColorEnd", m_GradientBackColorEnd, RGB(255, 255, 255))
End With
End Sub

Private Sub UserControl_Resize()
    Fresh
End Sub



'--------------------------------------------------
'--------Properties---------------------------
'--------------------------------------------------
Public Property Get Text() As String
    Text = m_Text
End Property

Public Property Let Text(new_Text As String)
    m_Text = new_Text
    PropertyChanged "Text"
    Fresh
End Property

Public Property Get AutoSize() As Boolean
    AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(new_AutoSize As Boolean)
    m_AutoSize = new_AutoSize
    PropertyChanged "AutoSize"
    Fresh
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(new_BackColor As OLE_COLOR)
    m_BackColor = new_BackColor
    PropertyChanged "BackColor"
    Fresh
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(new_ForeColor As OLE_COLOR)
    m_ForeColor = new_ForeColor
    PropertyChanged "ForeColor"
    Fresh
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(new_BorderColor As OLE_COLOR)
    m_BorderColor = new_BorderColor
    PropertyChanged "BorderColor"
    Fresh
End Property

Public Property Get BorderSize() As Long
    BorderSize = m_BorderSize
End Property

Public Property Let BorderSize(new_BorderSize As Long)
    If new_BorderSize >= 0 Then m_BorderSize = new_BorderSize
    PropertyChanged "BorderSize"
    Fresh
End Property

Public Property Get Font() As StdFont
    Set Font = m_Font
End Property

Public Property Set Font(new_Font As StdFont)
    Set m_Font = new_Font
    PropertyChanged "Font"
    Set UserControl.Font = m_Font
    Fresh
End Property

Public Property Get WordWrap() As Boolean
    WordWrap = m_WordWrap
End Property

Public Property Let WordWrap(new_WordWrap As Boolean)
    m_WordWrap = new_WordWrap
    PropertyChanged "WordWrap"
    Fresh
End Property

Public Property Get TiengViet() As Boolean
    TiengViet = m_TiengViet
End Property

Public Property Let TiengViet(new_TiengViet As Boolean)
    m_TiengViet = new_TiengViet
    PropertyChanged "TiengViet"
    Fresh
End Property

Public Property Get BorderStyle() As E_BorderStyle
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(new_BorderStyle As E_BorderStyle)
    m_BorderStyle = new_BorderStyle
    PropertyChanged "BorderStyle"
    Fresh
End Property

Public Property Get Transparent() As Boolean
    Transparent = m_Transparent
End Property

Public Property Let Transparent(new_Transparent As Boolean)
    m_Transparent = new_Transparent
    PropertyChanged "Transparent"
    Fresh
End Property

Public Property Get EdgeSpace() As Long
    EdgeSpace = m_EdgeSpace
End Property

Public Property Let EdgeSpace(new_EdgeSpace As Long)
    If new_EdgeSpace >= 0 Then m_EdgeSpace = new_EdgeSpace
    PropertyChanged "EdgeSpace"
    Fresh
End Property

Public Property Get OutlineColor() As OLE_COLOR
    OutlineColor = m_OutlineColor
End Property

Public Property Let OutlineColor(new_OutlineColor As OLE_COLOR)
    If new_OutlineColor >= 0 Then m_OutlineColor = new_OutlineColor
    PropertyChanged "OutlineColor"
    Fresh
End Property

Public Property Get Shadow() As Boolean
    Shadow = m_Shadow
End Property

Public Property Let Shadow(new_Shadow As Boolean)
    m_Shadow = new_Shadow
    PropertyChanged "Shadow"
    Fresh
End Property

Public Property Get ShadowDepth() As Integer
    ShadowDepth = m_ShadowDepth
End Property

Public Property Let ShadowDepth(new_ShadowDepth As Integer)
    If new_ShadowDepth >= 0 Then m_ShadowDepth = new_ShadowDepth
    PropertyChanged "ShadowDepth"
    Fresh
End Property

Public Property Get ShadowStyle() As E_ShadowStyle
    ShadowStyle = m_ShadowStyle
End Property

Public Property Let ShadowStyle(new_ShadowStyle As E_ShadowStyle)
    m_ShadowStyle = new_ShadowStyle
    PropertyChanged "ShadowStyle"
    Fresh
End Property

Public Property Get ShadowColorStart() As OLE_COLOR
    ShadowColorStart = m_ShadowColorStart
End Property

Public Property Let ShadowColorStart(new_ShadowColorStart As OLE_COLOR)
    m_ShadowColorStart = new_ShadowColorStart
    PropertyChanged "ShadowColorStart"
    Fresh
End Property

Public Property Get ShadowColorEnd() As OLE_COLOR
    ShadowColorEnd = m_ShadowColorEnd
End Property

Public Property Let ShadowColorEnd(new_ShadowColorEnd As OLE_COLOR)
    m_ShadowColorEnd = new_ShadowColorEnd
    PropertyChanged "ShadowColorEnd"
    Fresh
End Property

Public Property Get Alignment() As E_Alignment
    Alignment = m_Alignment
End Property

Public Property Let Alignment(new_Alignment As E_Alignment)
    m_Alignment = new_Alignment
    PropertyChanged "Alignment"
    Fresh
End Property

Public Property Get BackColorStyle() As E_BackColorStyle
    BackColorStyle = m_BackColorStyle
End Property

Public Property Let BackColorStyle(new_BackColorStyle As E_BackColorStyle)
    m_BackColorStyle = new_BackColorStyle
    PropertyChanged "BackColorStyle"
    Fresh
End Property

Public Property Get GradientBackColorStyle() As E_GradientBackColorStyle
    GradientBackColorStyle = m_GradientBackColorStyle
End Property

Public Property Let GradientBackColorStyle(new_GradientBackColorStyle As E_GradientBackColorStyle)
    m_GradientBackColorStyle = new_GradientBackColorStyle
    PropertyChanged "GradientBackColorStyle"
    Fresh
End Property

Public Property Get GradientBackColorStart() As OLE_COLOR
    GradientBackColorStart = m_GradientBackColorStart
End Property

Public Property Let GradientBackColorStart(new_GradientBackColorStart As OLE_COLOR)
    m_GradientBackColorStart = new_GradientBackColorStart
    PropertyChanged "GradientBackColorStart"
    Fresh
End Property

Public Property Get GradientBackColorEnd() As OLE_COLOR
    GradientBackColorEnd = m_GradientBackColorEnd
End Property

Public Property Let GradientBackColorEnd(new_GradientBackColorEnd As OLE_COLOR)
    m_GradientBackColorEnd = new_GradientBackColorEnd
    PropertyChanged "GradientBackColorEnd"
    Fresh
End Property

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Property Get hdc() As Long
    hdc = UserControl.hdc
End Property

Public Property Get LineCount() As Long
    LineCount = m_LineCount
End Property


'--------------------------------------------------
'--------Function more---------------------
'--------------------------------------------------
Public Function VNIToUnicode(s As String) As String
    VNIToUnicode = Text_To_Unicode(s)
End Function

Private Sub Fresh()
With UserControl
    If m_AutoSize Then
        MeChanged = True
        Dim Dong As Long
        If m_WordWrap Then
            '.ScaleWidth = (.TextWidth(IIf(m_TiengViet, XoaDau(VNIToUnicode(m_Text)), m_Text)) + 2 * BorderSize + 2 * EdgeSpace + IIf(m_Shadow, m_ShadowDepth, 0)) * 15
            Set tPic = Nothing
            Set tPic.Font = m_Font
            Dim t As RECT
            t.Left = BorderSize + EdgeSpace '+ IIf(m_Shadow And (m_ShadowStyle = sLeftTop Or m_ShadowStyle = sLeftBottom), m_ShadowDepth, 0)
            t.Top = BorderSize + EdgeSpace '+ IIf(m_Shadow And (m_ShadowStyle = sLeftTop Or m_ShadowStyle = sRightTop), m_ShadowDepth, 0)
            t.Bottom = 100 '.ScaleHeight - BorderSize - EdgeSpace
            t.Right = .ScaleWidth - BorderSize - EdgeSpace - IIf(m_Shadow, m_ShadowDepth, 0)
            Dong = DrawText(tPic.hdc, IIf(m_TiengViet, VNIToUnicode(m_Text), m_Text), t, True)
            Dong = Dong / tPic.TextHeight(m_Text)
        Else
            .Width = (.TextWidth(IIf(m_TiengViet, XoaDau(VNIToUnicode(m_Text)), m_Text)) + 2 * BorderSize + 2 * EdgeSpace + IIf(m_Shadow, m_ShadowDepth, 0)) * 15
            Dong = 1
        End If
        .Height = (.TextHeight(Left(m_Text, 1)) * Dong + 2 * BorderSize + 2 * EdgeSpace + IIf(m_Shadow, m_ShadowDepth, 0)) * 15
    End If
    Draw IIf(m_TiengViet, VNIToUnicode(m_Text), m_Text)
    RaiseEvent Change
End With
End Sub

Private Sub Draw(s As String)
Dim tR As RECT
Dim i As Integer
With UserControl
    If Trim(s) = "" Then Exit Sub
    'Set Backcolor
    tPic.BackColor = IIf(m_Transparent, TransColor, m_BackColor)
    Cls
    Set .Picture = Nothing
    Set tPic = Nothing
    Set tPic.Font = .Font
    tPic.Width = .ScaleWidth
    tPic.Height = .ScaleHeight
    tR.Left = BorderSize + EdgeSpace + IIf(m_Shadow And (m_ShadowStyle = sLeftTop Or m_ShadowStyle = sLeftBottom), m_ShadowDepth, 0)
    tR.Top = BorderSize + EdgeSpace + IIf(m_Shadow And (m_ShadowStyle = sLeftTop Or m_ShadowStyle = sRightTop), m_ShadowDepth, 0)
    tR.Bottom = .ScaleHeight - BorderSize - EdgeSpace
    tR.Right = .ScaleWidth - BorderSize - EdgeSpace
    If m_BackColorStyle = aGradientColor And m_Transparent = False Then DrawGradientBackColor tPic, m_GradientBackColorStyle
    If m_Shadow Then
        DrawShadow tPic, s, tR, m_WordWrap, m_ShadowColorStart, ShadowColorEnd, m_ShadowDepth, m_ShadowStyle
    End If
    tPic.ForeColor = m_ForeColor
    m_LineCount = DrawText(tPic.hdc, s, tR, m_WordWrap) / tPic.TextHeight(Left(s, 1))
    Select Case m_BorderStyle
        Case 0: 'None
        Case 1: 'Flat
            tPic.ForeColor = m_BorderColor
            For i = 1 To BorderSize
                Rectangle tPic.hdc, i - 1, i - 1, .ScaleWidth - i + 1, .ScaleHeight - i + 1
            Next i
            'tPic.ForeColor = m_ForeColor
        Case 2: 'Outline
            tPic.ForeColor = m_OutlineColor
            DrawOutline tPic.hdc, s, tR, m_WordWrap, BorderSize
            tPic.ForeColor = m_ForeColor
            DrawText tPic.hdc, s, tR, m_WordWrap
        Case 3, 4, 5: '3D,Frame1,Frame2
            tPic.ForeColor = m_ForeColor
            m_BorderSize = IIf(m_BorderStyle = 5, 3, 2)
            PropertyChanged "BorderSize"
            'DrawText tPic.Hdc, s, tR, m_WordWrap
            If m_BorderStyle = 3 Then
                tPic.Line (0, 0)-(.ScaleWidth - 1, 0), RGB(167, 166, 170)
                tPic.Line (0, 0)-(0, .ScaleHeight - 1), RGB(167, 166, 170)
                tPic.Line (.ScaleWidth - 1, 0)-(.ScaleWidth - 1, .ScaleHeight), RGB(255, 255, 255)
                tPic.Line (0, .ScaleHeight - 1)-(.ScaleWidth, .ScaleHeight - 1), RGB(255, 255, 255)
    
                tPic.Line (1, 1)-(.ScaleWidth - 2, 1), RGB(133, 135, 140)
                tPic.Line (1, 1)-(1, .ScaleHeight - 2), RGB(133, 135, 140)
                tPic.Line (.ScaleWidth - 2, 1)-(.ScaleWidth - 2, .ScaleHeight - 1), RGB(220, 223, 228)
                tPic.Line (1, .ScaleHeight - 2)-(.ScaleWidth - 1, .ScaleHeight - 2), RGB(220, 223, 228)
            ElseIf m_BorderStyle = 4 Then
                tPic.Line (0, 0)-(.ScaleWidth - 1, 0), RGB(167, 166, 170)
                tPic.Line (0, 0)-(0, .ScaleHeight - 1), RGB(167, 166, 170)
                tPic.Line (.ScaleWidth - 1, 0)-(.ScaleWidth - 1, .ScaleHeight), RGB(255, 255, 255)
                tPic.Line (0, .ScaleHeight - 1)-(.ScaleWidth, .ScaleHeight - 1), RGB(255, 255, 255)
    
                tPic.Line (1, 1)-(.ScaleWidth - 2, 1), RGB(255, 255, 255)
                tPic.Line (1, 1)-(1, .ScaleHeight - 2), RGB(255, 255, 255)
                tPic.Line (.ScaleWidth - 2, 1)-(.ScaleWidth - 2, .ScaleHeight - 1), RGB(167, 166, 170)
                tPic.Line (1, .ScaleHeight - 2)-(.ScaleWidth - 1, .ScaleHeight - 2), RGB(167, 166, 170)
            Else
                tPic.ForeColor = RGB(255, 255, 255)
                Rectangle tPic.hdc, 0, 0, .ScaleWidth, .ScaleHeight
                Rectangle tPic.hdc, 2, 2, .ScaleWidth - 2, .ScaleHeight - 2
                tPic.ForeColor = RGB(167, 166, 170)
                Rectangle tPic.hdc, 1, 1, .ScaleWidth - 1, .ScaleHeight - 1
            End If
    End Select
    Set .Picture = tPic.Image
    If m_Transparent Then
        .BackStyle = 0
        .MaskColor = TransColor
        .MaskPicture = tPic.Image
    Else
        .BackStyle = 1
    End If
End With
End Sub

Private Function DrawText(hdc As Long, s As String, tRect As RECT, wwrap As Boolean) As Long
    Dim Flags As Long
    Flags = DT_NOCLIP
    Flags = Flags Or IIf(wwrap, DT_WORDBREAK, 0)
    Flags = Flags Or IIf(m_Alignment = 0, 0, IIf(m_Alignment = 1, DT_RIGHT, DT_CENTER))
    DrawText = DrawTextW(hdc, StrPtr(s), Len(s), tRect, Flags)
End Function

Private Sub DrawOutline(hdc As Long, s As String, tRect As RECT, wwrap As Boolean, Size As Integer)
    Dim tx As Integer, ty As Integer
    Dim t As RECT
    t = tRect
    For ty = tRect.Top - Size To tRect.Top + Size
        For tx = tRect.Left - Size To tRect.Left + Size
            t.Right = t.Right - (t.Left - tx)
            t.Bottom = t.Bottom - (t.Top - ty)
            t.Left = tx
            t.Top = ty
            DrawText hdc, s, t, wwrap
        Next tx
    Next ty
End Sub

Private Sub DrawShadow(P As PictureBox, s As String, tR As RECT, wwrap As Boolean, Color_S As Long, Color_E As Long, Depth As Integer, Style As E_ShadowStyle)
Dim AColor() As Long
Dim t As RECT
Dim i As Integer, dx As Integer, dy As Integer
    t = tR
    GradientColor Color_E, Color_S, Depth + 1, AColor
    Select Case Style
        Case 0: dx = -1: dy = -1 'LeftTop
        Case 1: dx = 1: dy = -1 'RightTop
        Case 2: dx = -1: dy = 1 'LeftBottom
        Case 3: dx = 1: dy = 1  'RightBottom
    End Select
    For i = Depth To 1 Step -1
        P.ForeColor = AColor(Depth - i)
        t.Left = tR.Left + i * dx
        t.Top = tR.Top + i * dy
        t.Right = tR.Right + i * dx
        t.Bottom = tR.Bottom + i * dy
        DrawText P.hdc, s, t, wwrap
    Next i
End Sub

Private Sub DrawGradientBackColor(t As Object, Style As E_GradientBackColorStyle)
Dim c() As Long
Dim Depth As Integer
Dim i As Integer
With t
    '(IIf(i < .ScaleHeight, 0, i - .ScaleHeight + 1), IIf(i < .ScaleHeight, i, .ScaleHeight - 1))-(IIf(i < .ScaleWidth, i + 1, .ScaleWidth), IIf(i < .ScaleWidth, -1, i - .ScaleWidth)), c(i)
    If Style = aLeftToRight Or Style = aRightToLeft Then
        Depth = .ScaleWidth
    ElseIf Style = aTopToBottom Or Style = aBottomToTop Then
        Depth = .ScaleHeight
    ElseIf Style = aLeftTopToRightBottom Or Style = aLeftBottomToRightTop Then
        Depth = 2 * .ScaleWidth - 1
    ElseIf Style = aCenterToLeftRight Then
        Depth = (.ScaleWidth + 1) \ 2
    ElseIf Style = aCenterToTopBottom Then
        Depth = (.ScaleHeight + 1) \ 2
    ElseIf Style = aCenterToLeftTopNRightBottom Or aCenterToLeftBottomNRightTop Then
        Depth = .ScaleWidth
    End If
    GradientColor m_GradientBackColorStart, m_GradientBackColorEnd, Depth, c
    Select Case Style
        Case aLeftToRight:
            For i = 0 To .ScaleWidth - 1
                t.Line (i, 0)-(i, .ScaleHeight), c(i)
            Next i
        Case aRightToLeft:
            For i = 0 To .ScaleWidth - 1
                t.Line (i, 0)-(i, .ScaleHeight), c(.ScaleWidth - i - 1)
            Next i
        Case aTopToBottom:
            For i = 0 To .ScaleHeight - 1
                t.Line (0, i)-(.ScaleWidth, i), c(i)
            Next i
        Case aBottomToTop:
            For i = 0 To .ScaleHeight - 1
                t.Line (0, i)-(.ScaleWidth, i), c(.ScaleHeight - i - 1)
            Next i
        Case aLeftTopToRightBottom:
            'For i = 0 To .ScaleWidth + .ScaleHeight - 2
            '    t.Line (IIf(i < .ScaleHeight, 0, i - .ScaleHeight + 1), IIf(i < .ScaleHeight, i, .ScaleHeight - 1))-(IIf(i < .ScaleWidth, i + 1, .ScaleWidth), IIf(i < .ScaleWidth, -1, i - .ScaleWidth)), c(i)
            'Next i
            t.PSet (0, 0), c(0)
            t.PSet (.ScaleWidth - 1, .ScaleHeight - 1), c(.ScaleWidth + .ScaleWidth - 2)
            For i = 1 To .ScaleWidth + .ScaleWidth - 3
                t.Line (i - .ScaleWidth + 1, .ScaleHeight - 1)-(i + 1, -1), c(i)
            Next i
        Case aLeftBottomToRightTop:
            t.PSet (0, .ScaleHeight - 1), c(0)
            t.PSet (.ScaleWidth - 1, 0), c(.ScaleWidth + .ScaleWidth - 2)
            For i = 1 To .ScaleWidth + .ScaleWidth - 3
                t.Line (i - .ScaleWidth + 1, -1)-(i + 1, .ScaleHeight - 1), c(i)
            Next i
        Case aCenterToLeftRight:
            For i = 0 To .ScaleWidth \ 2 + (.ScaleWidth Mod 2) - 1
                t.Line (i, 0)-(i, .ScaleHeight), c(.ScaleWidth \ 2 + (.ScaleWidth Mod 2) - 1 - i)
                t.Line (i + .ScaleWidth \ 2, 0)-(i + .ScaleWidth \ 2, .ScaleHeight), c(i)
            Next i
        Case aCenterToTopBottom:
            For i = 0 To .ScaleHeight \ 2 + (.ScaleHeight Mod 2) - 1
                t.Line (0, i)-(.ScaleWidth, i), c(.ScaleHeight \ 2 + (.ScaleHeight Mod 2) - 1 - i)
                t.Line (0, i + .ScaleHeight \ 2)-(.ScaleWidth, i + .ScaleHeight \ 2), c(i)
            Next i
        Case aCenterToLeftTopNRightBottom:
            For i = 0 To Depth
                t.Line (i - .ScaleWidth + 1, .ScaleHeight - 1)-(i + 1, -1), c(Depth - i)
                t.Line (i, .ScaleHeight - 1)-(i + .ScaleWidth, -1), c(i)
            Next i
        Case aCenterToLeftBottomNRightTop
            For i = 0 To Depth
                t.Line (i - .ScaleWidth + 1, 0)-(i + 1, .ScaleHeight), c(Depth - i)
                t.Line (i, 0)-(i + .ScaleWidth, .ScaleHeight), c(i)
            Next i
    End Select
End With
End Sub
