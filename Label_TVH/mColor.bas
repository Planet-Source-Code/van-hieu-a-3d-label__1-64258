Attribute VB_Name = "mColor"

Option Explicit

Public Sub GradientColor(Color1 As Long, Color2 As Long, Depth As Integer, Result() As Long)
Dim VR, VG, VB As Single
Dim R, G, B, R2, G2, B2 As Integer
Dim t As Long
    t = (Color1 And 255)
    R = t And 255
    t = Int(Color1 / 256)
    G = t And 255
    t = Int(Color1 / 65536)
    B = t And 255
    t = (Color2 And 255)
    R2 = t And 255
    t = Int(Color2 / 256)
    G2 = t And 255
    t = Int(Color2 / 65536)
    B2 = t And 255
    VR = Abs(R - R2) / Depth
    VG = Abs(G - G2) / Depth
    VB = Abs(B - B2) / Depth
    If R2 < R Then VR = -VR
    If G2 < G Then VG = -VG
    If B2 < B Then VB = -VB
    ReDim Result(Depth)
    For t = 0 To Depth
        R2 = R + VR * t
        G2 = G + VG * t
        B2 = B + VB * t
        Result(t) = RGB(R2, G2, B2)
    Next t
End Sub
