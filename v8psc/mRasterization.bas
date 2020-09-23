Attribute VB_Name = "mRasterization"
Option Explicit

' =====================================================
' This module is responsible for drawing to the screen.
' =====================================================
Private Const mc_Name As String = "Peters3DEngine8.mRasterization"


' The SetPixel function sets the pixel at the specified coordinates to the specified color.
Public Declare Function SetPixel Lib "GDI32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long) As Long


' API Declarations used for drawing and filling "flat-shaded" triangles.
Public Type POINT_TYPE
  X As Long
  Y As Long
End Type
Public Declare Function Polygon Lib "gdi32.dll" (ByVal hdc As Long, lpPoint As POINT_TYPE, ByVal nCount As Long) As Long


' API Declarations used for drawing "Gouraud Shaded" triangles.
' (Can also be used for simple "flat shaded" triangles too.)
Public Type TRIVERTEX
    X As Long
    Y As Long
    Red As Integer
    Green As Integer
    Blue As Integer
    Alpha As Integer
End Type
Public Type GRADIENT_TRIANGLE
    Vertex1 As Long
    Vertex2 As Long
    Vertex3 As Long
End Type
Public Const GRADIENT_FILL_TRIANGLE As Long = &H2
Public Declare Function GradientFillTri Lib "msimg32" Alias "GradientFill" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_TRIANGLE, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Public sX As Double, sY As Double
Public vert(3) As TRIVERTEX
Public gTri(1) As GRADIENT_TRIANGLE

Private Function convert_Long2UShort(ULong As Long) As Integer

    ' This function converts a long integer to an unsigned integer (ie. a UShort)
    '
    ' All Visual Basic integers are "signed" integers, meaning they go from -32,768 to 32,767
    ' However our API routine needs an "unsigned" integer, meaning it goes from 0 to 65534.
    ' Both signed, and unsigned integers take up 16 bits. ie. 0000000000000000
    
    If ULong <= &H7FFF& Then
        convert_Long2UShort = ULong
    Else
        convert_Long2UShort = Not (&HFFFF& - ULong)
    End If
    
    
    ' Update debug & info. counters (conditionally compiled)
    #If g_blnCompiledDebugInfo = True Then
        g_lngSubtractionCount = g_lngSubtractionCount + 1
    #End If
    
End Function

Public Sub DrawTriangle(withHandle As Form, x1 As Single, y1 As Single, x2 As Single, y2 As Single, x3 As Single, y3 As Single, ByVal Red1 As Long, ByVal Green1 As Long, ByVal Blue1 As Long, ByVal Red2 As Long, ByVal Green2 As Long, ByVal Blue2 As Long, ByVal Red3 As Long, ByVal Green3 As Long, ByVal Blue3 As Long)

    ' Do basic error checking.
    If withHandle Is Nothing Then Exit Sub
    If withHandle.HasDC = False Then Exit Sub
            
    ' Clamp red light to safe values.
    ' ===============================
    If Red1 > 255 Then Red1 = 255
    If Red2 > 255 Then Red2 = 255
    If Red3 > 255 Then Red3 = 255
    If Red1 < 0 Then Red1 = 0
    If Red2 < 0 Then Red2 = 0
    If Red3 < 0 Then Red3 = 0
    
    ' Clamp green light to safe values.
    ' =================================
    If Green1 > 255 Then Green1 = 255
    If Green2 > 255 Then Green2 = 255
    If Green3 > 255 Then Green3 = 255
    If Green1 < 0 Then Green1 = 0
    If Green2 < 0 Then Green2 = 0
    If Green3 < 0 Then Green3 = 0
    
    ' Clamp blue light to safe values.
    ' ================================
    If Blue1 > 255 Then Blue1 = 255
    If Blue2 > 255 Then Blue2 = 255
    If Blue3 > 255 Then Blue3 = 255
    If Blue1 < 0 Then Blue1 = 0
    If Blue2 < 0 Then Blue2 = 0
    If Blue3 < 0 Then Blue3 = 0
    
    
    With vert(0)
        .X = x1
        .Y = y1
        .Red = convert_Long2UShort(Red1& * 256&)
        .Green = convert_Long2UShort(Green1& * 256&)
        .Blue = convert_Long2UShort(Blue1& * 256&)
        ' Note that GradientFill does not use the Alpha member of the TRIVERTEX structure.
    End With
    
    With vert(1)
        .X = x2
        .Y = y2
        .Red = convert_Long2UShort(Red2& * 256&)
        .Green = convert_Long2UShort(Green2& * 256&)
        .Blue = convert_Long2UShort(Blue2& * 256&)
        ' Note that GradientFill does not use the Alpha member of the TRIVERTEX structure.
    End With
    
    With vert(2)
        .X = x3
        .Y = y3
        .Red = convert_Long2UShort(Red3& * 256&)
        .Green = convert_Long2UShort(Green3& * 256&)
        .Blue = convert_Long2UShort(Blue3& * 256&)
        ' Note that GradientFill does not use the Alpha member of the TRIVERTEX structure.
    End With
    
    gTri(0).Vertex1 = 0
    gTri(0).Vertex2 = 1
    gTri(0).Vertex3 = 2
    
    GradientFillTri withHandle.hdc, vert(0), 3, gTri(0), 1, GRADIENT_FILL_TRIANGLE
    
End Sub

Public Sub DrawFlatTriangle(withHandle As Form, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single, lngDrawColour As Long, lngFillColour As Long, intDrawStyle As DrawStyleConstants, intFillStyle As FillStyleConstants)
    
    ' Do basic error checking.
    If withHandle Is Nothing Then Exit Sub
    If withHandle.HasDC = False Then Exit Sub
    
    Dim points(0 To 3) As POINT_TYPE
    
    points(0).X = x1: points(0).Y = y1
    points(1).X = x2: points(1).Y = y2
    points(2).X = x3: points(2).Y = y3
    points(3) = points(0)
    
    ' Fill Options
    ' =============
    If lngFillColour = -1 Then ' Turn off fill.
        withHandle.FillStyle = vbFSTransparent
    Else
        ' Fill polygon with specified colour.
        withHandle.FillColor = lngFillColour
        withHandle.FillStyle = intFillStyle
    End If
    
    ' Draw Options
    ' ============
    If lngDrawColour = -1 Then ' Turn off edges
        withHandle.DrawStyle = vbInvisible
    Else
        withHandle.ForeColor = lngDrawColour
        withHandle.DrawStyle = intDrawStyle
    End If
    
    ' Call the API to Draw and Fill the polygon.
    Call Polygon(withHandle.hdc, points(0), 4)
    
    
End Sub

Public Sub DrawFlatTriangle2(withHandle As Form, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single, lngDrawColour As Long, lngFillColour As Long, intDrawStyle As DrawStyleConstants, intFillStyle As FillStyleConstants)

    ' Do basic error checking.
    If withHandle Is Nothing Then Exit Sub
    If withHandle.HasDC = False Then Exit Sub

    ' Fill Options
    ' =============
    If lngFillColour = -1 Then ' Turn off fill.
        withHandle.FillStyle = vbFSTransparent
    Else
        ' Fill polygon with specified colour.
        withHandle.FillColor = lngFillColour
        withHandle.FillStyle = intFillStyle
    End If
    
    ' Draw Options
    ' ============
    If lngDrawColour = -1 Then ' Turn off edges
        withHandle.DrawStyle = vbInvisible
    Else
        withHandle.ForeColor = lngDrawColour
        withHandle.DrawStyle = intDrawStyle
    End If
    
    withHandle.Line (x1, y1)-(x2, y2)
    withHandle.Line -(x3, y3)
    withHandle.Line -(x1, y1)
    
End Sub


Public Sub DrawFlatTriangle3(withHandle As PictureBox, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single)

    ' Do basic error checking.
    If withHandle Is Nothing Then Exit Sub
    If withHandle.HasDC = False Then Exit Sub

    withHandle.Line (x1, y1)-(x2, y2)
    withHandle.Line -(x3, y3)
    withHandle.Line -(x1, y1)
    
End Sub

Public Sub RenderCrossHairs(withHandle As Form, p_Camera As mdr3DTargetCamera, CrossHairStyle As eCrossHairStyle)

    ' ===================================================================
    ' This routine is written so that cross-hair style's can be combined.
    ' ===================================================================

    If CrossHairStyle = mdrCRS_None Then Exit Sub
    
    Dim sngCenterX As Single
    Dim sngCenterY As Single
    With p_Camera
        sngCenterX = ((.VPXmax - .VPXmin) / 2) + .VPXmin
        sngCenterY = ((.VPYmin - .VPYmax) / 2) + .VPYmax
    End With
    
    withHandle.DrawStyle = vbSolid
    withHandle.DrawMode = vbCopyPen
    withHandle.ForeColor = RGB(128, 128, 128)
    withHandle.FillStyle = vbFSTransparent
    withHandle.DrawWidth = 1
    
    
    If (CrossHairStyle And mdrCRS_DotOnly) = mdrCRS_DotOnly Then
        withHandle.PSet (sngCenterX, sngCenterY)
    End If


    If (CrossHairStyle And mdrCRS_Cross) = mdrCRS_Cross Then
        withHandle.Line (sngCenterX - 16, sngCenterY)-(sngCenterX - 8, sngCenterY)
        withHandle.Line (sngCenterX + 16, sngCenterY)-(sngCenterX + 8, sngCenterY)
        withHandle.Line (sngCenterX, sngCenterY - 16)-(sngCenterX, sngCenterY - 8)
        withHandle.Line (sngCenterX, sngCenterY + 16)-(sngCenterX, sngCenterY + 8)
    End If


    If (CrossHairStyle And mdrCRS_X) = mdrCRS_X Then
        withHandle.Line (sngCenterX - 8, sngCenterY - 8)-(sngCenterX - 4, sngCenterY - 4)
        withHandle.Line (sngCenterX + 8, sngCenterY - 8)-(sngCenterX + 4, sngCenterY - 4)
        withHandle.Line (sngCenterX - 8, sngCenterY + 8)-(sngCenterX - 4, sngCenterY + 4)
        withHandle.Line (sngCenterX + 8, sngCenterY + 8)-(sngCenterX + 4, sngCenterY + 4)
    End If
    
    
    If (CrossHairStyle And mdrCRS_Scope1) = mdrCRS_Scope1 Then
    
        Dim sngDeg As Single, sngRad As Single
        Dim sngX1 As Single, sngX2 As Single
        Dim sngY1 As Single, sngY2 As Single
        
        Static sngTemp As Single
        Static sngTemp2 As Single
        sngTemp = sngTemp + 0
        If sngTemp > 360 Then sngTemp = 0

        Dim sngRadius As Single
        sngRadius = 64
        
        withHandle.Circle (sngCenterX, sngCenterY), sngRadius + 8, RGB(128, 128, 128)

        withHandle.DrawWidth = 3
        For sngDeg = 0 + sngTemp2 To 360 + sngTemp2 Step (360 / 4)
            sngRad = ConvertDeg2Rad(sngDeg)
            sngX1 = Cos(sngRad) * sngRadius + sngCenterX
            sngX2 = Cos(sngRad) * (sngRadius + 16) + sngCenterX
            sngY1 = Sin(sngRad) * sngRadius + sngCenterY
            sngY2 = Sin(sngRad) * (sngRadius + 16) + sngCenterY
            withHandle.Line (sngX1, sngY1)-(sngX2, sngY2)
        Next sngDeg
    
    
        withHandle.DrawWidth = 2
        
        For sngDeg = 0 - sngTemp To 360 - sngTemp Step (360 / 8)
            sngRad = ConvertDeg2Rad(sngDeg)
            sngX1 = Cos(sngRad) * sngRadius + sngCenterX
            sngX2 = Cos(sngRad) * (sngRadius + 8) + sngCenterX
            sngY1 = Sin(sngRad) * sngRadius + sngCenterY
            sngY2 = Sin(sngRad) * (sngRadius + 8) + sngCenterY
            withHandle.Line (sngX1, sngY1)-(sngX2, sngY2)
        Next sngDeg
    
    End If

End Sub

