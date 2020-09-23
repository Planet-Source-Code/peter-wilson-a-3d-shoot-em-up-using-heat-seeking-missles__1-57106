Attribute VB_Name = "mMouse"
Option Explicit

' =========================================================================================
' 3D Computer Graphics for Visual Basic Programmers: Theory, Practice, Source Code and Fun!
' Version: 5.0 - Game Edition.
'
' by Peter Wilson
' Copyright © 2003 - Peter Wilson - All rights reserved.
' http://dev.midar.com/
' =========================================================================================

' Define the name of this class/module for error-trap reporting.
Private Const m_strModuleName As String = "mMouse"


' API Declarations used to GET & SET the position of the mouse.
Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Attribute GetCursorPos.VB_Description = "GetCursorPos reads the current position of the mouse cursor."
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Attribute SetCursorPos.VB_Description = "SetCursorPos sets the current position of the mouse cursor."


' The ShowCursor function displays or hides the cursor.
' (This function sets an internal display counter that determines whether the cursor should be displayed.
'  The cursor is displayed only if the display count is greater than or equal to 0.
'  If a mouse is installed, the initial display count is 0. If no mouse is installed, the display count is –1.)
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Attribute ShowCursor.VB_Description = "The ShowCursor function displays or hides the cursor."

Public Sub GetMouseInput(intMouseSensitivity As Integer, strGameState As String, objPlayer As mdrPlayer, p_Camera As mdr3DTargetCamera, withHandle As Form)

    ' ===================================================================
    ' Capture the current mouse position using an API.
    ' Don't forget: API's use Pixels NOT Twips (Twips are the VB default)
    ' ===================================================================
    Dim lngRetVal As Long
    Dim myPoint As POINTAPI
    Static deltaX As Integer
    Static deltaY As Integer
    Dim vectA As mdrVector4
    
    
    Dim sngCenterX As Single
    Dim sngCenterY As Single
    With p_Camera
        sngCenterX = ((.VPXmax - .VPXmin) / 2) + .VPXmin
        sngCenterY = ((.VPYmin - .VPYmax) / 2) + .VPYmax
    End With
    
    ' Get the current position of the mouse.
    lngRetVal = GetCursorPos(myPoint)
    
    deltaX = (myPoint.X - sngCenterX)
    deltaY = (sngCenterY - myPoint.Y)
    
    Select Case strGameState
        Case ""
            ' Reset the rotation vector to a unit-vector (ie. 1 unit in length).
            objPlayer.VPN.X = 0
            objPlayer.VPN.Y = 0
            objPlayer.VPN.Z = 1
            objPlayer.VPN.w = 1
    
            vectA.X = 0
            vectA.Y = 1
            vectA.Z = 0
            vectA.w = 1
            objPlayer.LeftRightVector = VectorCrossProduct(vectA, objPlayer.VPN)
            
            ' Reset the mouse to the middle of the form. This is my lazy way to prevent large
            ' unwanted values the first time the user moves the mouse on purpose.
            lngRetVal = SetCursorPos(sngCenterX, sngCenterY)
            
        Case "run_game"
            
            ' ======================================
            ' Apply mouse movement to player object.
            ' This is a 2 step process.
            ' 1) Increment rotation offset.
            ' 2) Apply the offset.
            ' --------------------------------------
            objPlayer.XYPlane = objPlayer.XYPlane - (deltaX / intMouseSensitivity)
            objPlayer.XZPlane = objPlayer.XZPlane + (deltaY / intMouseSensitivity)
            
            ' Clip the up/down look values (the XZ plane) to prevent player from looking more than
            ' 90 degrees up, otherwise the screen will snap around violently confusing the player.
            ' =====================================================================================
            If objPlayer.XYPlane > 360 Then objPlayer.XYPlane = objPlayer.XYPlane - 360
            If objPlayer.XYPlane < 0 Then objPlayer.XYPlane = objPlayer.XYPlane + 360
            If objPlayer.XZPlane > 89 Then objPlayer.XZPlane = 89
            If objPlayer.XZPlane < -89 Then objPlayer.XZPlane = -89
            

            ' Reset the rotation vector to a unit-vector (ie. 1 unit in length).
            objPlayer.VPN.X = 0
            objPlayer.VPN.Y = 0
            objPlayer.VPN.Z = 1
            objPlayer.VPN.w = 1
    
            Dim matXY As mdrMatrix4
            Dim matXZ As mdrMatrix4
            Dim matResult As mdrMatrix4
            
            ' Rotate this unit-vector.
            matXY = MatrixRotationY(ConvertDeg2Rad(objPlayer.XYPlane))
            matXZ = MatrixRotationX(ConvertDeg2Rad(objPlayer.XZPlane))
            matResult = MatrixMultiply(matXZ, matXY)
            objPlayer.VPN = MatrixMultiplyVector(matResult, objPlayer.VPN)
            
            ' Find the cross-product of the resultant.
            vectA.X = 0
            vectA.Y = 1
            vectA.Z = 0
            vectA.w = 1
            objPlayer.LeftRightVector = VectorCrossProduct(vectA, objPlayer.VPN)
            
            
            ' Move the mouse position back to the middle, so that the user can continue scrolling in a single
            ' direction without the problem of the mouse stopping because it hits the boundry of the screen.
            ' Yes... this is what many professional games do.
'            lngRetVal = SetCursorPos(sngCenterX, sngCenterY)
            ' Hint: Remark-out the above line for some really funky mouse-action!!
                        
    End Select
    
End Sub
