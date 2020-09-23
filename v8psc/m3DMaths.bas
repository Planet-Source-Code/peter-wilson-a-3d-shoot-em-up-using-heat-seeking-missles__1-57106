Attribute VB_Name = "m3DMaths"
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
Private Const m_strModuleName As String = "m3DMaths"

' =========================================================================================
' Define a few constants.
' =========================================================================================
Public Const g_sngPI As Double = 3.14159265358979
Private Const g_sngPIDivideBy180 As Double = 1.74532925199433E-02
Private Const g_sng180DivideByPI As Double = 57.2957795130823

Public Function MatrixDeterminant(m As mdrMatrix4) As Single

    ' If a matrix can     be inverted, then it will have a non-zero determinant.
    ' If a matrix can not be inverted, then it will have a     zero determinant.

    With m
    
        MatrixDeterminant = _
            (.rc14 * .rc23 * .rc32 * .rc41) - (.rc13 * .rc24 * .rc32 * .rc41) - (.rc14 * .rc22 * .rc33 * .rc41) + (.rc12 * .rc24 * .rc33 * .rc41) + _
            (.rc13 * .rc22 * .rc34 * .rc41) - (.rc12 * .rc23 * .rc34 * .rc41) - (.rc14 * .rc23 * .rc31 * .rc42) + (.rc13 * .rc24 * .rc31 * .rc42) + _
            (.rc14 * .rc21 * .rc33 * .rc42) - (.rc11 * .rc24 * .rc33 * .rc42) - (.rc13 * .rc21 * .rc34 * .rc42) + (.rc11 * .rc23 * .rc34 * .rc42) + _
            (.rc14 * .rc22 * .rc31 * .rc43) - (.rc12 * .rc24 * .rc31 * .rc43) - (.rc14 * .rc21 * .rc32 * .rc43) + (.rc11 * .rc24 * .rc32 * .rc43) + _
            (.rc12 * .rc21 * .rc34 * .rc43) - (.rc11 * .rc22 * .rc34 * .rc43) - (.rc13 * .rc22 * .rc31 * .rc44) + (.rc12 * .rc23 * .rc31 * .rc44) + _
            (.rc13 * .rc21 * .rc32 * .rc44) - (.rc11 * .rc23 * .rc32 * .rc44) - (.rc12 * .rc21 * .rc33 * .rc44) + (.rc11 * .rc22 * .rc33 * .rc44)
            
    End With
    
    
    ' Update debug & info. counters (conditionally compiled)
    #If g_blnCompiledDebugInfo = True Then
        g_lngMultiplicationCount = g_lngMultiplicationCount + 72
        g_lngSubtractionCount = g_lngSubtractionCount + 12
        g_lngAdditionCount = g_lngAdditionCount + 11
    #End If
    
End Function

Public Function MatrixInverse(m As mdrMatrix4) As mdrMatrix4
    
    ' There are several ways to calculate the inverse of a matrix.
    ' This method is based on Cramers Rule. I don't fully understand it yet, but it is quite popular.
    ' The only downside to this method, is the number of calculations required. There are several easy methods to
    ' calculate the inverse of an independant transform (ie. scale, rotate and translate) which do not require
    ' complex calculations. However, if you don't know how the matrix was created, then you'll have to determine
    ' it's inverse the long, slow way. Sometimes, a "starting matrix" for an object is automatically imported and
    ' if you want to find it's inverse, you'll need to use this function.
    '
    ' Also see:
    '   http://www.euclideanspace.com/maths/algebra/matrix/functions/inverse/index.htm
    '
    ' Most of the time, you will only use an inverted matrix when you are trying to optomise your 3D code.
    ' If you don't care about optomizing code, then you'll never have to invert a matrix. They are optional.
    '
    ' Most of the time, you will use an inverted matrix to help speed up Back-Face-Cull or lighting calculations.
    ' Slow Way:
    '   Transform 10,000 faces into the Camera coordinate system,
    '   and then test these 10,000 faces against the Virtual Camera
    '   to see if they are backfacing.
    '
    ' Fast Way:
    '   Transform the Camera into the Object's coordinate system (using an inverse matrix),
    '   and then test against 10,000 faces to see if they are backfacing.
    '   This will save you 10,000 calculations.
    
    With m
        MatrixInverse.rc11 = (.rc23 * .rc34 * .rc42) - (.rc24 * .rc33 * .rc42) + (.rc24 * .rc32 * .rc43) - (.rc22 * .rc34 * .rc43) - (.rc23 * .rc32 * .rc44) + (.rc22 * .rc33 * .rc44)
        MatrixInverse.rc12 = (.rc14 * .rc33 * .rc42) - (.rc13 * .rc34 * .rc42) - (.rc14 * .rc32 * .rc43) + (.rc12 * .rc34 * .rc43) + (.rc13 * .rc32 * .rc44) - (.rc12 * .rc33 * .rc44)
        MatrixInverse.rc13 = (.rc13 * .rc24 * .rc42) - (.rc14 * .rc23 * .rc42) + (.rc14 * .rc22 * .rc43) - (.rc12 * .rc24 * .rc43) - (.rc13 * .rc22 * .rc44) + (.rc12 * .rc23 * .rc44)
        MatrixInverse.rc14 = (.rc14 * .rc23 * .rc32) - (.rc13 * .rc24 * .rc32) - (.rc14 * .rc22 * .rc33) + (.rc12 * .rc24 * .rc33) + (.rc13 * .rc22 * .rc34) - (.rc12 * .rc23 * .rc34)
        
        MatrixInverse.rc21 = (.rc24 * .rc33 * .rc41) - (.rc23 * .rc34 * .rc41) - (.rc24 * .rc31 * .rc43) + (.rc21 * .rc34 * .rc43) + (.rc23 * .rc31 * .rc44) - (.rc21 * .rc33 * .rc44)
        MatrixInverse.rc22 = (.rc13 * .rc34 * .rc41) - (.rc14 * .rc33 * .rc41) + (.rc14 * .rc31 * .rc43) - (.rc11 * .rc34 * .rc43) - (.rc13 * .rc31 * .rc44) + (.rc11 * .rc33 * .rc44)
        MatrixInverse.rc23 = (.rc14 * .rc23 * .rc41) - (.rc13 * .rc24 * .rc41) - (.rc14 * .rc21 * .rc43) + (.rc11 * .rc24 * .rc43) + (.rc13 * .rc21 * .rc44) - (.rc11 * .rc23 * .rc44)
        MatrixInverse.rc24 = (.rc13 * .rc24 * .rc31) - (.rc14 * .rc23 * .rc31) + (.rc14 * .rc21 * .rc33) - (.rc11 * .rc24 * .rc33) - (.rc13 * .rc21 * .rc34) + (.rc11 * .rc23 * .rc34)
        
        MatrixInverse.rc31 = (.rc22 * .rc34 * .rc41) - (.rc24 * .rc32 * .rc41) + (.rc24 * .rc31 * .rc42) - (.rc21 * .rc34 * .rc42) - (.rc22 * .rc31 * .rc44) + (.rc21 * .rc32 * .rc44)
        MatrixInverse.rc32 = (.rc14 * .rc32 * .rc41) - (.rc12 * .rc34 * .rc41) - (.rc14 * .rc31 * .rc42) + (.rc11 * .rc34 * .rc42) + (.rc12 * .rc31 * .rc44) - (.rc11 * .rc32 * .rc44)
        MatrixInverse.rc33 = (.rc12 * .rc24 * .rc41) - (.rc14 * .rc22 * .rc41) + (.rc14 * .rc21 * .rc42) - (.rc11 * .rc24 * .rc42) - (.rc12 * .rc21 * .rc44) + (.rc11 * .rc22 * .rc44)
        MatrixInverse.rc34 = (.rc14 * .rc22 * .rc31) - (.rc12 * .rc24 * .rc31) - (.rc14 * .rc21 * .rc32) + (.rc11 * .rc24 * .rc32) + (.rc12 * .rc21 * .rc34) - (.rc11 * .rc22 * .rc34)
        
        MatrixInverse.rc41 = (.rc23 * .rc32 * .rc41) - (.rc22 * .rc33 * .rc41) - (.rc23 * .rc31 * .rc42) + (.rc21 * .rc33 * .rc42) + (.rc22 * .rc31 * .rc43) - (.rc21 * .rc32 * .rc43)
        MatrixInverse.rc42 = (.rc12 * .rc33 * .rc41) - (.rc13 * .rc32 * .rc41) + (.rc13 * .rc31 * .rc42) - (.rc11 * .rc33 * .rc42) - (.rc12 * .rc31 * .rc43) + (.rc11 * .rc32 * .rc43)
        MatrixInverse.rc43 = (.rc13 * .rc22 * .rc41) - (.rc12 * .rc23 * .rc41) - (.rc13 * .rc21 * .rc42) + (.rc11 * .rc23 * .rc42) + (.rc12 * .rc21 * .rc43) - (.rc11 * .rc22 * .rc43)
        MatrixInverse.rc44 = (.rc12 * .rc23 * .rc31) - (.rc13 * .rc22 * .rc31) + (.rc13 * .rc21 * .rc32) - (.rc11 * .rc23 * .rc32) - (.rc12 * .rc21 * .rc33) + (.rc11 * .rc22 * .rc33)
    End With
    
    Dim sngDet As Single
    sngDet = MatrixDeterminant(m)
    
    Dim matScale As mdrMatrix4
    matScale = MatrixScale(1 / sngDet, 1 / sngDet, 1 / sngDet, 1 / sngDet)
    
    MatrixInverse = MatrixMultiply(MatrixInverse, matScale)
    
    
    ' Update debug & info. counters (conditionally compiled)
    #If g_blnCompiledDebugInfo = True Then
        g_lngMultiplicationCount = g_lngMultiplicationCount + 192
        g_lngDivisionCount = g_lngDivisionCount + 4
        g_lngSubtractionCount = g_lngSubtractionCount + 48
        g_lngAdditionCount = g_lngAdditionCount + 35
    #End If
    
End Function

Public Function MatrixLookAt(p_Origin As mdrVector4, p_Target As mdrVector4) As mdrMatrix4
    
    ' ======================================================================================
    '                       *** Synthetic [Virtual] Camera Values ***
    ' The four basic Viewing Parameters that define the View Reference Coordinates (VRC).
    ' They are all specified in World Coordinates (WC) except PRP; this is specified in VRC.
    '                       *** Synthetic [Virtual] Camera Values ***
    '
    ' ======================================================================================
    
    Dim vectVRP As mdrVector4   ' View Reference Point (VRP) - The world position of the virtual camera AND the virtual film!
    Dim vectVPN As mdrVector4   ' View Plane Normal (VPN) - The direction that the virtual camera is pointing "away from"!
    Dim vectPRP As mdrVector4   ' Projection Reference Point (PRP), also known as Centre Of Projection (COP) - This is the distance between the virtual camera's film, and the pin-hole lens of the virtual camera.
    Dim vectVUP As mdrVector4   ' View UP direction (VUP) - Which way is up? This is used for tilting (or not tilting) the camera.
    
    
    ' Define the View Reference Point (VRP)
    ' This is defined in the World Coordinate (WC) system.
    vectVRP = p_Origin
    
    
    ' Subtract the Camera's world position (VRP) from the 'LookingAt' point to give us the View Plane Normal (VPN).
    ' VPN means different things to different 3D packages, ie. PHIGS and OpenGL do not agree on this one.
    ' In this application, the VPN points in the opposite direction that the camera is facing!! I said, Opposite!
    vectVPN = VectorSubtract(vectVRP, p_Target)

    
    ' The VUP vector points in the direction of UP. This is usually x=0,y=1,z=0.
    ' By changing where the UP direction is, you can tilt the camera.
    ' In our 3D world increasing values of Y go UP towards the sky, and negative Y values are below ground.
    ' We can easily trick the computer into think UP is actually sideways. Try using x=1,y=0,z=0. The camera will
    ' appear tilted on it's side, like the camera has fallen over.
    vectVUP.X = 0#
    vectVUP.Y = 1#
    vectVUP.Z = 0#
    vectVUP.w = 1#
        
    
    ' =====================================================
    ' Rotate VRC such that the:
    '   * n axis becomes the z axis,
    '   * u axis becomes the x axis and
    '   * v axis becomes the y axis.
    ' =====================================================

    Dim matRotateVRC As mdrMatrix4
    Dim vectN As mdrVector4
    Dim vectU As mdrVector4
    Dim vectV As mdrVector4


    '         VPN
    ' n* = ¯¯¯¯¯¯¯¯¯¯¯
    '       | VPN |
    '
    ' * also referred to as Rz (eq. 6.25)
    vectN = VectorNormalize(vectVPN)


    '       VUP x n
    ' u* = ¯¯¯¯¯¯¯¯¯¯¯¯¯
    '     | VUP x n |
    '
    ' * Also referred to as Rx (eq.6.26)
    vectU = CrossProduct(vectVUP, vectN)
    vectU = VectorNormalize(vectU)


    ' v* = n x u
    '
    ' * Also referred to as Ry (eq.6.27)
    vectV = CrossProduct(vectN, vectU)


    ' Define the Rotate matrix such that the n-axis (VPN) becomes the z-axis,
    ' the u-axis becomes the x-axis and the v-axis becomes the y-axis.
    MatrixLookAt = MatrixIdentity()
    With MatrixLookAt
        .rc11 = vectU.X: .rc12 = vectU.Y: .rc13 = vectU.Z
        .rc21 = vectV.X: .rc22 = vectV.Y: .rc23 = vectV.Z
        .rc31 = vectN.X: .rc32 = vectN.Y: .rc33 = vectN.Z
    End With
    
    ' Note sure why I used the following line of code... can't remember.
'    MatrixLookAt = MatrixTranspose(MatrixLookAt)
    
End Function
Public Function MakeWorldMatrix(Identity As mdrMatrix4, WorldPosition As mdrVector4, Pitch_x As Single, Yaw_y As Single, Roll_z As Single, UniformScale As Single) As mdrMatrix4
    
    Dim matTranslation As mdrMatrix4
    Dim matRotateX As mdrMatrix4
    Dim matRotateY As mdrMatrix4
    Dim matRotateZ As mdrMatrix4
    Dim matUniformScale As mdrMatrix4
    
    ' Setup the Scaling matrix.
    matUniformScale = MatrixScale(UniformScale, UniformScale, UniformScale, 1)
    
    ' Setup the Rotation matrices.
    matRotateX = MatrixRotationX(ConvertDeg2Rad(Pitch_x))
    matRotateY = MatrixRotationY(ConvertDeg2Rad(Yaw_y))
    matRotateZ = MatrixRotationZ(ConvertDeg2Rad(Roll_z))
    
    ' Setup the Translation matrix.
    matTranslation = MatrixTranslation(WorldPosition.X, WorldPosition.Y, WorldPosition.Z)
    
    
    ' Ok the part above is easy, because it doesn't really do anything except setup the matrices.
    ' The hard part comes when deciding the order in which to multiply them.  This is the part which WILL
    ' confuse you because you probably won't find two people that will agree on the same way of doing it.
    ' If this is the case, then "Don't Panic" this is normal and does not mean that you are wrong, or that
    ' other people are wrong.  There are in fact several ways to multiply matrices together. I am just
    ' showing one way of doing it. For instructional purposes, I have favoured clear code instead of fast code.
    '
    ' Note: DirectX and OpenGL multiply vectors-n-stuff in completely different ways, so be on your toes!
    '
    '       DirectX uses    Row-Vector notation.
    '       OpenGL  uses Column-Vector notation (like this program)
    '
    ' Overview of the process
    ' =======================
    ' NewWorldPoints = ModelingTransformation * OriginalObjectPoints
    '
    '
    ' Working it out long-hand, step by step, rotation by rotation, etc.
    ' (This part explains why we read and calculate from Right to Left. There are other methods that work
    '  equally as well from Left to Right, however this application uses Column Vectors and everything is Right to Left.)
    ' ===================================================================================================================
    ' Step/Equation 1) NewObjectPoints1 = ScaleMatrix         *   OriginalObjectPoint
    ' Step/Equation 2) NewObjectPoints2 = RollMatrix          *   NewObjectPoints1
    ' Step/Equation 3) NewObjectPoints3 = PitchMatrix         *   NewObjectPoints2
    ' Step/Equation 4) NewObjectPoints4 = YawMatrix           *   NewObjectPoints3
    ' Step/Equation 5) NewObjectPoints5 = TranslationMatrix   *   NewObjectPoints4
    '                    NewWorldPoints = NewObjectPoints5
    '
    '   A word on substitution
    '   ======================
    '   Substituting Equation 1 into Equation 2 we get...
    '   NewObjectPoints2 = RollMatrix * Scale * OriginalObjectPoint
    '
    '   If we expand this concept to all equations eventually we'll get...
    '   NewWorldPoints = Translation * Yaw * Pitch * Roll * Scale * OldObjectPoints
    '
    '   Ok, here's the bit that might get confusing... Your probably read the above equation from Left to Right,
    '   however you need to read it from Right to Left, because we are starting with our OldObjectPoints and then
    '   applying a series of transforms to get it into the NewWorldPoints (going from right to left).
    '
    '   For this routine, we won't be touching the OldObjectPoints at all.... we just want the concatenation of all
    '   of the matrices...
    '
    '       WorldMatrix = (Translation * Yaw * Pitch * Roll * Scale)
    '
    '       Remember, this is read from Right to Left (at least in this program), so Scale gets applied first, then Roll,
    '       then pitch and so on. Most 3D programmers get confused about this issue (including me), so it's important NOT
    '       to take guesses (as is common with many people), but to make sure you know exactly why you are applying a
    '       transform and the exact order that is applied with respect to the *whole program*.  If you don't understand
    '       this concept, you'll be forever taking guesses as to how to rotate complex objects, or setup special cameras.
    '
    ' I see a lot of 3D programmers rotate objects, like cubes, space ships, etc.  This is what most 3D programmers
    ' attempt to do in their first programs (including me). This is actually pretty easy to do, and there are several
    ' ways of doing it. However, this can also be a bit of a trap. Setting up a real synthetic camera is not so easy.
    ' Instead of rotating objects, you leave the objects alone and rotate the coordinate system instead. This requires a
    ' leap in logic that may seem hard to follow at first, and I suppose what I'm trying to say is, this is what
    ' separates the men from the boys (sorry girls). However once you fully understand the synthetic camera and all of the
    ' transforms to make it happen, many new possabilites begin to open up! Exciting possabilites!
    '
    ' eg. A real synthetic camera allows you to easily setup "rear vision mirrors", "split screen", or even splitting
    ' a view accross several monitors to achieve a wrap-around effect. Hmmmm.... I might make a network version
    ' of this program so that I can have three monitors in front of me to increase my 3D view of the world!.
    ' You could even program three separate computers to work on the same image on the same monitor - ie. Network rendering.
    ' If you're just fudging your synthetic camera, then the fudging becomes harder (if not impossible), when you try
    ' to do these sorts of exotic transforms, but this stuff is easy (well.. much easier) when you do the maths properly.
    '
    '
    ' Theory continuted...
    ' ========================================================================
    ' MakeWorldMatrix = Translation * Yaw * Pitch * Roll * Scale
    '                   (Remember, read this and calculate from Right to Left)
    ' ========================================================================

    ' Reset the WorldMatrix (If we didn't, then everything would be multiplied by zero and we'd have nothing!)
    MakeWorldMatrix = Identity
    
    ' Scale the object first,
    '   (MakeWorldMatrix = MakeWorldMatrix * Scale)
    MakeWorldMatrix = MatrixMultiply(MakeWorldMatrix, matUniformScale)

    ' ...then Roll the object... like banking in an areoplane,
    '   (MakeWorldMatrix = MakeWorldMatrix * Roll)
    MakeWorldMatrix = MatrixMultiply(MakeWorldMatrix, matRotateZ)

    ' ...then Pitch the object up or down,
    '   (MakeWorldMatrix = MakeWorldMatrix * Pitch)
    MakeWorldMatrix = MatrixMultiply(MakeWorldMatrix, matRotateX)

    ' ...then Yaw the object (point your aeroplane in the right compass direction,
    '   (MakeWorldMatrix = MakeWorldMatrix * Yaw)
    MakeWorldMatrix = MatrixMultiply(MakeWorldMatrix, matRotateY)

    ' ...and finally apply the translation/movement.
    '   (MakeWorldMatrix = MakeWorldMatrix * Translation)
    MakeWorldMatrix = MatrixMultiply(MakeWorldMatrix, matTranslation)
    

    ' Comments about "Gimbal Lock"
    ' ============================
    ' At some time during your rotations, you may notice a phenomenon called "Gimbal Lock".
    ' You will notice gimbal lock when your object won't rotate how you think it should.
    '
    ' For example...
    ' If your aeroplane is flying straight and level, then banking and yawing the plane works normally.
    ' However, if you Pitch your plane down into a 90 deg nose dive, then banking and yawing seem to do
    ' *exactly* the same thing!  You will have *lost* the ability to rotate your airplane on one
    ' of it's axis; This is the gimbal lock phenomenon. This also occurs in first-person shoot-em-up games
    ' where you use the mouse to look around; when you look straight up, you almost always get gimbal lock.
    
End Function

Public Function MakeWorldMatrixInverse(Identity As mdrMatrix4, WorldPosition As mdrVector4, Pitch_x As Single, Yaw_y As Single, Roll_z As Single, UniformScale As Single) As mdrMatrix4
    
    ' This function creates an inverse world martix. Basically, it's the complete opposite
    ' of 'MakeWorldMatrix'.
    ' eg.
    ' In the matrix world, the reversal of:
    '   a * b * c
    ' is NOT
    '   c * b * a
    '
    ' The correct answer is:
    '   -c * -b * -a
    
    Dim matTranslation As mdrMatrix4
    Dim matRotateX As mdrMatrix4
    Dim matRotateY As mdrMatrix4
    Dim matRotateZ As mdrMatrix4
    Dim matUniformScale As mdrMatrix4
    
    ' Setup the Scaling matrix.
    matUniformScale = MatrixScale(1 / UniformScale, 1 / UniformScale, 1 / UniformScale, 1)
    
    ' Setup the Rotation matrices.
    matRotateX = MatrixRotationX(ConvertDeg2Rad(-Pitch_x))
    matRotateY = MatrixRotationY(ConvertDeg2Rad(-Yaw_y))
    matRotateZ = MatrixRotationZ(ConvertDeg2Rad(-Roll_z))
    
    ' Setup the Translation matrix.
    matTranslation = MatrixTranslation(-WorldPosition.X, -WorldPosition.Y, -WorldPosition.Z)
    
    ' Reset.
    MakeWorldMatrixInverse = MatrixIdentity
    
    MakeWorldMatrixInverse = MatrixMultiply(MakeWorldMatrixInverse, matTranslation)
    MakeWorldMatrixInverse = MatrixMultiply(MakeWorldMatrixInverse, matRotateY)
    MakeWorldMatrixInverse = MatrixMultiply(MakeWorldMatrixInverse, matRotateX)
    MakeWorldMatrixInverse = MatrixMultiply(MakeWorldMatrixInverse, matRotateZ)
    MakeWorldMatrixInverse = MatrixMultiply(MakeWorldMatrixInverse, matUniformScale)
    
    
'    MakeWorldMatrixInverse = MatrixMultiply(MakeWorldMatrixInverse, matUniformScale)
'    MakeWorldMatrixInverse = MatrixMultiply(MakeWorldMatrixInverse, matRotateZ)
'    MakeWorldMatrixInverse = MatrixMultiply(MakeWorldMatrixInverse, matRotateX)
'    MakeWorldMatrixInverse = MatrixMultiply(MakeWorldMatrixInverse, matRotateY)
'    MakeWorldMatrixInverse = MatrixMultiply(MakeWorldMatrixInverse, matTranslation)
    
    
    ' Update debug & info. counters (conditionally compiled)
    #If g_blnCompiledDebugInfo = True Then
        g_lngDivisionCount = g_lngDivisionCount + 3
    #End If
    
End Function

Public Function IsBackFace(vectCamera As mdrVector4, vect1 As mdrVector4, vect2 As mdrVector4, vect3 As mdrVector4) As Double

    Dim vectR As mdrVector4
    Dim vectQ As mdrVector4
    Dim vectNormal As mdrVector4
    
    Dim COP As mdrVector4
    COP = vectCamera
    
    ' Get two vectors R and Q, that define "two edges" of a Polygon/Triangle
    ' (This is the part that checks if a triangle is drawn in clockwise or counter-clockwise direction)
    vectR = VectorSubtract(vect2, vect1)
    vectQ = VectorSubtract(vect3, vect2)
    
    ' Normal = R * Q
    vectNormal = CrossProduct(vectR, vectQ)
    
    ' Normalize the vectors (optional - only do if you want to calculate the cosine of the angle)
    ' If you just want to determine if the face is visible or not, then you don't have to normalize the
    ' vectors (rem the following two lines out) You will only need to know if the DotProduct is
    ' positive or negative. If you want to know the angle that the lights hits a face,
    ' then you'll need to normalize the vectors first.
'    vectNormal = VectorNormalize(vectNormal)    ' << This line may be remarked out.
'    COP = VectorNormalize(COP)                  ' << This line may be remarked out.
    
    IsBackFace = DotProduct(COP, vectNormal)
    
End Function

Public Function MatrixShadow(LightPosition As mdrVector4, PlaneEquation As mdrVector4) As mdrMatrix4
    
    '   You will need to define planes that you want to cast your shadows on (Walls, floor, and ceiling).
    '   As you may already know, a plane is defined by Ax+By+Cz + D = 0.  Given three points, you can easily
    '   calculate the parameters of this equations.  (hint: normal, Ax + By + Cz = -D)
    '
    '   Also, given a plane p defined by (A,B,C,D), light_position in homogeneous coordinate, and an
    '   object, a projection matrix is defined by:
    '
    '   dot = dotproduct(p, light_position)
    '
    '   Shadow Matrix =
    '       dot-(light_pos[0]*p[0]) -(light_pos[0]*p[1]) -(light_pos[0]*p[2]) -(light_pos[0]*p[3])
    '       -(light_pos[1]*p[0]) dot-(light_pos[1]*p[1]) -(light_pos[1]*p[2]) -(light_pos[1]*p[3])
    '       -(light_pos[2]*p[0]) -(light_pos[2]*p[1]) dot-(light_pos[2]*p[2]) -(light_pos[2]*p[3])
    '       -(light_pos[3]*p[0]) -(light_pos[3]*p[1]) -(light_pos[3]*p[2]) dot-(light_pos[3]*p[3])
    
    Dim sngDP As Single ' ie. DotProduct
    
    sngDP = DotProduct4(LightPosition, PlaneEquation)


    With MatrixShadow
        .rc11 = sngDP - (LightPosition.X * PlaneEquation.X)
        .rc12 = -(LightPosition.X * PlaneEquation.Y)
        .rc13 = -(LightPosition.X * PlaneEquation.Z)
        .rc14 = -(LightPosition.X * PlaneEquation.w)
        
        .rc21 = -(LightPosition.Y * PlaneEquation.X)
        .rc22 = sngDP - (LightPosition.Y * PlaneEquation.Y)
        .rc23 = -(LightPosition.Y * PlaneEquation.Z)
        .rc24 = -(LightPosition.Y * PlaneEquation.w)
        
        .rc31 = -(LightPosition.Z * PlaneEquation.X)
        .rc32 = -(LightPosition.Z * PlaneEquation.Y)
        .rc33 = sngDP - (LightPosition.Z * PlaneEquation.Z)
        .rc34 = -(LightPosition.Z * PlaneEquation.w)
        
        .rc41 = -(LightPosition.w * PlaneEquation.X)
        .rc42 = -(LightPosition.w * PlaneEquation.Y)
        .rc43 = -(LightPosition.w * PlaneEquation.Z)
        .rc44 = sngDP - (LightPosition.w * PlaneEquation.w)
    End With


    ' Update debug & info. counters (conditionally compiled)
    #If g_blnCompiledDebugInfo = True Then
        g_lngMultiplicationCount = g_lngMultiplicationCount + 16
        g_lngSubtractionCount = g_lngSubtractionCount + 4
    #End If

End Function

Public Function MatrixShear(ShearX As Single, ShearY As Single) As mdrMatrix4
    
    ' Create a new Identity matrix (i.e. Reset)
    MatrixShear = MatrixIdentity()
    
    ' Shear along the X and Y axis.
    '
    ' Shearing is used to distort an image for a very particular purpose. More specifically, it's
    ' used to help correctly reorient both the 3D objects and the observer (you)
    ' so that the image does NOT look distorted.
    ' You will NOT need to apply the Shear matrix to any part of your 3D objects (planes, tanks, etc.)
    
    MatrixShear.rc13 = ShearX
    MatrixShear.rc23 = ShearY
    
End Function

Public Function MatrixViewMapping_Per(p_Camera As mdr3DTargetCamera) As mdrMatrix4
        
    Dim vectCW As mdrVector4            '   Centre of Window
    Dim vectDOP As mdrVector4           '   Direction Of Projection
    Dim matTranslate As mdrMatrix4
    
    Dim sngShearX As Single
    Dim sngShearY As Single
    Dim matShear As mdrMatrix4
    
    Dim sngScaleX As Single
    Dim sngScaleY As Single
    Dim sngScaleZ As Single
    Dim matScale As mdrMatrix4
    
    Dim matPerspective As mdrMatrix4
    
    
    ' ===========================================================================================
    ' Translate such that the centre of projection (COP), given by PRP, is at the origin (p. 268)
    ' ===========================================================================================
    matTranslate = MatrixTranslation(-p_Camera.PRP.X, -p_Camera.PRP.Y, -p_Camera.PRP.Z)
    
    
    ' ===================================
    ' Calculate the Centre of the Window.
    ' ===================================
    vectCW.X = (p_Camera.Umax + p_Camera.Umin) / 2
    vectCW.Y = (p_Camera.Vmax + p_Camera.Vmin) / 2
    vectCW.Z = 0
    vectCW.w = 1
    
    
    ' ======================================================================================
    ' Calculate the difference between the Centre of the Window, and the PRP.
    ' The result is the Direction Of Projectsion (DOP), which should be the opposite of VPN.
    ' The DOP points in the direction of the camera, ie. Direction of Projection.
    ' ======================================================================================
    vectDOP = VectorSubtract(vectCW, p_Camera.PRP)
    
    
    ' ======================================================================
    ' Shear such that the center line of the view volume becomes the z-axis.
    ' ======================================================================
    If vectDOP.Z <> 0 Then
        sngShearX = -(vectDOP.X / vectDOP.Z)
        sngShearY = -(vectDOP.Y / vectDOP.Z)
    End If
    matShear = MatrixShear(sngShearX, sngShearY)
    
    
    ' ==========================================================================================
    ' Calculate the Perspective Scale transformation.
    ' Scale such that the view volume becomes the canonical perspective view volume, the
    ' truncated right pyramid defined by the six planes (ready for clipping) Eq. 6.39 on p.269.
    ' Here at the 6 planes after this step: x=z, x=-z, y=z, y=-z, z=-min, z=-1
    ' Diagram 6.56 - BEFORE.
    ' ==========================================================================================
    Dim sngTemp As Double
    sngScaleX = (2 * -p_Camera.PRP.Z) / ((p_Camera.Umax - p_Camera.Umin) * (-p_Camera.PRP.Z + p_Camera.ClipFar))
    sngScaleY = (2 * -p_Camera.PRP.Z) / ((p_Camera.Vmax - p_Camera.Vmin) * (-p_Camera.PRP.Z + p_Camera.ClipFar))
    sngScaleZ = -1 / (-p_Camera.PRP.Z + p_Camera.ClipFar)
    matScale = MatrixScale(sngScaleX, sngScaleY, sngScaleZ, 1)
    
    
    ' =======================================================================================================
    ' Ok... now that we have the "perspective-projection canonical view volume" (above), it is normal to
    ' covert this into the "parallel-projection canonical view volume". This is so a single clipping procedure
    ' can be used for both perspective and parallel.
    '
    ' zMin is the transform front clipping plane (Eq. 6.48)
    ' Here at the 6 planes equations after this step: x=-w, x=w, y=-w, y=w, z=-w, z=0
    ' -1 <= x/w =< 1, -1 <= y/w =< 1, -1 <= z/w =< 0
    ' Diagram 6.56 - AFTER.
    ' =======================================================================================================
    Dim sngZmin As Double
    sngZmin = -((-p_Camera.PRP.Z + p_Camera.ClipNear) / (-p_Camera.PRP.Z + p_Camera.ClipFar))
    matPerspective = MatrixIdentity
    If sngZmin <> -1 Then ' Minus one is the only value not allowed!
        matPerspective.rc33 = 1 / (1 + sngZmin)
        matPerspective.rc34 = -sngZmin / (1 + sngZmin)
        matPerspective.rc43 = -1
        matPerspective.rc44 = 0
    End If
    
    
    MatrixViewMapping_Per = MatrixIdentity()
    MatrixViewMapping_Per = MatrixMultiply(MatrixViewMapping_Per, matTranslate)
    MatrixViewMapping_Per = MatrixMultiply(MatrixViewMapping_Per, matShear)
    MatrixViewMapping_Per = MatrixMultiply(MatrixViewMapping_Per, matScale)
    MatrixViewMapping_Per = MatrixMultiply(MatrixViewMapping_Per, matPerspective)
    
    
    ' Update debug & info. counters (conditionally compiled)
    #If g_blnCompiledDebugInfo = True Then
        g_lngMultiplicationCount = g_lngMultiplicationCount + 4
        g_lngDivisionCount = g_lngDivisionCount + 10
        g_lngSubtractionCount = g_lngSubtractionCount + 2
        g_lngAdditionCount = g_lngAdditionCount + 7
    #End If
    
End Function


Public Function Matrix_vv3dv(withHandle As Form, Xmin As Single, Xmax As Single, Ymin As Single, Ymax As Single, Zmin As Single, Zmax As Single, ShowClipBoundary As Boolean, DayMode As Boolean) As mdrMatrix4
    
    Dim matTranslateA As mdrMatrix4
    Dim matTranslateB As mdrMatrix4
    Dim sngScaleX As Single
    Dim sngScaleY As Single
    Dim sngScaleZ As Single
    Dim matScale As mdrMatrix4
    
    
    ' Translate the canonical parallel projection view volume so that it's corner (-1,-1,-1) becomes
    ' the origin. This is so that the scaling process does not distort the geometry.
    ' ==============================================================================================
    matTranslateA = MatrixTranslation(1, 1, 1)
    
    
    ' The translated view volume is scaled into the size of the 3D viewport, with the following scale.
    ' This is the part where you turn those 'virtual' camera coordinates into real screen pixels!
    ' This scale matrix can also flip the y-coordinates so that (0,0) is at the bottom-left instead
    ' of the default windows location at top-left. You can also adjust the aspect ratio of the output here.
    ' NOTE for VB users: Visual Basic lets you change the scale-mode setting of a form,picturebox or pinter,
    '   This routine does pretty much the same thing, so if you wanted you could skip this
    '   "Matrix_vv3dv" step altogether and simply adjust the form/picturebox scale settings instead.
    '   If your not careful, you may get confused. There are many ways to do this, they all work.
    ' =========================================================================================================
    sngScaleX = (Xmax - Xmin) / 2
    sngScaleY = (Ymax - Ymin) / 2
    sngScaleZ = (Zmax - Zmin) / 1
    matScale = MatrixScale(sngScaleX, sngScaleY, sngScaleZ, 1)
    
    
    ' Finally, the properly scaled view volume is translated
    ' to the lower-left corner of the viewport
    ' =======================================================
    matTranslateB = MatrixTranslation(Xmin, Ymin, Zmin)
    
    
    ' Section: 6.5.5
    Matrix_vv3dv = MatrixIdentity()
    Matrix_vv3dv = MatrixMultiply(Matrix_vv3dv, matTranslateA)
    Matrix_vv3dv = MatrixMultiply(Matrix_vv3dv, matScale)
    Matrix_vv3dv = MatrixMultiply(Matrix_vv3dv, matTranslateB)


    ' ==============================================================
    ' Draw View Port Bounds (optional: can remark out / remove code)
    ' ==============================================================
    If DayMode = True Then
        withHandle.FillColor = RGB(220, 220, 220)
    Else
        withHandle.FillColor = RGB(48, 48, 48)
    End If
    withHandle.BackColor = withHandle.FillColor
    withHandle.FillStyle = vbFSSolid
    withHandle.DrawStyle = vbSolid
    withHandle.DrawWidth = 1
    If ShowClipBoundary = True Then
        If DayMode = True Then
            withHandle.Line (Xmin, Ymin)-(Xmax, Ymax), RGB(0, 0, 255), B ' Erase space & draw border..
        Else
            withHandle.Line (Xmin, Ymin)-(Xmax, Ymax), RGB(96, 0, 0), B ' Erase space & draw border..
        End If
    Else
        withHandle.Cls
    End If


    ' Update debug & info. counters (conditionally compiled)
    #If g_blnCompiledDebugInfo = True Then
        g_lngSubtractionCount = g_lngSubtractionCount + 3
        g_lngDivisionCount = g_lngDivisionCount + 3
    #End If

End Function
Public Function MatrixScale(ScaleX As Single, ScaleY As Single, ScaleZ As Single, ScaleW As Single) As mdrMatrix4
    
    ' Create a new Identity matrix (i.e. Reset)
    MatrixScale = MatrixIdentity()
    
    ' Makes an object bigger or smaller on any of the three axes.
    ' Note: Some imported .x files need scaling in the order of 100, 200 even 1000, while other objects
    '       do not need scaling at all.
    '       ie. If the scale factor is 2,2,2 then the object is doubled in size on all three
    '           axes. A scale factor of 0.5,0.5,0.5 will shrink an object on all three axes.
    '
    ' Normally you scale on all three axis, by the same amount, otherwise your 3D object may get
    ' distorted in a way you didn't expect. If in doubt, just make all the numbers the same....
    ' a Uniform Scale.
    
    MatrixScale.rc11 = ScaleX
    MatrixScale.rc22 = ScaleY
    MatrixScale.rc33 = ScaleZ
    MatrixScale.rc44 = ScaleW
    
End Function

Public Function ConvertFOVtoZoom(FOV As Single) As Single
    
    ' Given a Field Of View, calculate the Zoom.
    ConvertFOVtoZoom = 1 / Tan(ConvertDeg2Rad(FOV) / 2)
    
    
    ' Update debug & info. counters (conditionally compiled)
    #If g_blnCompiledDebugInfo = True Then
        g_lngDivisionCount = g_lngDivisionCount + 2
    #End If
    
End Function

Public Function ConvertZoomtoFOV(Zoom As Single) As Single
    
    ' Given a Zoom value, calculate the 'Field Of View'
    ConvertZoomtoFOV = ConvertRad2Deg(2 * Atn(1 / Zoom))
    
    
    ' Update debug & info. counters (conditionally compiled)
    #If g_blnCompiledDebugInfo = True Then
        g_lngMultiplicationCount = g_lngMultiplicationCount + 1
        g_lngDivisionCount = g_lngDivisionCount + 1
    #End If
    
End Function

Public Function ConvertDeg2Rad(Degress As Single) As Single
Attribute ConvertDeg2Rad.VB_Description = "Converts Degrees to Radians."

    ' Converts Degrees to Radians
    ConvertDeg2Rad = Degress * (g_sngPIDivideBy180)
    
    
    ' Update debug & info. counters (conditionally compiled)
    #If g_blnCompiledDebugInfo = True Then
        g_lngMultiplicationCount = g_lngMultiplicationCount + 1
    #End If
    
End Function

Public Function ConvertRad2Deg(Radians As Single) As Single
Attribute ConvertRad2Deg.VB_Description = "Converts Radians to Degrees."
 
    ' Converts Radians to Degrees
    ConvertRad2Deg = Radians * (g_sng180DivideByPI)
    
    
    ' Update debug & info. counters (conditionally compiled)
    #If g_blnCompiledDebugInfo = True Then
        g_lngMultiplicationCount = g_lngMultiplicationCount + 1
    #End If
    
End Function

Public Function DotProduct(VectorU As mdrVector4, VectorV As mdrVector4) As Single
Attribute DotProduct.VB_Description = "Returns to the DotProduct of two vectors."
Attribute DotProduct.VB_HelpID = 1

    ' Determines the dot-product of two 4D vectors (ignoring the W component)
    DotProduct = (VectorU.X * VectorV.X) + (VectorU.Y * VectorV.Y) + (VectorU.Z * VectorV.Z)
    
    
    ' Update debug & info. counters (conditionally compiled)
    #If g_blnCompiledDebugInfo = True Then
        g_lngMultiplicationCount = g_lngMultiplicationCount + 3
        g_lngAdditionCount = g_lngAdditionCount + 2
    #End If
    
End Function

Public Function DotProduct4(VectorU As mdrVector4, VectorV As mdrVector4) As Single

    ' Determines the dot-product of two 4D vectors.
    ' ==============================================
    DotProduct4 = (VectorU.X * VectorV.X) + (VectorU.Y * VectorV.Y) + (VectorU.Z * VectorV.Z) + (VectorU.w * VectorV.w)
    
    
    ' Update debug & info. counters (conditionally compiled)
    #If g_blnCompiledDebugInfo = True Then
        g_lngMultiplicationCount = g_lngMultiplicationCount + 4
        g_lngAdditionCount = g_lngAdditionCount + 3
    #End If
    
End Function

Public Function MatrixViewOrientation(vectVPN As mdrVector4, vectVUP As mdrVector4, vectVRP As mdrVector4) As mdrMatrix4
Attribute MatrixViewOrientation.VB_Description = "Builds a ViewOrientation Matrix to correctly orientate the scene. VPN = View Plane Normal, VUP=Up Vector, VRP=View Reference Point."
    
    ' =====================================================
    ' Rotate VRC such that the:
    '   * n axis becomes the z axis,
    '   * u axis becomes the x axis and
    '   * v axis becomes the y axis.
    ' =====================================================
    
    Dim matRotateVRC As mdrMatrix4
    Dim matTranslateVRP As mdrMatrix4
    
    Dim vectN As mdrVector4
    Dim vectU As mdrVector4
    Dim vectV As mdrVector4
    
        
    '         VPN
    ' n* = ¯¯¯¯¯¯¯¯¯¯¯
    '       | VPN |
    '
    ' * also referred to as Rz (eq. 6.25)
    vectN = VectorNormalize(vectVPN)
    
    
    '         VUP x n
    ' u* = ¯¯¯¯¯¯¯¯¯¯¯¯¯
    '       | VUP x n |
    '
    ' * Also referred to as Rx (eq.6.26)
    vectU = CrossProduct(vectVUP, vectN)
    vectU = VectorNormalize(vectU)
    
    
    ' v* = n x u
    '
    ' * Also referred to as Ry (eq.6.27)
    vectV = CrossProduct(vectN, vectU)
    
    
    ' Define the Rotate matrix such that the n-axis (VPN) becomes the z-axis,
    ' the u-axis becomes the x-axis and the v-axis becomes the y-axis.
    matRotateVRC = MatrixIdentity()
    With matRotateVRC
        .rc11 = vectU.X: .rc12 = vectU.Y: .rc13 = vectU.Z
        .rc21 = vectV.X: .rc22 = vectV.Y: .rc23 = vectV.Z
        .rc31 = vectN.X: .rc32 = vectN.Y: .rc33 = vectN.Z
    End With
    
    
    ' Define a Translation matrix to transform the VRP to the origin.
    matTranslateVRP = MatrixTranslation(-vectVRP.X, -vectVRP.Y, -vectVRP.Z)
    
    
    ' Theory
    ' ===============================================================================
    ' MatrixViewOrientation =  matTranslateVRP * matRotateVRC
    '                          (Remember, read this and calculate from Right to Left)
    ' ===============================================================================
    MatrixViewOrientation = MatrixIdentity()
    MatrixViewOrientation = MatrixMultiply(MatrixViewOrientation, matTranslateVRP)
    MatrixViewOrientation = MatrixMultiply(MatrixViewOrientation, matRotateVRC)
    
    
End Function

Public Function VectorCrossProduct(v1 As mdrVector4, v2 As mdrVector4) As mdrVector4

    ' Returns the Cross-Product of two vectors.
    With VectorCrossProduct
        .X = (v1.Y * v2.Z) - (v1.Z * v2.Y)
        .Y = (v1.Z * v2.X) - (v1.X * v2.Z)
        .Z = (v1.X * v2.Y) - (v1.Y * v2.X)
        .w = 1
    End With
    
    ' Update debug & info. counters (conditionally compiled)
    #If g_blnCompiledDebugInfo = True Then
        g_lngMultiplicationCount = g_lngMultiplicationCount + 6
        g_lngSubtractionCount = g_lngSubtractionCount + 3
    #End If
    
End Function

Public Function VectorSubtract(v1 As mdrVector4, v2 As mdrVector4) As mdrVector4
Attribute VectorSubtract.VB_Description = "Returns the result of Vector2 subtracted from Vector1."

    ' Subtracts vector 2 away from vector 1.
    With VectorSubtract
        .X = v1.X - v2.X
        .Y = v1.Y - v2.Y
        .Z = v1.Z - v2.Z
        .w = 1 ' Ignore W
    End With
    
    
    ' Update debug & info. counters (conditionally compiled)
    #If g_blnCompiledDebugInfo = True Then
        g_lngSubtractionCount = g_lngSubtractionCount + 3
    #End If
    
End Function

Public Function MatrixTranspose(MIn As mdrMatrix4) As mdrMatrix4
    
    ' Swaps Rows for Columns (and visa-versa) in a 4x4 matrix.
    ' Useful for converting between the OpenGL matrix format and the Direct-X matrix format.
    
    With MatrixTranspose
        
        .rc11 = MIn.rc11: .rc12 = MIn.rc21: .rc13 = MIn.rc31: .rc14 = MIn.rc41
        .rc21 = MIn.rc12: .rc22 = MIn.rc22: .rc23 = MIn.rc32: .rc24 = MIn.rc42
        .rc31 = MIn.rc13: .rc32 = MIn.rc23: .rc33 = MIn.rc33: .rc34 = MIn.rc43
        .rc41 = MIn.rc14: .rc42 = MIn.rc24: .rc43 = MIn.rc34: .rc44 = MIn.rc44
        
    End With
    
End Function

Public Function VectorAddition(v1 As mdrVector4, v2 As mdrVector4) As mdrVector4
Attribute VectorAddition.VB_Description = "Returns the result of two Vectors added together."

    ' Adds two vectors together.
    With VectorAddition
        .X = v1.X + v2.X
        .Y = v1.Y + v2.Y
        .Z = v1.Z + v2.Z
        .w = 1 ' Ignore W
    End With
    
    
    ' Update debug & info. counters (conditionally compiled)
    #If g_blnCompiledDebugInfo = True Then
        g_lngAdditionCount = g_lngAdditionCount + 3
    #End If
    
End Function


Public Function MatrixTranslation(OffsetX As Single, OffsetY As Single, OffsetZ As Single) As mdrMatrix4
Attribute MatrixTranslation.VB_Description = "Given X, Y and Z offsets, builds a Translation Matrix."
    
    ' Translation is another word for "move".
    ' ie. You can translate an object from one location to another.
    '     You can    move   an object from one location to another.
    '
    ' The ability to combine a Rotation with a Translation within a single matrix, is the main
    ' reason why I have used a 4x4 matrix and NOT a 3x3 matrix.
    
    ' Create a new Identity matrix (i.e. Reset)
    MatrixTranslation = MatrixIdentity()
    
    With MatrixTranslation
        .rc14 = OffsetX
        .rc24 = OffsetY
        .rc34 = OffsetZ
    End With
    
    ' Very important note about this matrix
    ' =====================================
    ' If you see other programmers placing their Offset's in different positions (like the columns
    ' and rows have been swapped over - ie. Transposed) then this probably means that they have coded all
    ' of their algorithims to a different "notation standard". This subroutine follows the conventions used
    ' in the ledgendary bible "Computer Graphics Principles and Practice", Foley·vanDam·Feiner·Hughes which
    ' illustrates mathematical formulas using Column-Vector notation. Other books like "3D Math Primer for
    ' Graphics and Game Development", Fletcher Dunn·Ian Parberry, use Row-Vector notation. Both are correct,
    ' however it's important to know which standard you code to, because it affects the way in which you
    ' build your matrices and the order in which you should multiply them to obtain the correct result.
    '
    ' OpenGL uses Column Vectors (like this application).
    ' DirectX uses Row Vectors.
    
End Function

Public Function MatrixIdentity() As mdrMatrix4
Attribute MatrixIdentity.VB_Description = "Returns an Identity matrix."

    ' The identity matrix is used as the starting point for matrices
    ' that will modify vertex values to create rotations, translations,
    ' and any other transformations that can be represented by a 4×4 matrix
    '
    ' Notice that...
    '   * the 1's go diagonally down?
    '   * rc stands for Row Column. Therefore, rc12 means Row1, Column 2.
    '
    ' Comments:
    ' You'll often hear people talking about the "identity matrix"... well this is it!
    ' Sometimes the identity matrix also contains pre-calculated rotations and translations. This is usually
    ' the case when you import a 3D object from another application.
    
    With MatrixIdentity
        .rc11 = 1: .rc12 = 0: .rc13 = 0: .rc14 = 0
        .rc21 = 0: .rc22 = 1: .rc23 = 0: .rc24 = 0
        .rc31 = 0: .rc32 = 0: .rc33 = 1: .rc34 = 0
        .rc41 = 0: .rc42 = 0: .rc43 = 0: .rc44 = 1
    End With
    
End Function

Public Function MatrixMultiply(m1 As mdrMatrix4, m2 As mdrMatrix4) As mdrMatrix4
Attribute MatrixMultiply.VB_Description = "Returns the result of Matrix1 multiplied by Matrix2."
    
    ' Re-declare m1 & m2
    Dim m1b As mdrMatrix4
    Dim m2b As mdrMatrix4
    m1b = m1
    m2b = m2
    
    ' Matrix multiplication is a set of "dot products" between the rows of the left matrix and columns of the right matrix.
    '
    ' Matrix A and B below
    ' ====================
    '                          | a, b, c |       | j, k, l |
    '  Let A*B represent...    | d, e, f |   *   | m, n, o |
    '                          | g, h, i |       | p, q, r |
    '
    '  Multipling out we get...
    '
    '   | (a*j)+(b*m)+(c*p), (a*k)+(b*n)+(c*q), (a*l)+(b*o)+(c*r) |
    '   | (d*j)+(e*m)+(f*p), (d*k)+(e*n)+(f*q), (d*l)+(e*o)+(f*r) |
    '   | (g*j)+(h*m)+(i*p), (g*k)+(h*n)+(i*q), (g*l)+(h*o)+(i*r) |
    '
    ' To put this another way...
    '
    '  | a, b, c |     | j, k, l |     | (a*j)+(b*m)+(c*p), (a*k)+(b*n)+(c*q), (a*l)+(b*o)+(c*r) |
    '  | d, e, f |  *  | m, n, o |  =  | (d*j)+(e*m)+(f*p), (d*k)+(e*n)+(f*q), (d*l)+(e*o)+(f*r) |
    '  | g, h, i |     | p, q, r |     | (g*j)+(h*m)+(i*p), (g*k)+(h*n)+(i*q), (g*l)+(h*o)+(i*r) |
    '
    ' Note: This was only a 3x3 matrix show... however this routine is actually bigger, using a 4x4.
    ' I just wanted to keep the example short.
    
    
    ' =====================
    ' About this subroutine
    ' =====================
    ' This is the kind of routine that is hard coded into the electronic circuts of many CPU's and
    ' all 3D video cards (actually most of this module is hard coded into the video-cards, in some way or another)
    ' For additional research try searching for "Matrix Multiplication"
    '
    ' Multiply two 4x4 matrices (m2 & m1) and return the result in 'MatrixMultiply'.
    '   64 Floating point multiplications
    '   48 Floating point additions
    '
    ' This matrix multiplies a full 4x4 matrix, however some programmers and/or algorithms only
    ' multiply the top-left 3x3; yes, you can do this, however a 4x4 matrix lets you combine rotation
    ' and movement in a single matrix. If you are using a 3x3 matrix then you can't do this and
    ' will have to calculate rotation and movement as separate steps. A 3x3 matrix also makes it
    ' harder to rotate an object around a point that is not it's origin. Heck! There's a lot of
    ' agruments about 3x3 vs. 4x4, and I can't be bothered getting into them. Just do it the correct
    ' way and everyone will be happy! ;-)
    
    
    ' Reset the matrix to identity.
    MatrixMultiply = MatrixIdentity()
    
    
    With MatrixMultiply
        .rc11 = (m1b.rc11 * m2b.rc11) + (m1b.rc21 * m2b.rc12) + (m1b.rc31 * m2b.rc13) + (m1b.rc41 * m2b.rc14)
        .rc12 = (m1b.rc12 * m2b.rc11) + (m1b.rc22 * m2b.rc12) + (m1b.rc32 * m2b.rc13) + (m1b.rc42 * m2b.rc14)
        .rc13 = (m1b.rc13 * m2b.rc11) + (m1b.rc23 * m2b.rc12) + (m1b.rc33 * m2b.rc13) + (m1b.rc43 * m2b.rc14)
        .rc14 = (m1b.rc14 * m2b.rc11) + (m1b.rc24 * m2b.rc12) + (m1b.rc34 * m2b.rc13) + (m1b.rc44 * m2b.rc14)
        
        .rc21 = (m1b.rc11 * m2b.rc21) + (m1b.rc21 * m2b.rc22) + (m1b.rc31 * m2b.rc23) + (m1b.rc41 * m2b.rc24)
        .rc22 = (m1b.rc12 * m2b.rc21) + (m1b.rc22 * m2b.rc22) + (m1b.rc32 * m2b.rc23) + (m1b.rc42 * m2b.rc24)
        .rc23 = (m1b.rc13 * m2b.rc21) + (m1b.rc23 * m2b.rc22) + (m1b.rc33 * m2b.rc23) + (m1b.rc43 * m2b.rc24)
        .rc24 = (m1b.rc14 * m2b.rc21) + (m1b.rc24 * m2b.rc22) + (m1b.rc34 * m2b.rc23) + (m1b.rc44 * m2b.rc24)
        
        .rc31 = (m1b.rc11 * m2b.rc31) + (m1b.rc21 * m2b.rc32) + (m1b.rc31 * m2b.rc33) + (m1b.rc41 * m2b.rc34)
        .rc32 = (m1b.rc12 * m2b.rc31) + (m1b.rc22 * m2b.rc32) + (m1b.rc32 * m2b.rc33) + (m1b.rc42 * m2b.rc34)
        .rc33 = (m1b.rc13 * m2b.rc31) + (m1b.rc23 * m2b.rc32) + (m1b.rc33 * m2b.rc33) + (m1b.rc43 * m2b.rc34)
        .rc34 = (m1b.rc14 * m2b.rc31) + (m1b.rc24 * m2b.rc32) + (m1b.rc34 * m2b.rc33) + (m1b.rc44 * m2b.rc34)
        
        .rc41 = (m1b.rc11 * m2b.rc41) + (m1b.rc21 * m2b.rc42) + (m1b.rc31 * m2b.rc43) + (m1b.rc41 * m2b.rc44)
        .rc42 = (m1b.rc12 * m2b.rc41) + (m1b.rc22 * m2b.rc42) + (m1b.rc32 * m2b.rc43) + (m1b.rc42 * m2b.rc44)
        .rc43 = (m1b.rc13 * m2b.rc41) + (m1b.rc23 * m2b.rc42) + (m1b.rc33 * m2b.rc43) + (m1b.rc43 * m2b.rc44)
        .rc44 = (m1b.rc14 * m2b.rc41) + (m1b.rc24 * m2b.rc42) + (m1b.rc34 * m2b.rc43) + (m1b.rc44 * m2b.rc44)
    End With
    
    
    ' Update debug & info. counters (conditionally compiled)
    #If g_blnCompiledDebugInfo = True Then
        g_lngMultiplicationCount = g_lngMultiplicationCount + 64
        g_lngAdditionCount = g_lngAdditionCount + 48
    #End If
    
End Function

Public Function MatrixMultiplyVector(m1 As mdrMatrix4, v1 As mdrVector4) As mdrVector4
Attribute MatrixMultiplyVector.VB_Description = "Returns the result of a Matrix multiplied by a Vector."
        
    ' Here is a Column Vector (having three letters/numbers)...
    '
    '   | a |
    '   | b |
    '   | c |
    '
    ' Here is the Row Vector equivalent...
    '
    '   | a, b, c |
    '
    ' The two different conventions (Column Vector, Row Vector) store exactly the same information,
    ' so the issue of which is best will not even be discussed!  Just remember that different authors use different
    ' conventions, and it's quite easy to get them mixed up with each other!
    
    
    
    ' Matrix multiplication is a set of "dot products" between the rows of the left matrix and columns of the right matrix.
    '
    ' Matrix A and B below
    ' ====================
    '                            | a, b, c |     | x |
    '  Note the following...     | d, e, f |  *  | y |
    '                            | g, h, i |     | z |
    '
    '  ...multipling out we get...
    '
    '   | (a*x)+(b*y)+(c*z) |
    '   | (d*x)+(e*y)+(f*z) |
    '   | (g*x)+(h*y)+(i*z) |
    
    '
    ' Therefore...
    '
    '   | a, b, c |     | x |     | (a*x)+(b*y)+(c*z) |
    '   | d, e, f |  *  | y |  =  | (d*x)+(e*y)+(f*z) |
    '   | g, h, i |     | z |     | (g*x)+(h*y)+(i*z) |
    
    
    
    
    
    ' Multiply two matrices (m1 & v1) and returns the result in VOut.
    '
    ' m1 is a 4x4 matrix (ColumnsN = 4)
    ' v1 is a Column vector matrix (RowsM = 4 rows)
    '
    ' Because ColumnsN equals RowsM, this is considered a 'Square Matrix' and can be multiplied.
    ' (Notice how the reverse is NOT true: Columns of v1 = 1, Rows of m1 = 4, they are not the
    '  same and thus can't be multiplied in reverse order.)
    '
    ' 16 Floating point multiplications
    ' 12 Floating point additions
    
    With MatrixMultiplyVector
        .X = (m1.rc11 * v1.X) + (m1.rc12 * v1.Y) + (m1.rc13 * v1.Z) + (m1.rc14 * v1.w)
        .Y = (m1.rc21 * v1.X) + (m1.rc22 * v1.Y) + (m1.rc23 * v1.Z) + (m1.rc24 * v1.w)
        .Z = (m1.rc31 * v1.X) + (m1.rc32 * v1.Y) + (m1.rc33 * v1.Z) + (m1.rc34 * v1.w)
        .w = (m1.rc41 * v1.X) + (m1.rc42 * v1.Y) + (m1.rc43 * v1.Z) + (m1.rc44 * v1.w)
    End With
    
    
    ' Update debug & info. counters (conditionally compiled)
    #If g_blnCompiledDebugInfo = True Then
        g_lngMultiplicationCount = g_lngMultiplicationCount + 16
        g_lngAdditionCount = g_lngAdditionCount + 12
    #End If
    
End Function

Public Function VectorNormalize(v As mdrVector4) As mdrVector4
Attribute VectorNormalize.VB_Description = "Returns the normalized version of a Vector. The resulting Vector will have a length equal to 1.0"

    ' Returns the normalized version of a vector.
    
    Dim sngLength As Single
    
    sngLength = VectorLength(v)
    If sngLength = 0 Then sngLength = 1
    
    With VectorNormalize
        .X = v.X / sngLength
        .Y = v.Y / sngLength
        .Z = v.Z / sngLength
        .w = v.w ' Ignore W
    End With
    
    
    ' Update debug & info. counters (conditionally compiled)
    #If g_blnCompiledDebugInfo = True Then
        g_lngDivisionCount = g_lngDivisionCount + 3
    #End If
    
End Function

Public Function VectorLength(v As mdrVector4) As Single
Attribute VectorLength.VB_Description = "Returns the length of a Vector using Pythagoras therom."

    ' Returns the length of a Vector.
    '
    ' In Mathematic books, the "length of a vector" is often written with two vertical bars on either
    ' side, like this:  ||v||
    ' It took me ages to figure this out! Nobody explained it, they just assumed I knew it!
    '
    ' The length of a vector is from the origin (0,0,0) to x,y,z
    ' Do you remember high schools maths, Pythagoras theorem?  c^2 = a^2 + b^2
    '   "In a right-angled triangle, the area of the square of the hypotenuse (the longest side)
    '    is equal to the sum of the areas of the squares drawn on the other two sides."
    
    VectorLength = Sqr((v.X ^ 2) + (v.Y ^ 2) + (v.Z ^ 2))
    ' Ignore W
    
    ' Update debug & info. counters (conditionally compiled)
    #If g_blnCompiledDebugInfo = True Then
        g_lngAdditionCount = g_lngAdditionCount + 2
        g_lngSquareRootCount = g_lngSquareRootCount + 3
    #End If
    
End Function

Public Function CrossProduct(vectV As mdrVector4, VectW As mdrVector4) As mdrVector4
Attribute CrossProduct.VB_Description = "Returns the CrossProduct of two vectors."

    ' Determines the cross-product of two 3-D vectors (V and W).
    ' The cross-product is used to find a vector that is perpendicular to the plane defined by VectV and VectW.
    
    With CrossProduct
        .X = (vectV.Y * VectW.Z) - (vectV.Z * VectW.Y)
        .Y = (vectV.Z * VectW.X) - (vectV.X * VectW.Z)
        .Z = (vectV.X * VectW.Y) - (vectV.Y * VectW.X)
        .w = 1 ' Ignore W
    End With
    
    
    
    ' Update debug & info. counters (conditionally compiled)
    #If g_blnCompiledDebugInfo = True Then
        g_lngMultiplicationCount = g_lngMultiplicationCount + 6
        g_lngSubtractionCount = g_lngSubtractionCount + 3
    #End If
    
End Function

Public Function VectorMultiplyByScalar(VectorIn As mdrVector4, Scalar As Single) As mdrVector4
Attribute VectorMultiplyByScalar.VB_Description = "Returns the result of a Vector multiplied by a scalar value. Useful for making vectors bigger or smaller."
    
    With VectorMultiplyByScalar
        .X = CSng(VectorIn.X) * CSng(Scalar)
        .Y = CSng(VectorIn.Y) * CSng(Scalar)
        .Z = CSng(VectorIn.Z) * CSng(Scalar)
        .w = VectorIn.w ' Ignore W
    End With
    
    ' Update debug & info. counters (conditionally compiled)
    #If g_blnCompiledDebugInfo = True Then
        g_lngMultiplicationCount = g_lngMultiplicationCount + 3
    #End If
    
End Function

Public Function MatrixRotationX(Radians As Single) As mdrMatrix4
Attribute MatrixRotationX.VB_Description = "Given an angle expressed as Radians, this function builds a Rotation Matrix for the X Axis."

    ' ===========================================================================================
    '               *** This application uses a Right-Handed Coordinate System ***
    ' ===========================================================================================
    '   The positive X axis points towards the right.
    '   The positive Y axis points upwards to the top of the screen.
    '   The positive Z axis points out of the monitor towards you.
    '
    ' Note: DirectX uses a Left-Handed Coordinate system, which many people find more intuitive.
    ' ===========================================================================================
    
    Dim sngCosine As Double
    Dim sngSine As Double
    
    sngCosine = Round(Cos(Radians), 6)
    sngSine = Round(Sin(Radians), 6)
    
    ' Create a new Identity matrix (i.e. Reset)
    MatrixRotationX = MatrixIdentity()
    
    ' =======================================================================================================
    ' Positive rotations in a right-handed coordinate system are such that, when looking from a
    ' positive axis back towards the origin (0,0,0), a 90° "counter-clockwise" rotation will
    ' transform one positive axis into the other:
    '
    ' X-Axis rotation.
    ' A positive rotation of 90° transforms the +Y axis into the +Z axis.
    ' An additional positive rotation of 90° transforms the +Z axis into the -Y axis.
    ' An additional positive rotation of 90° transforms the -Y axis into the -Z axis.
    ' An additional positive rotation of 90° transforms the -Z axis into the +Y axis (back where we started).
    ' =======================================================================================================
    With MatrixRotationX
        .rc22 = sngCosine
        .rc23 = -sngSine
        .rc32 = sngSine
        .rc33 = sngCosine
    End With
    
    ' Very important note about this matrix
    ' =====================================
    ' If you see other programmers placing their Sines and Cosines in different positions (like the columns
    ' and rows have been swapped over - ie. Transposed) then this probably means that they have coded all
    ' of their algorithims to a different "notation standard". This subroutine follows the conventions used
    ' in the ledgendary bible "Computer Graphics Principles and Practice", Foley·vanDam·Feiner·Hughes which
    ' illustrates mathematical formulas using Column-Vector notation. Other books like "3D Math Primer for
    ' Graphics and Game Development", Fletcher Dunn·Ian Parberry, use Row-Vector notation. Both are correct,
    ' however it's important to know which standard you code to, because it affects the way in which you
    ' build your matrices and the order in which you should multiply them to obtain the correct result.
    '
    ' OpenGL uses Column Vectors (like this application).
    ' DirectX uses Row Vectors.
    
End Function

Public Function MatrixRotationY(Radians As Single) As mdrMatrix4
Attribute MatrixRotationY.VB_Description = "Given an angle expressed as Radians, this function builds a Rotation Matrix for the Y Axis."

    ' ===========================================================================================
    '               *** This application uses a Right-Handed Coordinate System ***
    ' ===========================================================================================
    '   The positive X axis points towards the right.
    '   The positive Y axis points upwards to the top of the screen.
    '   The positive Z axis points out of the monitor towards you.
    '
    ' Note: DirectX uses a Left-Handed Coordinate system, which many people find more intuitive.
    ' ===========================================================================================
    
    Dim sngCosine As Double
    Dim sngSine As Double
    
    sngCosine = Round(Cos(Radians), 6)
    sngSine = Round(Sin(Radians), 6)
    
    ' Create a new Identity matrix (i.e. Reset)
    MatrixRotationY = MatrixIdentity()
    
    ' =======================================================================================================
    ' Positive rotations in a right-handed coordinate system are such that, when looking from a
    ' positive axis back towards the origin (0,0,0), a 90° "counter-clockwise" rotation will
    ' transform one positive axis into the other:
    '
    ' Y-Axis rotation.
    ' A positive rotation of 90° transforms the +Z axis into the +X axis
    ' An additional positive rotation of 90° transforms the +X axis into the -Z axis.
    ' An additional positive rotation of 90° transforms the -Z axis into the -X axis.
    ' An additional positive rotation of 90° transforms the -X axis into the +Z axis (back where we started).
    ' =======================================================================================================
    With MatrixRotationY
        .rc11 = sngCosine
        .rc31 = -sngSine
        .rc13 = sngSine
        .rc33 = sngCosine
    End With
    
    ' Very important note about this matrix
    ' =====================================
    ' If you see other programmers placing their Sines and Cosines in different positions (like the columns
    ' and rows have been swapped over - ie. Transposed) then this probably means that they have coded all
    ' of their algorithims to a different "notation standard". This subroutine follows the conventions used
    ' in the ledgendary bible "Computer Graphics Principles and Practice", Foley·vanDam·Feiner·Hughes which
    ' illustrates mathematical formulas using Column-Vector notation. Other books like "3D Math Primer for
    ' Graphics and Game Development", Fletcher Dunn·Ian Parberry, use Row-Vector notation. Both are correct,
    ' however it's important to know which standard you code to, because it affects the way in which you
    ' build your matrices and the order in which you should multiply them to obtain the correct result.
    '
    ' OpenGL uses Column Vectors (like this application).
    ' DirectX uses Row Vectors.

End Function

Public Function MatrixRotationZ(Radians As Single) As mdrMatrix4
Attribute MatrixRotationZ.VB_Description = "Given an angle expressed as Radians, this function builds a Rotation Matrix for the Z Axis."

    ' ===========================================================================================
    '               *** This application uses a Right-Handed Coordinate System ***
    ' ===========================================================================================
    '   The positive X axis points towards the right.
    '   The positive Y axis points upwards to the top of the screen.
    '   The positive Z axis points out of the monitor towards you.
    '
    ' Note: DirectX uses a Left-Handed Coordinate system, which many people find more intuitive.
    ' ===========================================================================================
    
    
    Dim sngCosine As Double
    Dim sngSine As Double
    
    sngCosine = Round(Cos(Radians), 6)
    sngSine = Round(Sin(Radians), 6)
    
    ' Create a new Identity matrix (i.e. Reset)
    MatrixRotationZ = MatrixIdentity()

    ' =======================================================================================================
    ' Positive rotations in a right-handed coordinate system are such that, when looking from a
    ' positive axis back towards the origin (0,0,0), a 90° "counter-clockwise" rotation will
    ' transform one positive axis into the other:
    '
    ' Z-Axis rotation.
    ' A positive rotation of 90° transforms the +X axis into the +Y axis.
    ' An additional positive rotation of 90° transforms the +Y axis into the -X axis.
    ' An additional positive rotation of 90° transforms the -X axis into the -Y axis.
    ' An additional positive rotation of 90° transforms the -Y axis into the +X axis (back where we started).
    ' =======================================================================================================
    With MatrixRotationZ
        .rc11 = sngCosine
        .rc21 = sngSine
        .rc12 = -sngSine
        .rc22 = sngCosine
    End With
    
    ' Very important note about this matrix
    ' =====================================
    ' If you see other programmers placing their Sines and Cosines in different positions (like the columns
    ' and rows have been swapped over - ie. Transposed) then this probably means that they have coded all
    ' of their algorithims to a different "notation standard". This subroutine follows the conventions used
    ' in the ledgendary bible "Computer Graphics Principles and Practice", Foley·vanDam·Feiner·Hughes which
    ' illustrates mathematical formulas using Column-Vector notation. Other books like "3D Math Primer for
    ' Graphics and Game Development", Fletcher Dunn·Ian Parberry, use Row-Vector notation. Both are correct,
    ' however it's important to know which standard you code to, because it affects the way in which you
    ' build your matrices and the order in which you should multiply them to obtain the correct result.
    '
    ' OpenGL uses Column Vectors (like this application).
    ' DirectX uses Row Vectors.

End Function

