VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Application"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Define the name of this class/module for error-trap reporting.
Private Const mc_ModuleName As String = "Peters3DEngine8.Application"

' ============================
' Simulation State / Settings.
' ============================
Public GameState As String
Private m_blnVideoFullScreen As Boolean
Private m_blnDetailClipBoundary As Boolean
Private m_blnDetailPartialPolygons As Boolean
Private m_blnDetailPolygonsOutside As Boolean
Private m_blnDetailLCDScreen As Boolean
Private m_intDetailDrawMode As Integer
Private m_blnDetailEdgedFaces As Boolean
Private m_blnDetailDayMode As Boolean
Private m_blnDetailShadows As Boolean
Private m_intInputMouseSensitivity As Integer
Private m_blnIgnoreEvent As Boolean

' This is where most of the drawing occurs.
' =========================================
Public WithEvents m_frmCanvas As frmCanvas
Attribute m_frmCanvas.VB_VarHelpID = -1
'Private WithEvents m_frmSettings As frmSettings


' This big-fat array holds ALL 3D objects in this aplication!
' (Arrays are much faster than VB Classes & Collections)
' ===========================================================
Private m_AllObjects()          As mdr3DObject
Private m_Particles()           As mdr3DObject

Private m_Primitives()          As mdrPrimitive

Private m_Camera              As mdr3DTargetCamera
Private m_matViewOrientation    As mdrMatrix4
Private m_matViewMapping        As mdrMatrix4
Private m_matViewPort           As mdrMatrix4

Private m_PlayerOne             As mdrPlayer

Private m_blnMouseCapture As Boolean
Private m_RenderOptions As Integer ' Temporary rendering options.


Private Sub DoMasterReset()
    
    ' =======================================
    ' General Settings. Change these for fun!
    ' =======================================
    m_blnVideoFullScreen = True
    m_blnDetailClipBoundary = False
    m_blnDetailLCDScreen = False
    m_intDetailDrawMode = 0 ' << Wireframe, Flat, Gouraud.
    m_blnDetailEdgedFaces = False
    m_blnDetailDayMode = False
    m_blnDetailShadows = True
    m_intInputMouseSensitivity = 8
    m_blnDetailPartialPolygons = False
    
    ' ====================
    ' Setup form / canvas.
    ' ====================
    m_frmCanvas.ScaleMode = vbPixels
    
    ' ====================================
    ' Load Animation Scene from Text File.
    ' ====================================
''    Call ReadAnimationFile
    
    ' ==============================
    ' Reserve space for the objects.
    ' ==============================
    ReDim m_AllObjects(3) As mdr3DObject
    Randomize Now
    
    
    ' ==========================================
    ' This routine is used to load some objects.
    ' ==========================================
    Dim objTestGrid As mdr3DObject
    Dim obj3DStudio As mdr3DObject
    Dim objTree As mdr3DObject
    Dim objRock As mdr3DObject
              
    ' ================
    ' Load "TestGrid".
    ' ================
    objTestGrid = LoadGroundGrid()
    objTestGrid.ID = "TestGrid"
    objTestGrid.Class = "GeometryDots"
    objTestGrid.Enabled = True
    objTestGrid.Caption = "TestGrid"
    objTestGrid.Description = "TestGrid"
    objTestGrid.UniformScale = 1
    m_AllObjects(0) = objTestGrid

    ' Load "3DStudio" Text.
    ' =====================
    obj3DStudio = mDirectXParser2.LoadXFile(App.Path & "\xFiles\ufo_6b.X")
'    obj3DStudio = mDirectXParser2.LoadXFile(App.Path & "\xFiles\3DStudio_only.X")
    'obj3DStudio = mDirectXParser2.LoadXFile(App.Path & "\xFiles\dolphin.X")
    obj3DStudio.ID = "3DStudio_only.X"
    obj3DStudio.Class = "Geometry"
    obj3DStudio.Enabled = True
    obj3DStudio.Caption = "3DStudio_only.X"
    obj3DStudio.Description = "3DStudio_only.X"
    obj3DStudio.UniformScale = 5
    obj3DStudio.WorldPosition.Y = 16
    m_AllObjects(1) = obj3DStudio
    
    ' ======
    ' Trees.
    ' ======
    objTree = mDirectXParser2.LoadXFile(App.Path & "\xFiles\tree_1.X")
    objTree.ID = "objTree.X"
    objTree.Class = "Geometry"
    objTree.Enabled = True
    objTree.Caption = "objTree.X"
    objTree.Description = "objTree.X"
    objTree.UniformScale = 1
    
    Dim intN As Integer
    For intN = 2 To UBound(m_AllObjects)
        m_AllObjects(intN) = objTree
        m_AllObjects(intN).WorldPosition.X = (Rnd * 4000) - 2000
        m_AllObjects(intN).WorldPosition.Z = (Rnd * 4000) - 2000
        m_AllObjects(intN).WorldPosition.Y = 0
        m_AllObjects(intN).ID = "objTree.X"
        m_AllObjects(intN).Class = "tree"
        m_AllObjects(intN).Enabled = True
        m_AllObjects(intN).Caption = "objTree.X"
        m_AllObjects(intN).Description = "objTree.X"
        m_AllObjects(intN).UniformScale = Rnd * 5
        m_AllObjects(intN).Yaw = Rnd * 360
    Next intN

    
''    ' ======
''    ' Rocks.
''    ' ======
''    objRock = mDirectXParser2.LoadXFile(App.Path & "\xFiles\rock1.X")
''    objRock.ID = "objTree.X"
''    objRock.Class = "Geometry"
''    objRock.Enabled = True
''    objRock.Caption = "objRock.X"
''    objRock.Description = "objRock.X"
''    objRock.UniformScale = 1
''    For intN = 11 To UBound(m_AllObjects)
''        m_AllObjects(intN) = objRock
''        m_AllObjects(intN).WorldPosition.X = (Rnd * 3000) - 1500
''        m_AllObjects(intN).WorldPosition.Z = (Rnd * 3000) - 1500
''        m_AllObjects(intN).WorldPosition.Y = 0
''        m_AllObjects(intN).ID = "objRock.X"
''        m_AllObjects(intN).Class = "tree"
''        m_AllObjects(intN).Enabled = True
''        m_AllObjects(intN).Caption = "objRock.X"
''        m_AllObjects(intN).Description = "objRock.X"
''        m_AllObjects(intN).UniformScale = Rnd * 10
''        m_AllObjects(intN).Yaw = Rnd * 360
''    Next intN




     
    ReDim m_Particles(128) As mdr3DObject
    
    ' ===============================
    ' Reset Camera / Player Position.
    ' ===============================
    m_PlayerOne.WorldPosition.X = 0
    m_PlayerOne.WorldPosition.Y = 0
    m_PlayerOne.WorldPosition.Z = 256
    m_PlayerOne.VPN.X = 0
    m_PlayerOne.VPN.Y = 0
    m_PlayerOne.VPN.Z = -1
    
    m_Camera.FreeCamera = False
    
    m_Camera.WorldPosition.X = 0
    m_Camera.WorldPosition.Y = 0
    m_Camera.WorldPosition.Z = 256
    m_Camera.WorldPosition.w = 1
    
    m_Camera.LookAtPoint.X = 0
    m_Camera.LookAtPoint.Y = 0
    m_Camera.LookAtPoint.Z = 0
    m_Camera.LookAtPoint.w = 1
    
    m_Camera.VPN.X = 0
    m_Camera.VPN.Y = 0
    m_Camera.VPN.Z = 1
    m_Camera.VPN.w = 1
    
    m_Camera.VUP.X = 0
    m_Camera.VUP.Y = 1
    m_Camera.VUP.Z = 0
    m_Camera.VUP.w = 1#
    
    ' Remember: PRP is specified in Camera Coordinates.
    m_Camera.PRP.X = 0 ' u
    m_Camera.PRP.Y = 0 ' v
    m_Camera.PRP.Z = 2.666 ' n << Change this value to distort the perspective!
    m_Camera.PRP.Z = 1#
    m_Camera.PRP.w = 1#
    
    
    ' ==========================================================================
    ' The UV coordinate system is our 'virtual window' into the 3D world.
    ' --------------------------------------------------------------------------
    ' All objects in our scene are clipped to this 'virtual window'. After we
    ' have finished doing all the 3D calculations, we then transform everything
    ' within this 'virtual window' into the 'actual window'; this is our Windows
    ' Form. If the 'virtual window' is a perfect square, then we need to make
    ' sure the form we are drawing too is also a perfect square, otherwise we'll
    ' get an image that looks squashed. The normal aspect ratio of a computer
    ' monitor is 1:1.333  Some typicaly ratios are 640x480, 800x600, 1024x768
    '   Example: 1024 / 768 = 1.333
    ' For this reason, I often make the 'virtual window' the same shape as our
    ' form by making it the same aspect ratio.
    ' ==========================================================================
    m_Camera.Umin = (-1 + m_Camera.UPan) * 1.333
    m_Camera.Umax = (1 + m_Camera.UPan) * 1.333
    m_Camera.Vmin = -1 + m_Camera.VPan
    m_Camera.Vmax = 1 + m_Camera.VPan
    
    m_Camera.ClipNear = 0
    m_Camera.ClipFar = -8192 ' (any negative number)
        
    ' =================================================================
    ' The ViewPort (VP) bounds must be specified in pixels (not twips).
    ' -----------------------------------------------------------------
    ' Typical values are as follows:
    '       VPXmin = 0
    '       VPXmax = 1024
    '       VPYmin = 0
    '       VPYmax = 768
    ' Note: I usually put a 64 pixel border around this for testing
    '       the clipping procedure.
    ' =================================================================
    If m_blnVideoFullScreen = True Then
        m_Camera.VPXmin = m_frmCanvas.ScaleLeft
        m_Camera.VPXmax = m_frmCanvas.ScaleWidth
        m_Camera.VPYmin = m_frmCanvas.ScaleHeight
        m_Camera.VPYmax = m_frmCanvas.ScaleTop
    Else
        m_Camera.VPXmin = m_frmCanvas.ScaleLeft + 64
        m_Camera.VPXmax = m_frmCanvas.ScaleWidth - 64
        m_Camera.VPYmin = m_frmCanvas.ScaleHeight - 64
        m_Camera.VPYmax = m_frmCanvas.ScaleTop + 128
    End If
    
End Sub

Private Sub DoOutputStatistics(withHandle As Form, DayMode As Boolean)

    Dim lngTempOnly As Long
    Dim curTemp64BitInteger As Currency
    Dim lngColBright As Long
    Dim lngColDim As Long
    
    
    If DayMode = True Then
        lngColBright = RGB(0, 0, 255)
        lngColDim = RGB(160, 160, 255)
    Else
        lngColBright = RGB(255, 0, 0)
        lngColDim = RGB(128, 0, 0)
    End If
    
    
    ' Update debug & info. counters (conditionally compiled)
    ' NOTE: Putting this part in the display loop, isn't really a great idea, it slows things down.
    #If g_blnCompiledDebugInfo = True Then
    
        withHandle.Font = "Courier"
        withHandle.FontSize = 7
        withHandle.FontBold = False
        withHandle.ForeColor = lngColDim
        withHandle.CurrentY = withHandle.ScaleTop
        withHandle.CurrentX = withHandle.ScaleLeft

        lngTempOnly = g_lngMultiplicationCount + g_lngDivisionCount + g_lngAdditionCount + g_lngSubtractionCount + g_lngSquareRootCount
        If lngTempOnly = 0 Then lngTempOnly = 1
        
        withHandle.Print "Multiplications /frame : " & Format(g_lngMultiplicationCount, "###,000,000") & "  " & Format(g_lngMultiplicationCount / lngTempOnly, "00%")
        withHandle.Print "Divisions              : " & Format(g_lngDivisionCount, "###,000,000") & "  " & Format(g_lngDivisionCount / lngTempOnly, "00%")
        withHandle.Print "Additions              : " & Format(g_lngAdditionCount, "###,000,000") & "  " & Format(g_lngAdditionCount / lngTempOnly, "00%")
        withHandle.Print "Subtractions           : " & Format(g_lngSubtractionCount, "###,000,000") & "  " & Format(g_lngSubtractionCount / lngTempOnly, "00%")
        withHandle.Print "Square Roots           : " & Format(g_lngSquareRootCount, "###,000,000") & "  " & Format(g_lngSquareRootCount / lngTempOnly, "00%")
        withHandle.ForeColor = lngColBright
        withHandle.Print "Polygons Drawn         : " & Format(g_lngPolygonCount, "###,000,000")
        withHandle.ForeColor = lngColDim
        
        curTemp64BitInteger = (GetPerformanceCounter - g_curFramePerfCount) / g_curPerformanceFrequency
        withHandle.Print "Frame Draw Time (sec.) : " & curTemp64BitInteger
        
        withHandle.ForeColor = lngColBright
        withHandle.Print "Frame Rate (approx.)   : " & Format(1 / curTemp64BitInteger, "#,##0.0")
        
    #End If

End Sub

Private Sub DoUpdateSettingsWindow()

''    If (m_frmSettings Is Nothing) Then Exit Sub
''    If m_frmSettings.IsFormLoaded = False Then Exit Sub
''
''    With m_frmSettings
''        .txtCameraX.Text = m_Camera.WorldPosition.X
''        .txtCameraY.Text = m_Camera.WorldPosition.Y
''        .txtCameraZ.Text = m_Camera.WorldPosition.Z
''
''        .txtClipNear.Text = m_Camera.ClipNear
''        .txtClipFar.Text = m_Camera.ClipFar
''
''        .txtDocResWidth.Text = m_Camera.VPXmax - m_Camera.VPXmin
''        .txtDocResHeight.Text = m_Camera.VPYmin - m_Camera.VPYmax
''
''        .txtFOV.Text = m_Camera.FOV
''        .txtZoom.Text = m_Camera.Zoom
''        .txtPRPz.Text = m_Camera.PRP.Z
''
''    End With
    
End Sub

Private Sub PrintAnimationText(FontSize As Integer, Text As String)

    m_frmCanvas.Font = "Arial Bold"
    m_frmCanvas.FontSize = FontSize
    If m_blnDetailDayMode = True Then
        m_frmCanvas.ForeColor = vbBlack
    Else
        m_frmCanvas.ForeColor = vbWhite
    End If
    
    Dim sngTextWidth As Single
    Dim sngTextHeight As Single
    
    sngTextWidth = m_frmCanvas.TextWidth(Text)
    sngTextHeight = m_frmCanvas.TextHeight(Text)
    
    m_frmCanvas.CurrentX = (m_frmCanvas.ScaleWidth - sngTextWidth) / 2
    m_frmCanvas.CurrentY = (m_frmCanvas.ScaleHeight - sngTextHeight) / 2
    
    m_frmCanvas.Print Text
    
End Sub

Private Sub zTest()

    Dim intObject As Integer
    Dim intPart As Integer
    Dim intFace As Integer
    
    For intObject = LBound(m_AllObjects) To (UBound(m_AllObjects) - 1)
    
        For intPart = 0 To UBound(m_AllObjects(intObject).Parts)
            For intFace = 0 To UBound(m_AllObjects(intObject).Parts(intPart).Faces)
            
            Next intFace
        Next intPart
    Next intObject

End Sub

Private Function LoadGroundGrid() As mdr3DObject
    
    Dim sngX As Double
    Dim sngY As Double
    Dim sngZ As Double
    Dim lngVertexCount As Long
    
    ReDim LoadGroundGrid.Parts(0)
    
    Const sngMax = 4000!
    Const sngStep = 400!
    
    For sngZ = -sngMax To sngMax Step sngStep
        For sngX = -sngMax To sngMax Step sngStep
            For sngY = 0 To sngMax Step sngStep
                With LoadGroundGrid.Parts(0)
                    If (Abs(sngX) = sngMax) And (Abs(sngY) = sngMax) Then
'                        ReDim Preserve LoadGroundGrid.Parts(0).Vertices(lngVertexCount)
'                        .Vertices(lngVertexCount).Pxyz.X = sngX
'                        .Vertices(lngVertexCount).Pxyz.Y = sngY
'                        .Vertices(lngVertexCount).Pxyz.Z = sngZ
'                        .Vertices(lngVertexCount).Pxyz.w = 1
'                        .Vertices(lngVertexCount).RGB_Red = 255: .Vertices(lngVertexCount).RGB_Green = 127: .Vertices(lngVertexCount).RGB_Blue = 127
'                        lngVertexCount = lngVertexCount + 1
                    ElseIf (Abs(sngY) = sngMax) And (Abs(sngZ) = sngMax) Then
'                        ReDim Preserve LoadGroundGrid.Parts(0).Vertices(lngVertexCount)
'                        .Vertices(lngVertexCount).Pxyz.X = sngX
'                        .Vertices(lngVertexCount).Pxyz.Y = sngY
'                        .Vertices(lngVertexCount).Pxyz.Z = sngZ
'                        .Vertices(lngVertexCount).Pxyz.w = 1
'                        .Vertices(lngVertexCount).RGB_Red = 127: .Vertices(lngVertexCount).RGB_Green = 255: .Vertices(lngVertexCount).RGB_Blue = 127
'                        lngVertexCount = lngVertexCount + 1
                    ElseIf (Abs(sngX) = sngMax) And (Abs(sngZ) = sngMax) Then
'                        ReDim Preserve LoadGroundGrid.Parts(0).Vertices(lngVertexCount)
'                        .Vertices(lngVertexCount).Pxyz.X = sngX
'                        .Vertices(lngVertexCount).Pxyz.Y = sngY
'                        .Vertices(lngVertexCount).Pxyz.Z = sngZ
'                        .Vertices(lngVertexCount).Pxyz.w = 1
'                        .Vertices(lngVertexCount).RGB_Red = 127: .Vertices(lngVertexCount).RGB_Green = 127: .Vertices(lngVertexCount).RGB_Blue = 255
'                        lngVertexCount = lngVertexCount + 1
                    ElseIf (sngX = 0) Then
'                        ReDim Preserve LoadGroundGrid.Parts(0).Vertices(lngVertexCount)
'                        .Vertices(lngVertexCount).Pxyz.X = sngX
'                        .Vertices(lngVertexCount).Pxyz.Y = sngY
'                        .Vertices(lngVertexCount).Pxyz.Z = sngZ
'                        .Vertices(lngVertexCount).Pxyz.w = 1
'                        .Vertices(lngVertexCount).RGB_Red = 255: .Vertices(lngVertexCount).RGB_Green = 0: .Vertices(lngVertexCount).RGB_Blue = 0
'                        lngVertexCount = lngVertexCount + 1
                    ElseIf (sngY = 0) Then
                        ReDim Preserve LoadGroundGrid.Parts(0).Vertices(lngVertexCount)
                        .Vertices(lngVertexCount).Pxyz.X = Rnd * (sngMax * 2) - sngMax 'sngX
                        .Vertices(lngVertexCount).Pxyz.Y = sngY
                        .Vertices(lngVertexCount).Pxyz.Z = Rnd * (sngMax * 2) - sngMax 'sngZ
                        .Vertices(lngVertexCount).Pxyz.w = 1
                        .Vertices(lngVertexCount).RGB_Red = 0: .Vertices(lngVertexCount).RGB_Green = 255: .Vertices(lngVertexCount).RGB_Blue = 0
                        lngVertexCount = lngVertexCount + 1
                    ElseIf (sngZ = 0) Then
'                        ReDim Preserve LoadGroundGrid.Parts(0).Vertices(lngVertexCount)
'                        .Vertices(lngVertexCount).Pxyz.X = sngX
'                        .Vertices(lngVertexCount).Pxyz.Y = sngY
'                        .Vertices(lngVertexCount).Pxyz.Z = sngZ
'                        .Vertices(lngVertexCount).Pxyz.w = 1
'                        .Vertices(lngVertexCount).RGB_Red = 127: .Vertices(lngVertexCount).RGB_Green = 127: .Vertices(lngVertexCount).RGB_Blue = 255
'                        lngVertexCount = lngVertexCount + 1
                    End If
                End With
            Next sngY
        Next sngX
    Next sngZ
        
End Function

Private Sub RefreshF5()
    
    #If g_blnCompiledDebugInfo = True Then
        ' Reset debug/info. counters (conditionally compiled)
        g_lngMultiplicationCount = 0
        g_lngDivisionCount = 0
        g_lngAdditionCount = 0
        g_lngSubtractionCount = 0
        g_lngSquareRootCount = 0
        g_lngPolygonCount = 0
        g_curFramePerfCount = GetPerformanceCounter
    #End If
    
    ' =================
    ' Erase the screen.
    ' =================
'    If m_blnDetailDayMode = True Then
'        m_frmCanvas.BackColor = vbWhite
'    Else
'        m_frmCanvas.BackColor = vbBlack
'    End If
    
    
    ' =================================================
    ' Place a 128 pixel border around the viewing area.
    ' =================================================
    If m_blnVideoFullScreen = True Then
        m_Camera.VPXmin = m_frmCanvas.ScaleLeft
        m_Camera.VPXmax = m_frmCanvas.ScaleWidth
        m_Camera.VPYmin = m_frmCanvas.ScaleHeight
        m_Camera.VPYmax = m_frmCanvas.ScaleTop
    Else
        m_Camera.VPXmin = m_frmCanvas.ScaleLeft + 64
        m_Camera.VPXmax = m_frmCanvas.ScaleWidth - 64
        m_Camera.VPYmin = m_frmCanvas.ScaleHeight - 64
        m_Camera.VPYmax = m_frmCanvas.ScaleTop + 128
    End If
    
    
    m_Camera.WorldPosition = m_PlayerOne.WorldPosition
    'm_Camera.WorldPosition = m_Particles(0).WorldPosition
    
    Call DoSetDrawingParamaters(m_Camera)
    m_Camera.FOV = ConvertRad2Deg(2 * Atn((m_Camera.Umax - m_Camera.Umin) / m_Camera.PRP.Z))
    m_Camera.Zoom = Round(1 / Tan(ConvertDeg2Rad(m_Camera.FOV) / 2), 5)
    
    
    ' =======================================================
    ' Map canonical view volume into 3D viewport view volume.
    ' =======================================================
    With m_frmCanvas
        m_matViewPort = Matrix_vv3dv(m_frmCanvas, m_Camera.VPXmin, m_Camera.VPXmax, m_Camera.VPYmin, m_Camera.VPYmax, -1, 0, m_blnDetailClipBoundary, m_blnDetailDayMode)
    End With
    
    ' =====================================================================================
    ' Draw shadows first. Note: To draw shadows, you have to traverse through the render
    ' pipeline twice, thus slowing everything down and doubling the number of calculations.
    ' =====================================================================================
    If (m_blnDetailShadows = True) And (m_Camera.WorldPosition.Y > 0) Then
        Call RenderPipeline0(m_AllObjects, True)
        Call RenderPipeline0(m_Particles, True)
    End If
    
    ' ==============
    ' Draw normally.
    ' ==============
    Call RenderPipeline0(m_AllObjects, False)
    Call RenderPipeline0(m_Particles, False)
    
    
    Call RenderCrossHairs(m_frmCanvas, m_Camera, mdrCRS_DotOnly + mdrCRS_Scope1)
    
    ' ==========================================
    ' Output some text to the screen. (optional)
    ' ==========================================
'    Call DoOutputStatistics(m_frmCanvas, m_blnDetailDayMode)
    

    Screen.MousePointer = vbDefault

End Sub

Private Sub RenderPipeline0(p_3DObjects() As mdr3DObject, p_blnDrawShadows As Boolean)

'    On Error GoTo errTrap
    
    Dim intObject As Integer
    Dim lngPart As Long
    Dim matOutput As mdrMatrix4
    Dim matWorld As mdrMatrix4
    Dim blnClipMe As Boolean
    Dim sngW As Double
    Dim intJ As Integer
    Dim vectLight As mdrVector4
    Dim vectPlaneEq As mdrVector4
    Dim matShadow As mdrMatrix4
    
    
    For intObject = LBound(p_3DObjects) To UBound(p_3DObjects)
    
        With p_3DObjects(intObject)
            If .Enabled = True Then
            
                ' Draw each Part within the 3DObject.
                For lngPart = 0 To UBound(.Parts)
                
                    ' ====================
                    ' Init. Object Matrix.
                    ' ====================
                    Dim tempA As mdrVector4
                    tempA.X = 0
                    tempA.Y = 0
                    tempA.Z = 0
                    tempA.w = 1
                    
                    Dim matI As mdrMatrix4
                    
                    If .Class = "projectile" Then
'                        .Parts(lngPart).IdentityMatrix = MatrixInverse(MatrixLookAt(m_AllObjects(1).WorldPosition, .WorldPosition))
                    ElseIf .Class = "Geometry" Then
                        .Parts(lngPart).IdentityMatrix = MatrixIdentity()
                    Else
                        .Parts(lngPart).IdentityMatrix = MatrixIdentity()
                    End If
                    
                    matOutput = MakeWorldMatrix(.Parts(lngPart).IdentityMatrix, .WorldPosition, .Pitch, .Yaw, .Roll, .UniformScale)
                
                    ' ================
                    ' Process Shadows.
                    ' ================
                    If (p_blnDrawShadows = True) And (.Class <> "GeometryDots") Then
                    
                        vectLight.X = 8192
                        vectLight.Y = 8192
                        vectLight.Z = 16384
                        vectLight.w = 1

'                        vectLight = m_Particles(0).WorldPosition
                        'vectLight = VectorMultiplyByScalar(m_PlayerOne.WorldPosition, 1)
                        'vectLight.Y = Abs(vectLight.Y)
                        
                        
                        vectPlaneEq.X = 0
                        vectPlaneEq.Y = 1
                        vectPlaneEq.Z = 0
                        vectPlaneEq.w = 1
                                                
                        matShadow = MatrixShadow(vectLight, vectPlaneEq)
                    
                        matOutput = MatrixMultiply(matOutput, matShadow)
                        
                    End If ' Shadows?
                    
                    
                    ' =====================
                    ' Concatenate Matrices.
                    ' =====================
                    matOutput = MatrixMultiply(matOutput, m_matViewOrientation)
                    matOutput = MatrixMultiply(matOutput, m_matViewMapping)
                    
                    
                    ' For each vertex in the Part, recalculate.
                    For intJ = 0 To UBound(.Parts(lngPart).Vertices)
                        
                        With .Parts(lngPart).Vertices(intJ)
                            
                            ' Do some heavy number crunching here.
                            .Wxyz = MatrixMultiplyVector(matOutput, .Pxyz)
                            
                            .Clipped = False
                            sngW = .Wxyz.w
                            If sngW > 0 Then
                                ' W is positive.
                                If (.Wxyz.X < -sngW) Or (.Wxyz.X > sngW) Then .Clipped = True
                                If (.Wxyz.Y < -sngW) Or (.Wxyz.Y > sngW) Then .Clipped = True
                                If (.Wxyz.Z < -sngW) Or (.Wxyz.Z > 0) Then .Clipped = True
                            Else
                                ' W is negative
                                If (.Wxyz.X < -sngW) Or (.Wxyz.X > sngW) Then .Clipped = True
                                If (.Wxyz.Y < -sngW) Or (.Wxyz.Y > sngW) Then .Clipped = True
                                If (.Wxyz.Z < -sngW) Or (.Wxyz.Z > 0) Then .Clipped = True
                            End If
                            
                            If (.Clipped = False) Then
                                .Brightness = Abs(1 - .Wxyz.w)
                                If .Brightness > 1 Then .Brightness = 1
                            End If
                            
                        End With
                    Next intJ
                    
                    
                ' ==========================
                ' >>> Clip Polygons Here <<<
                ' ==========================
    ''            Call DoClipping(.Parts(lngPart))
                
                
SkipOver:
                Next lngPart
                
                ' Convert canonical view volume into 3D viewport view volume.
                ' ClipSpace -> ViewPort (ie. to pixels)
                On Error GoTo errTrap2
                For lngPart = 0 To UBound(.Parts)
                    For intJ = 0 To UBound(.Parts(lngPart).Vertices)
                        With .Parts(lngPart).Vertices(intJ)
                            .Txyz = MatrixMultiplyVector(m_matViewPort, .Wxyz)
                        End With
                    Next intJ
SkipOver2:
                Next lngPart
    
                
                ' Convert from 4D down to 3D
                ' ==========================
                Call Convert4Dto3D(p_3DObjects(intObject))
                
                
                ' ==============
                ' Draw 3D Object
                ' ==============
                Select Case p_3DObjects(intObject).Class
                    Case "GeometryDots"
                        Call Draw_DotsOnly(p_3DObjects(intObject))
                    
                    Case Else
                        Call Draw_SimpleWireframe(p_3DObjects(intObject), p_blnDrawShadows)
                        
                End Select
            
            End If ' Is object Enabled?
            
        End With
    Next intObject
    
    
    Exit Sub
errTrap:
    Resume SkipOver
    
errTrap2:
    
    Resume SkipOver2
    
End Sub


Private Sub RenderPipeline00()
   
   ' =======================================================================================
   ' -------------------[ Double-Zero is the default wire-frame renderer ]------------------
   '
   ' * Backfaces are included. Nothing complicated.
   ' * Worse-case rendering used to measure performance against improved rendering piplines.
   ' =======================================================================================
   On Error GoTo errTrap
   
    Dim intObject As Integer
    Dim lngPart As Long
    Dim lngFace As Long
    Dim matOutput As mdrMatrix4
    Dim matWorld As mdrMatrix4
    Dim matIWorld As mdrMatrix4
    Dim vectLight As mdrVector4
    Dim vectPlaneEq As mdrVector4
    Dim matShadow As mdrMatrix4
    Dim blnClipMe As Boolean
    Dim sngW As Double
    Dim intJ As Integer
    Dim vectCamera As mdrVector4
    Dim vectVertex As mdrVector4
    Dim lngIndex0 As Long, lngIndex1 As Long, lngIndex2 As Long
    Dim vect0 As mdrVector4, vect1 As mdrVector4, vect2 As mdrVector4
    Dim sngDotProduct As Single
    
    For intObject = LBound(m_AllObjects) To (UBound(m_AllObjects))
        With m_AllObjects(intObject)
            
            ' Draw each Part within the 3DObject.
            For lngPart = 0 To UBound(.Parts)
            
                matWorld = MakeWorldMatrix(.Parts(lngPart).IdentityMatrix, .WorldPosition, 0, 0, 0, 1)
'                matIWorld = MatrixInverse(matWorld)
'                vectCamera = MatrixMultiplyVector(matIWorld, m_Camera.WorldPosition)
                
                matOutput = MatrixMultiply(matWorld, m_matViewOrientation)
                matOutput = MatrixMultiply(matOutput, m_matViewMapping)
                matOutput = MatrixMultiply(matOutput, m_matViewPort)
                
                ' Draw each polygon for the given polyhedra.
                For lngFace = 0 To UBound(.Parts(lngPart).Faces)
                
                    lngIndex0 = .Parts(lngPart).Faces(lngFace)(0)
                    lngIndex1 = .Parts(lngPart).Faces(lngFace)(1)
                    lngIndex2 = .Parts(lngPart).Faces(lngFace)(2)
                    
                    vect0 = .Parts(lngPart).Vertices(lngIndex0).Pxyz
                    vect1 = .Parts(lngPart).Vertices(lngIndex1).Pxyz
                    vect2 = .Parts(lngPart).Vertices(lngIndex2).Pxyz
                    
                    .Parts(lngPart).Vertices(lngIndex0).Txyz = MatrixMultiplyVector(matOutput, .Parts(lngPart).Vertices(lngIndex0).Pxyz)
                    .Parts(lngPart).Vertices(lngIndex1).Txyz = MatrixMultiplyVector(matOutput, .Parts(lngPart).Vertices(lngIndex1).Pxyz)
                    .Parts(lngPart).Vertices(lngIndex2).Txyz = MatrixMultiplyVector(matOutput, .Parts(lngPart).Vertices(lngIndex2).Pxyz)
                    
                    .Parts(lngPart).Vertices(lngIndex0).Txyz.X = .Parts(lngPart).Vertices(lngIndex0).Txyz.X / .Parts(lngPart).Vertices(lngIndex0).Txyz.w
                    .Parts(lngPart).Vertices(lngIndex0).Txyz.Y = .Parts(lngPart).Vertices(lngIndex0).Txyz.Y / .Parts(lngPart).Vertices(lngIndex0).Txyz.w
                    .Parts(lngPart).Vertices(lngIndex0).Txyz.Z = .Parts(lngPart).Vertices(lngIndex0).Txyz.Z / .Parts(lngPart).Vertices(lngIndex0).Txyz.w
                    
                    .Parts(lngPart).Vertices(lngIndex1).Txyz.X = .Parts(lngPart).Vertices(lngIndex1).Txyz.X / .Parts(lngPart).Vertices(lngIndex1).Txyz.w
                    .Parts(lngPart).Vertices(lngIndex1).Txyz.Y = .Parts(lngPart).Vertices(lngIndex1).Txyz.Y / .Parts(lngPart).Vertices(lngIndex1).Txyz.w
                    .Parts(lngPart).Vertices(lngIndex1).Txyz.Z = .Parts(lngPart).Vertices(lngIndex1).Txyz.Z / .Parts(lngPart).Vertices(lngIndex1).Txyz.w
                    
                    .Parts(lngPart).Vertices(lngIndex2).Txyz.X = .Parts(lngPart).Vertices(lngIndex2).Txyz.X / .Parts(lngPart).Vertices(lngIndex2).Txyz.w
                    .Parts(lngPart).Vertices(lngIndex2).Txyz.Y = .Parts(lngPart).Vertices(lngIndex2).Txyz.Y / .Parts(lngPart).Vertices(lngIndex2).Txyz.w
                    .Parts(lngPart).Vertices(lngIndex2).Txyz.Z = .Parts(lngPart).Vertices(lngIndex2).Txyz.Z / .Parts(lngPart).Vertices(lngIndex2).Txyz.w
                    
                    vect0 = .Parts(lngPart).Vertices(lngIndex0).Txyz
                    vect1 = .Parts(lngPart).Vertices(lngIndex1).Txyz
                    vect2 = .Parts(lngPart).Vertices(lngIndex2).Txyz
                    
                    Call DrawFlatTriangle(m_frmCanvas, vect0.X, vect0.Y, vect1.X, vect1.Y, vect2.X, vect2.Y, RGB(0, 0, 255), -1, vbSolid, vbFSSolid)
                    g_lngPolygonCount = g_lngPolygonCount + 1
                    
                Next lngFace
ResumePoint:
            Next lngPart
        End With
    Next intObject
    
    Exit Sub
errTrap:
    Resume ResumePoint
   
End Sub
Private Sub Draw_SimpleWireframe(withObject As mdr3DObject, p_blnDrawShadows As Boolean)
    
    ' Displays 3D usin Simple Wireframe

    On Error GoTo errTrap

    Dim intN As Long, intJ As Long, intK As Long
    Dim blnClipped As Boolean
    Dim intClipCount As Integer

    Dim lngIndex0 As Long
    Dim vect0 As mdrVector4
    Dim Red0 As Double
    Dim Green0 As Double
    Dim Blue0 As Double
    Dim sngBrightness0 As Double
    
    Dim lngIndex1 As Long
    Dim vect1 As mdrVector4
    Dim Red1 As Double
    Dim Green1 As Double
    Dim Blue1 As Double
    Dim sngBrightness1 As Double
    
    Dim lngIndex2 As Long
    Dim vect2 As mdrVector4
    Dim Red2 As Double
    Dim Green2 As Double
    Dim Blue2 As Double
    Dim sngBrightness2 As Double

    
    If m_blnDetailLCDScreen = True Then
        m_frmCanvas.DrawWidth = 2
    Else
        m_frmCanvas.DrawWidth = 1
    End If
    If m_intDetailDrawMode <> 0 Then m_frmCanvas.DrawWidth = 1
    
    
    With withObject

        ' Draw each object's polyhedra
        For intN = 0 To UBound(.Parts)
            m_frmCanvas.ForeColor = vbBlack
            
            ' Draw each polygon for the given polyhedra.
            For intJ = 0 To UBound(.Parts(intN).Faces)

                If .Parts(intN).Selected = True Then
                    m_frmCanvas.ForeColor = vbWhite
                End If

                With .Parts(intN)
                
                    lngIndex0 = .Faces(intJ)(0)
                    lngIndex1 = .Faces(intJ)(1)
                    lngIndex2 = .Faces(intJ)(2)
                    
                    ' Enough faces for a Triangle?
                    If (UBound(.Faces(intJ)) > 1) Then
                    
                        ' Should the triangle be clipped?
                        blnClipped = False
                        blnClipped = blnClipped Or .Vertices(lngIndex0).Clipped
                        blnClipped = blnClipped Or .Vertices(lngIndex1).Clipped
                        blnClipped = blnClipped Or .Vertices(lngIndex2).Clipped
                        
                        intClipCount = 0
                        If .Vertices(lngIndex0).Clipped = True Then intClipCount = intClipCount + 1
                        If .Vertices(lngIndex1).Clipped = True Then intClipCount = intClipCount + 1
                        If .Vertices(lngIndex2).Clipped = True Then intClipCount = intClipCount + 1
                        
                        If blnClipped = True Then
                            vect0 = .Vertices(lngIndex0).Txyz
                            vect1 = .Vertices(lngIndex1).Txyz
                            vect2 = .Vertices(lngIndex2).Txyz
                            If intClipCount = 3 Then
                                If m_blnDetailPolygonsOutside = True Then
                                    Call DrawFlatTriangle(m_frmCanvas, vect0.X, vect0.Y, vect1.X, vect1.Y, vect2.X, vect2.Y, RGB(64, 32, 32), 0, vbSolid, vbFSTransparent)
                                End If
                            Else
                                If m_blnDetailPartialPolygons = True Then
                                    Call DrawFlatTriangle(m_frmCanvas, vect0.X, vect0.Y, vect1.X, vect1.Y, vect2.X, vect2.Y, RGB(255, 128, 0), 0, vbSolid, vbFSTransparent)
                                End If
                            End If
                            
                        Else
                            vect0 = .Vertices(lngIndex0).Txyz
                            vect1 = .Vertices(lngIndex1).Txyz
                            vect2 = .Vertices(lngIndex2).Txyz
                            
                            ' ==================
                            ' Back-Face Culling.
                            ' ==================
                            Dim vectA As mdrVector4, vectB As mdrVector4, vectC As mdrVector4
                            vectA = .Vertices(lngIndex0).Txyz
                            vectB = .Vertices(lngIndex1).Txyz
                            vectC = .Vertices(lngIndex2).Txyz

                            Dim sngDotProduct As Single
                            sngDotProduct = IsBackFace(m_Camera.PRP, vectA, vectB, vectC)
                            If sngDotProduct < 0 Then
                                
' Update debug & info. counters (conditionally compiled)
#If g_blnCompiledDebugInfo = True Then
    g_lngPolygonCount = g_lngPolygonCount + 1
#End If

                                If p_blnDrawShadows = False Then
                                
                                    ' ============================================================================================
                                    ' This colour part is much more complicated that it needs to be. You could probably delete all
                                    ' this colour stuff and it would not effect the program at all.
                                    ' ============================================================================================
                                    Dim sngRed0 As Single, sngGreen0 As Single, sngBlue0 As Single
                                    sngBrightness0 = .Vertices(lngIndex0).Brightness
                                    If m_blnDetailDayMode = True Then
                                        Call HSV2(sngRed0, sngGreen0, sngBlue0, 360 * sngBrightness0 * 6, sngBrightness0, 1)
                                    Else
                                        Call HSV2(sngRed0, sngGreen0, sngBlue0, 360 * sngBrightness0 * 6, 1, sngBrightness0)
                                    End If
                                    Red0 = 255 * sngRed0
                                    Green0 = 255 * sngGreen0
                                    Blue0 = 255 * sngBlue0

                                    Dim sngRed1 As Single, sngGreen1 As Single, sngBlue1 As Single
                                    sngBrightness1 = .Vertices(lngIndex1).Brightness
                                    If m_blnDetailDayMode = True Then
                                        Call HSV2(sngRed1, sngGreen1, sngBlue1, 360 * sngBrightness1 * 6, sngBrightness1, 1)
                                    Else
                                        Call HSV2(sngRed1, sngGreen1, sngBlue1, 360 * sngBrightness1 * 6, 1, sngBrightness1)
                                    End If
                                    Red1 = 255 * sngRed1
                                    Green1 = 255 * sngGreen1
                                    Blue1 = 255 * sngBlue1

                                    Dim sngRed2 As Single, sngGreen2 As Single, sngBlue2 As Single
                                    sngBrightness2 = .Vertices(lngIndex2).Brightness
                                    If m_blnDetailDayMode = True Then
                                        Call HSV2(sngRed2, sngGreen2, sngBlue2, 360 * sngBrightness2 * 6, sngBrightness2, 1)
                                    Else
                                        Call HSV2(sngRed2, sngGreen2, sngBlue2, 360 * sngBrightness2 * 6, 1, sngBrightness2)
                                    End If
                                    Red2 = 255 * sngRed2
                                    Green2 = 255 * sngGreen2
                                    Blue2 = 255 * sngBlue2
                                    
                                    Dim sngRedAvg As Single
                                    Dim sngGrnAvg As Single
                                    Dim sngBluAvg As Single
                                    sngRedAvg = (Red0 + Red1 + Red2) / 3
                                    sngGrnAvg = (Green0 + Green1 + Green2) / 3
                                    sngBluAvg = (Blue0 + Blue1 + Blue2) / 3
                                    
                                    ' Draw Wireframe.
                                    If m_intDetailDrawMode = 0 Then
                                        Call DrawFlatTriangle(m_frmCanvas, vect0.X, vect0.Y, vect1.X, vect1.Y, vect2.X, vect2.Y, RGB(sngRedAvg * 0.9, sngGrnAvg * 0.9, sngBluAvg * 0.9), RGB(sngRedAvg, sngGrnAvg, sngBluAvg), vbSolid, vbFSTransparent)
                                    End If
                                    
                                    ' Flat-Shaded Polygon
                                    If m_intDetailDrawMode = 1 Then
                                        If m_blnDetailEdgedFaces = True Then
                                            Call DrawFlatTriangle(m_frmCanvas, vect0.X, vect0.Y, vect1.X, vect1.Y, vect2.X, vect2.Y, RGB(sngRedAvg * 0.9, sngGrnAvg * 0.9, sngBluAvg * 0.9), RGB(sngRedAvg, sngGrnAvg, sngBluAvg), vbSolid, vbFSSolid)
                                        Else
                                            Call DrawFlatTriangle(m_frmCanvas, vect0.X, vect0.Y, vect1.X, vect1.Y, vect2.X, vect2.Y, 0, RGB(sngRedAvg, sngGrnAvg, sngBluAvg), vbInvisible, vbFSSolid)
                                        End If
                                    End If
                                    
                                    ' Gouraud Shaded
                                    If m_intDetailDrawMode = 2 Then
                                        Call DrawTriangle(m_frmCanvas, vect0.X, vect0.Y, vect1.X, vect1.Y, vect2.X, vect2.Y, Red0, Green0, Blue0, Red1, Green1, Blue1, Red2, Green2, Blue2)
                                        If m_blnDetailEdgedFaces = True Then
                                            Call DrawFlatTriangle(m_frmCanvas, vect0.X, vect0.Y, vect1.X, vect1.Y, vect2.X, vect2.Y, RGB(sngRedAvg * 0.9, sngGrnAvg * 0.9, sngBluAvg * 0.9), RGB(sngRedAvg, sngGrnAvg, sngBluAvg), vbSolid, vbFSTransparent)
                                        End If
                                    End If
                                    
                                Else
                                    ' Draw Shadow..
                                    If m_intDetailDrawMode <> 0 Then ' if wireframe
                                        Call DrawFlatTriangle(m_frmCanvas, vect0.X, vect0.Y, vect1.X, vect1.Y, vect2.X, vect2.Y, RGB(1, 1, 1), RGB(Blue0, Blue0, Blue0), vbSolid, vbFSSolid)
                                    Else
                                        Call DrawFlatTriangle(m_frmCanvas, vect0.X, vect0.Y, vect1.X, vect1.Y, vect2.X, vect2.Y, RGB(1, 1, 1), RGB(Blue0, Blue0, Blue0), vbSolid, vbFSTransparent)
                                    End If
                                End If
                            Else
                                'Call DrawFlatTriangle(m_frmCanvas, vect0.X, vect0.Y, vect1.X, vect1.Y, vect2.X, vect2.Y, RGB(0, 128, 0), -1, vbDot, vbFSSolid)
                            End If
                            
                        End If ' Clipped=False
                    End If ' Enough faces for a Triangle?
                End With
SkipOver1:
            Next intJ
        Next intN
    End With
    
    Exit Sub
errTrap:
    Select Case Err.Number
        Case 9 ' Subscript Error
            Resume SkipOver1

        Case Else
            MsgBox Err.Number & " - " & Err.Description, vbCritical
            ' Do nothing

    End Select
    
End Sub

Private Sub Draw_SimpleWireframe2(withObject As mdr3DObject, p_blnDrawShadows As Boolean)
    
    ' Displays 3D usin Simple Wireframe

    On Error GoTo errTrap

    Dim intN As Long, intJ As Long, intK As Long
    Dim blnClipped As Boolean
    Dim intClipCount As Integer

    Dim lngIndex0 As Long
    Dim vect0 As mdrVector4
    Dim Red0 As Double
    Dim Green0 As Double
    Dim Blue0 As Double
    Dim sngBrightness0 As Double
    
    Dim lngIndex1 As Long
    Dim vect1 As mdrVector4
    Dim Red1 As Double
    Dim Green1 As Double
    Dim Blue1 As Double
    Dim sngBrightness1 As Double
    
    Dim lngIndex2 As Long
    Dim vect2 As mdrVector4
    Dim Red2 As Double
    Dim Green2 As Double
    Dim Blue2 As Double
    Dim sngBrightness2 As Double

    
    If m_blnDetailLCDScreen = True Then
        m_frmCanvas.DrawWidth = 2
    Else
        m_frmCanvas.DrawWidth = 1
    End If
    If m_intDetailDrawMode <> 0 Then m_frmCanvas.DrawWidth = 1
    
    
    With withObject

        ' Draw each object's polyhedra
        For intN = 0 To UBound(.Parts)
            m_frmCanvas.ForeColor = vbBlack
            
            ' Draw each polygon for the given polyhedra.
            For intJ = 0 To UBound(.Parts(intN).Faces)

                If .Parts(intN).Selected = True Then
                    m_frmCanvas.ForeColor = vbWhite
                End If

                With .Parts(intN)
                
                    lngIndex0 = .Faces(intJ)(0)
                    lngIndex1 = .Faces(intJ)(1)
                    lngIndex2 = .Faces(intJ)(2)
                    
                    ' Enough faces for a Triangle?
                    If (UBound(.Faces(intJ)) > 1) Then
                    
                        ' Should the triangle be clipped?
                        blnClipped = False
                        blnClipped = blnClipped Or .Vertices(lngIndex0).Clipped
                        blnClipped = blnClipped Or .Vertices(lngIndex1).Clipped
                        blnClipped = blnClipped Or .Vertices(lngIndex2).Clipped
                        
                        intClipCount = 0
                        If .Vertices(lngIndex0).Clipped = True Then intClipCount = intClipCount + 1
                        If .Vertices(lngIndex1).Clipped = True Then intClipCount = intClipCount + 1
                        If .Vertices(lngIndex2).Clipped = True Then intClipCount = intClipCount + 1
                        
                        If blnClipped = True Then
                            vect0 = .Vertices(lngIndex0).Txyz
                            vect1 = .Vertices(lngIndex1).Txyz
                            vect2 = .Vertices(lngIndex2).Txyz
                            If intClipCount = 3 Then
                                If m_blnDetailPolygonsOutside = True Then
                                    Call DrawFlatTriangle(m_frmCanvas, vect0.X, vect0.Y, vect1.X, vect1.Y, vect2.X, vect2.Y, RGB(64, 32, 32), 0, vbSolid, vbFSTransparent)
                                End If
                            Else
                                If m_blnDetailPartialPolygons = True Then
                                    Call DrawFlatTriangle(m_frmCanvas, vect0.X, vect0.Y, vect1.X, vect1.Y, vect2.X, vect2.Y, RGB(255, 128, 0), 0, vbSolid, vbFSTransparent)
                                End If
                            End If
                            
                        Else
                            vect0 = .Vertices(lngIndex0).Txyz
                            vect1 = .Vertices(lngIndex1).Txyz
                            vect2 = .Vertices(lngIndex2).Txyz
                            
                            ' ==================
                            ' Back-Face Culling.
                            ' ==================
                            Dim vectA As mdrVector4, vectB As mdrVector4, vectC As mdrVector4
                            vectA = .Vertices(lngIndex0).Txyz
                            vectB = .Vertices(lngIndex1).Txyz
                            vectC = .Vertices(lngIndex2).Txyz

                            Dim sngDotProduct As Single
                            sngDotProduct = IsBackFace(m_Camera.PRP, vectA, vectB, vectC)
                            If sngDotProduct < 0 Then
                                
' Update debug & info. counters (conditionally compiled)
#If g_blnCompiledDebugInfo = True Then
    g_lngPolygonCount = g_lngPolygonCount + 1
#End If

                                If p_blnDrawShadows = False Then
                                
                                    ' ============================================================================================
                                    ' This colour part is much more complicated that it needs to be. You could probably delete all
                                    ' this colour stuff and it would not effect the program at all.
                                    ' ============================================================================================
                                    Dim sngRed0 As Single, sngGreen0 As Single, sngBlue0 As Single
                                    sngBrightness0 = .Vertices(lngIndex0).Brightness
                                    Red0 = .Vertices(lngIndex0).RGB_Red
                                    Green0 = .Vertices(lngIndex0).RGB_Green
                                    Blue0 = .Vertices(lngIndex0).RGB_Blue

                                    Dim sngRed1 As Single, sngGreen1 As Single, sngBlue1 As Single
                                    sngBrightness1 = .Vertices(lngIndex1).Brightness
                                    Red1 = .Vertices(lngIndex1).RGB_Red
                                    Green1 = .Vertices(lngIndex1).RGB_Green
                                    Blue1 = .Vertices(lngIndex1).RGB_Blue

                                    Dim sngRed2 As Single, sngGreen2 As Single, sngBlue2 As Single
                                    sngBrightness2 = .Vertices(lngIndex2).Brightness
                                    Red2 = .Vertices(lngIndex2).RGB_Red
                                    Green2 = .Vertices(lngIndex2).RGB_Green
                                    Blue2 = .Vertices(lngIndex2).RGB_Blue
                                    
                                    Dim sngRedAvg As Single
                                    Dim sngGrnAvg As Single
                                    Dim sngBluAvg As Single
                                    sngRedAvg = (Red0 + Red1 + Red2) / 3
                                    sngGrnAvg = (Green0 + Green1 + Green2) / 3
                                    sngBluAvg = (Blue0 + Blue1 + Blue2) / 3
                                    
                                    ' Draw Wireframe.
                                    If m_intDetailDrawMode = 0 Then
                                        Call DrawFlatTriangle(m_frmCanvas, vect0.X, vect0.Y, vect1.X, vect1.Y, vect2.X, vect2.Y, RGB(sngRedAvg * 0.9, sngGrnAvg * 0.9, sngBluAvg * 0.9), RGB(sngRedAvg, sngGrnAvg, sngBluAvg), vbSolid, vbFSTransparent)
                                    End If
                                    
                                    ' Flat-Shaded Polygon
                                    If m_intDetailDrawMode = 1 Then
                                        If m_blnDetailEdgedFaces = True Then
                                            Call DrawFlatTriangle(m_frmCanvas, vect0.X, vect0.Y, vect1.X, vect1.Y, vect2.X, vect2.Y, RGB(sngRedAvg * 0.9, sngGrnAvg * 0.9, sngBluAvg * 0.9), RGB(sngRedAvg, sngGrnAvg, sngBluAvg), vbSolid, vbFSSolid)
                                        Else
                                            Call DrawFlatTriangle(m_frmCanvas, vect0.X, vect0.Y, vect1.X, vect1.Y, vect2.X, vect2.Y, 0, RGB(sngRedAvg, sngGrnAvg, sngBluAvg), vbInvisible, vbFSSolid)
                                        End If
                                    End If
                                    
                                    ' Gouraud Shaded
                                    If m_intDetailDrawMode = 2 Then
                                        Call DrawTriangle(m_frmCanvas, vect0.X, vect0.Y, vect1.X, vect1.Y, vect2.X, vect2.Y, Red0, Green0, Blue0, Red1, Green1, Blue1, Red2, Green2, Blue2)
                                        If m_blnDetailEdgedFaces = True Then
                                            Call DrawFlatTriangle(m_frmCanvas, vect0.X, vect0.Y, vect1.X, vect1.Y, vect2.X, vect2.Y, RGB(sngRedAvg * 0.9, sngGrnAvg * 0.9, sngBluAvg * 0.9), RGB(sngRedAvg, sngGrnAvg, sngBluAvg), vbSolid, vbFSTransparent)
                                        End If
                                    End If
                                    
                                Else
                                    ' Draw Shadow..
                                    If m_intDetailDrawMode <> 0 Then ' if wireframe
                                        Call DrawFlatTriangle(m_frmCanvas, vect0.X, vect0.Y, vect1.X, vect1.Y, vect2.X, vect2.Y, RGB(1, 1, 1), RGB(Blue0, Blue0, Blue0), vbSolid, vbFSSolid)
                                    Else
                                        Call DrawFlatTriangle(m_frmCanvas, vect0.X, vect0.Y, vect1.X, vect1.Y, vect2.X, vect2.Y, RGB(1, 1, 1), RGB(Blue0, Blue0, Blue0), vbSolid, vbFSTransparent)
                                    End If
                                End If
                            Else
                                'Call DrawFlatTriangle(m_frmCanvas, vect0.X, vect0.Y, vect1.X, vect1.Y, vect2.X, vect2.Y, RGB(0, 128, 0), -1, vbDot, vbFSSolid)
                            End If
                            
                        End If ' Clipped=False
                    End If ' Enough faces for a Triangle?
                End With
SkipOver1:
            Next intJ
        Next intN
    End With
    
    Exit Sub
errTrap:
    Select Case Err.Number
        Case 9 ' Subscript Error
            Resume SkipOver1

        Case Else
            MsgBox Err.Number & " - " & Err.Description, vbCritical
            ' Do nothing

    End Select
    
End Sub
Private Sub Draw_DotsOnly(withObject As mdr3DObject)
    
    Dim intPart As Long, intVertex As Long
    Dim blnClipped As Boolean
    Dim intClipCount As Integer

    Dim lngIndex0 As Long
    Dim vect0 As mdrVector4
    Dim Red0 As Double
    Dim Green0 As Double
    Dim Blue0 As Double
    Dim sngBrightness0 As Double
    
    Dim sngX As Single
    Dim sngY As Single
    
    If m_blnDetailLCDScreen = True Then
        m_frmCanvas.DrawWidth = 3
    Else
        m_frmCanvas.DrawWidth = 1
    End If
    
    With withObject

        ' Draw each object's part.
        For intPart = 0 To UBound(.Parts)
        
            ' Loop through all vertices for the part.
            For intVertex = 0 To UBound(.Parts(intPart).Vertices)
            
                If .Parts(intPart).Vertices(intVertex).Clipped = False Then
                
                    If m_blnDetailDayMode = True Then
                        sngBrightness0 = 1 - .Parts(intPart).Vertices(intVertex).Brightness
                    Else
                        sngBrightness0 = .Parts(intPart).Vertices(intVertex).Brightness
                    End If
                    
                    Red0 = .Parts(intPart).Vertices(intVertex).RGB_Red * sngBrightness0
                    Green0 = .Parts(intPart).Vertices(intVertex).RGB_Green * sngBrightness0
                    Blue0 = .Parts(intPart).Vertices(intVertex).RGB_Blue * sngBrightness0
                    
                    sngX = .Parts(intPart).Vertices(intVertex).Txyz.X
                    sngY = .Parts(intPart).Vertices(intVertex).Txyz.Y
                    
                    m_frmCanvas.PSet (sngX, sngY), RGB(Red0, Green0, Blue0)
                    
                End If
                
            Next intVertex
        Next intPart
        
    End With
        
End Sub

Private Sub Convert4Dto3D(p_3DObject As mdr3DObject)

    On Error GoTo errTrap
    
    Dim lngN As Long, lngJ As Long
    
    With p_3DObject
    
        ' For each object, calculate it's Parts
        For lngN = 0 To UBound(.Parts)
        
            For lngJ = 0 To UBound(.Parts(lngN).Vertices)
                With .Parts(lngN).Vertices(lngJ)
                        
                        ' Ignore invalid co-ordinates
                        If (.Txyz.w = 0) = False Then
                            ' Convert from 4D down to 3D
                            ' ==========================
                            ' Apply the Perspective transformation by converting 4 dimensions down to
                            ' 3 dimensions, by dividing the x, y & z co-ordinates with w. Don't forget the w
                            ' value was previous calculated using the 'MatrixPerspective' routine.
                            ' I personally like to call this part of the code, 'dimensionally downshifting'.
                            .Txyz.X = .Txyz.X / .Txyz.w
                            .Txyz.Y = .Txyz.Y / .Txyz.w
                            .Txyz.Z = .Txyz.Z / .Txyz.w
                            
' Update debug & info. counters (conditionally compiled)
#If g_blnCompiledDebugInfo = True Then
    ' Probably a bad idea to have this code here in a loop... oh well.
    g_lngDivisionCount = g_lngDivisionCount + 3
#End If
                            
                    End If
                End With
SkipPointA:
            Next lngJ ' .Parts(lngN).Vertices
        Next lngN ' .Parts
    End With ' p_3DObject
    
    
    Exit Sub
errTrap:
    Select Case Err.Number
        Case 9 ' Subscript out of range
            Resume SkipPointA
            
        Case Else
            MsgBox Err.Number & " - " & Err.Description, vbExclamation
            
    End Select
    
End Sub

Private Sub DoSetDrawingParamaters(p_Camera As mdr3DTargetCamera)
    
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
    
    
    ' ====================================================
    ' Define the View Reference Point (VRP)
    ' This is defined in the World Coordinate (WC) system.
    ' ====================================================
    vectVRP = p_Camera.WorldPosition
    
    
    If p_Camera.FreeCamera = True Then
        ' Programmer is responsible for creating and rotating View Plane Normal (VPN) manually.
        ' This is a last minute hack... I'll probably clean this up a little so that I'm not
        ' mixing up cameras with human players; their code should be kept separate. Note, this
        ' is a 'camera' routine, not a 'human player' routine.
        vectVPN = m_PlayerOne.VPN
        'vectVPN = VectorMultiplyByScalar(m_Particles(0).Vector, -1)
        
    Else
        ' =============================================================================================================
        ' Subtract the Camera's world position (VRP) from the 'LookingAt' point to give us the View Plane Normal (VPN).
        ' In this application, the VPN points in the opposite direction that the camera is facing!! I said, Opposite!
        ' Fig. 6.14 to 6.16
        ' =============================================================================================================
        vectVPN = VectorSubtract(vectVRP, p_Camera.LookAtPoint)
    End If
    If (vectVPN.X = 0) And (vectVPN.Y = 0) And (vectVPN.Z = 0) Then
        vectVPN.X = 0# ' Do not allow VPN to be all zero's (shouldn't happen anyway, but still check)
        vectVPN.Y = 0#
        vectVPN.Z = 1#
    End If
    
    
    ' ==============================================================================================
    ' PRP is specified in the View Reference Coordinate system (and NOT the world coordinate system)
    ' ==============================================================================================
    vectPRP = p_Camera.PRP
    
    
    ' ==================================================================================
    ' The VUP vector is usually x=0,y=1,z=0. This is used to tilt the camera.
    ' If you had a camera firmly mounted on the handle bars of a motorcycle, you would
    ' expect to see the camera tilt when the motorcycle rounded the corners.
    ' (To achieve this, multiply the vectVUP vector with a rotation matrix)
    ' ==================================================================================
    vectVUP = p_Camera.VUP
    
    ' ============================================================================
    ' View Orientation.
    ' ============================================================================
    m_matViewOrientation = MatrixViewOrientation(vectVPN, vectVUP, vectVRP)
    
    
    ' ============================================================================
    ' Map projection view volume into canonical view volume suitable for clipping.
    ' ============================================================================
    m_matViewMapping = MatrixViewMapping_Per(p_Camera)
    
    
End Sub
Public Sub ShowApplication()
    
    ' Create the Canvas window (which is now 'owned' by the Application class.)
    If (m_frmCanvas Is Nothing) Then
    
        Set m_frmCanvas = New frmCanvas
        Load m_frmCanvas
        m_frmCanvas.Show
        
    End If
    
End Sub


Private Sub m_frmCanvas_OnAnimate()
    
    ' ===========================================================================
    ' The entire program is controlled from here.
    ' --------------------------------------------
    ' The game is always in a known state (ie. the GameState). If you program in
    ' this manner, you will be able to expand your game as time goes on without
    ' the headache of increasing the complexity (a common problem with newbie's).
    ' Each "game state" can have it's own "sub-state". That is, when the game is
    ' in "run_game" state, I could then sub-divide this state into various parts
    ' such as "do enemies", "do A.I", "do check for collisions" and so on.
    ' ===========================================================================
    Select Case Me.GameState
        Case ""
            ' ==========================================
            ' Master Reset - first time program has run.
            ' ==========================================
            Call DoMasterReset
            Call GetMouseInput(m_intInputMouseSensitivity, Me.GameState, m_PlayerOne, m_Camera, m_frmCanvas)
            
            Me.GameState = "run_game"

        
        Case "show_settings"
''            Me.GameState = "settings"
''            If m_frmSettings Is Nothing Then Set m_frmSettings = New frmSettings
''            m_frmSettings.Show vbModeless, m_frmCanvas
''
''            m_frmSettings.chkVideoFullScreen.Value = m_blnVideoFullScreen
''            m_frmSettings.chkDetailClipBoundry.Value = m_blnDetailClipBoundary
''            m_frmSettings.chkDetailEdgedFaces.Value = m_blnDetailEdgedFaces
''            m_frmSettings.chkDetailDayMode.Value = m_blnDetailDayMode
''            m_frmSettings.chkDetailShadows.Value = m_blnDetailShadows
''            m_frmSettings.chkDetailPartialPolygons.Value = m_blnDetailPartialPolygons
''            m_frmSettings.chkDetailPolygonsOutside.Value = m_blnDetailPolygonsOutside
''            m_frmSettings.chkDetailLCDScreen.Value = m_blnDetailLCDScreen
''            m_frmSettings.optDetailDrawMode.Item(m_intDetailDrawMode).Value = True
        
        Case "settings"
''            ' Just refresh the screen whilst user is choosing the settings.
''            Call GetMouseInput(m_intInputMouseSensitivity, Me.GameState, m_PlayerOne, m_Camera, m_frmCanvas)
''            Call GetKeyboardInput(Me.GameState, m_PlayerOne)
''            Call RefreshF5

            
        Case "run_game"
            ' ===============================================================
            ' Run game/simulation (which works slightly differently to demo).
            ' ===============================================================
            m_Camera.FreeCamera = True
            Call GetMouseInput(64, Me.GameState, m_PlayerOne, m_Camera, m_frmCanvas)
            Call GetKeyboardInput(Me.GameState, m_PlayerOne)
            Call ProcessGameElements
            Call ProcessParticles
            Call RefreshF5
            
            
        Case "quit"
            ' =======================================
            ' The user wants to quit the application.
            ' =======================================
            Me.GameState = "quitting"
            m_frmCanvas.TimerDoAnimation.Enabled = False
            Unload m_frmCanvas
            Set m_frmCanvas = Nothing
            
    End Select
    
End Sub



Private Sub m_frmCanvas_OnMouseDown()

    Call CreateParticle("projectile", m_Particles, m_PlayerOne, m_PlayerOne.WorldPosition)

End Sub

Private Sub m_frmCanvas_OnMouseUp()

    ' =====================================================================
    ' Ignore this routine. It is used to help create my animation sequence.
    ' =====================================================================
    
    Dim vectA As mdrVector4
    
    With m_PlayerOne
        
        vectA = .VPN
        vectA = VectorMultiplyByScalar(vectA, -1000)
        
        Debug.Print Round(vectA.X, 2) & "," & Round(vectA.Y, 2) & "," & Round(vectA.Z, 2) & vbTab;
        Debug.Print Round(.WorldPosition.X, 2) & "," & Round(.WorldPosition.Y, 2) & "," & Round(.WorldPosition.Z, 2)
        Debug.Print
    
    End With
    
End Sub



Public Sub ProcessParticles()

    Dim intIndex As Integer
    
    For intIndex = 0 To UBound(m_Particles)
    
        With m_Particles(intIndex)
            If .Enabled = True Then
            
            ' Process the different types of particles.
            Select Case .Class
                Case "projectile"
                    
                    ' ================================================
                    ' Rotate object (by setting the Identity), so that
                    ' it points at 'm_AllObjects(1).WorldPosition'
                    ' ================================================
                    .Parts(0).IdentityMatrix = MatrixInverse(MatrixLookAt(m_AllObjects(1).WorldPosition, .WorldPosition))
                    
                    If .LifeTime > 0.9 Then
                        ' Engine is OFF. That is, we've just launched the rocket!
                        .Vector.Y = .Vector.Y - 0.7
                    
                    ElseIf .LifeTime < 0.5 Then
                        .Vector.Y = .Vector.Y - 0.7
                        .Yaw = .Yaw + 4
                        .Pitch = .Pitch - 2
                        .LifeTime = .LifeTime + (1 / 200) ' Counter-act normal degredation
                    Else
                        ' Engine is ON.
                        
                        ' ===============================================================
                        ' Find new displacement vector (for the thrusting of the engine).
                        ' ===============================================================
                        Dim tempV As mdrVector4
                        tempV = VectorSubtract(m_AllObjects(1).WorldPosition, .WorldPosition)
                        tempV = VectorNormalize(tempV)
                        tempV = VectorMultiplyByScalar(tempV, 4) ' Adjust speed of rocket acceleration!!!
                                            
                        ' ===================================================================
                        ' Add the vectors together to get a "heat-seeking" missle effect.
                        ' Note: Don't let this get too fast. Overshooting a target is common.
                        ' ===================================================================
                        .Vector = VectorAddition(.Vector, tempV)
                        If VectorLength(.Vector) > 32 Then
                            .Vector = VectorMultiplyByScalar(.Vector, 0.7) ' Slow down everything by 70%
                        End If
                        
                        ' ===============
                        ' Generate Smoke.
                        ' ===============
                        Static lngBeanCounter As Long
                        lngBeanCounter = lngBeanCounter + 1
                        If (lngBeanCounter Mod 1) = 0 Then
                            Call CreateParticle("smoke", m_Particles, m_PlayerOne, .WorldPosition)
                        End If
                        
                    End If
                    
                    
                    ' ===========================================================
                    ' Update the world position based on the displacement vector.
                    ' ===========================================================
                    .WorldPosition = VectorAddition(.WorldPosition, .Vector)
                    
                    
                    ' =========================
                    ' Detect Crash with Ground.
                    ' =========================
                    If .WorldPosition.Y < (.UniformScale * 3) Then ' Has particle hit the ground?
                        .WorldPosition.Y = (.UniformScale * 3)
                        .Vector = VectorMultiplyByScalar(.Vector, 0.8) ' Slow down particle, now that it's hit the ground.
                        .LifeTime = .LifeTime - (1 / 10) ' Decrease it's life span very quickly when on the ground.
                        If (lngBeanCounter Mod 1) = 0 Then
                            ' Smoke it up, if boggie hits the deck!!
                            Call CreateParticle("smoke", m_Particles, m_PlayerOne, .WorldPosition)
                        End If
                        
                    End If
                    .LifeTime = .LifeTime - (1 / 200)
                    
                    
                    
                Case "smoke"
                    .UniformScale = .UniformScale + (Rnd / 3)
                    .Yaw = .Yaw - 0.5
                    If .Yaw < 0 Then .Yaw = 0
                    .Pitch = .Pitch + 3
                    .Roll = .Roll + 3
                    If .WorldPosition.Y < 0 Then .LifeTime = 0
                    .LifeTime = .LifeTime - (1 / 10)
                                        
            End Select
            
            
            ' Disable particle if it reaches it's lifespan.
            If .LifeTime <= 0 Then
                .LifeTime = 0
                .Enabled = False
            End If
            
            
            End If
        End With
        
    Next intIndex
    
End Sub
Public Sub ProcessGameElements()

    Static sngDir As Single
    Static sngDeg As Single
    Dim sngRad As Single
    
    If sngDir = 0 Then sngDir = 1
    
    sngDeg = sngDeg + sngDir
    
    If sngDeg + sngDir > 360 Then sngDir = -sngDir
    If sngDeg + sngDir < 0 Then sngDir = -sngDir
    
    
    sngRad = ConvertDeg2Rad(sngDeg)
    
    Dim sngX As Single
    Dim sngY As Single
    Dim sngZ As Single
    
    If sngDeg > 360 Then
        sngX = Cos(sngRad) * 700
        sngY = Sin(sngRad) * 700
        sngZ = Sin(sngRad) * 1500 + 1600
    Else
        sngX = Sin(sngRad) * 700
        sngY = Cos(sngRad) * 700
        sngZ = Cos(sngRad) * 1500 + 1600
    End If
    
    Dim intIndex As Integer
    
    For intIndex = 0 To UBound(m_AllObjects)
    
        With m_AllObjects(intIndex)
            If .Enabled = True Then
            
                ' Process the different types of game objects.
                Select Case .Class
                    Case "Geometry" ' Rotate the UFO gently.
                        .WorldPosition.Y = Abs(sngZ)
                        .WorldPosition.X = sngX
                        .WorldPosition.Z = sngY
                        .Yaw = .Yaw + 3
                        .Roll = .Roll + 0.1
                                        
                End Select
            
            End If
        End With
    Next intIndex
    
End Sub

