Attribute VB_Name = "mDataStructures"
Option Explicit

' See Help File; Section N.N.N - Constants » PrimitiveTypes
Public Enum mdrPrimitiveType
    mdrPoints = 0
    mdrLines = 1
    mdrLineLoop = 2
    mdrLineStrip = 3
    mdrTriangles = 4
    mdrTriangleStrip = 5
    mdrTriangleFan = 6
    mdrQuads = 7
    mdrQuadStrip = 8
    mdrPolygon = 9
End Enum

Public Enum eCrossHairStyle
    mdrCRS_None = 0
    mdrCRS_DotOnly = 1
    mdrCRS_Cross = 2
    mdrCRS_X = 4
    mdrCRS_Large = 8
    mdrCRS_Scope1 = 16
    mdrCRS_Scope2 = 32
End Enum


' Section N.N.N - Structures » Vectors » 4D
Public Type mdrVector4
    X As Single
    Y As Single
    Z As Single
    w As Single
End Type


' Section N.N.N - Structures » Matrix » 4x4
Public Type mdrMatrix4
    rc11 As Single: rc12 As Single: rc13 As Single: rc14 As Single
    rc21 As Single: rc22 As Single: rc23 As Single: rc24 As Single
    rc31 As Single: rc32 As Single: rc33 As Single: rc34 As Single
    rc41 As Single: rc42 As Single: rc43 As Single: rc44 As Single
End Type


' Section N.N.N - Structures » Vertex
Public Type mdrVertex
    Pxyz As mdrVector4
    Txyz As mdrVector4
    Wxyz As mdrVector4
    RGB_Red As Single
    RGB_Green As Single
    RGB_Blue As Single
    Brightness As Single
    Clipped As Boolean
End Type


' Section N.N.N - Structures » Primitive
Public Type mdrPrimitive
    ID              As String           ' General purpose: Reference number or string.
    Class           As String           ' General purpose: Class of Object. ie. Geometry, Helper Object, etc.
    Enabled         As Boolean          ' General purpose: Enable/Disable processing?
    Visible         As Boolean          ' General purpose: Visible or Hidden from GUI.
    Selected        As Boolean          ' General purpose: Determins if the object is selected or not.
    Caption         As String           ' General purpose: point, face, edge, wall, door, etc. (Optional)
    Description     As String           ' General purpose: A Caption should always have a Description. (Optional)
    
    PrimitiveType   As mdrPrimitiveType ' Primitive Type. ie. Point, Line, Triangles, Polygon, Quads, etc.
    Vertices()      As mdrVertex        ' A Vertex List.
End Type


Public Type mdr3DPart
    Caption As String                   ' Helicopter Blades, Landing Gear, Gun Turret, Leg, Head, Arm, etc. (Optional)
    Description As String               ' A Caption should always have a Description. (Optional)
    Enabled As Boolean                  ' General purpose: Enable/Disable processing?
    Selected As Boolean                 ' General purpose: Determins if the object is selected or not.
    
    RGB_Red As Single                   ' 0 to 1
    RGB_Green As Single                 ' 0 to 1
    RGB_Blue As Single                  ' 0 to 1
    
    Vertices() As mdrVertex             ' The original vertices that make up the object (these never changed once defined)
    VerticesT() As mdrVertex            ' The transformed vertices; a temporary working area.
    Faces() As Variant                  ' Connect the dots [Vertices] together to form shapes.
    
    PointingAt As mdrVector4            ' Defines the direction the object is pointing. (Alternative to Pitch, Roll & Yaw.)
    
    IdentityMatrix As mdrMatrix4        ' This holds the initial or default starting position for the polyhedron (rotation, size & position). (Optional)
    
    Primitives()   As mdrPrimitive      ' Primitives that make up the object.
End Type


' ======================================================================
' A 3D object is usually a collection of smaller objects (ie. Parts)
' ======================================================================
Public Type mdr3DObject
    ID              As String       ' General purpose: Reference number or string.
    Class           As String       ' General purpose: Class of Object. ie. Geometry, Helper Object, etc.
    Enabled         As Boolean      ' General purpose: Enable/Disable processing?
    Visible         As Boolean      ' General purpose: Visible or Hidden from GUI.
    Selected        As Boolean      ' General purpose: Determins if the object is selected or not.
    Caption         As String       ' General purpose: Helicopter, Tank, Space Ship, Monster, etc. (Optional)
    Description     As String       ' General purpose: A Caption should always have a Description. (Optional)
    
    WorldPosition   As mdrVector4   ' Position of the Object in World Coordinates.
    PointingAt      As mdrVector4   ' Defines the direction the object is pointing. (Alternative to Pitch, Roll & Yaw.)
    Pitch As Single                ' Angle in degrees
    Roll As Single                 ' Angle in degrees
    Yaw As Single                  ' Angle in degrees
    UniformScale    As Single       ' Uniform scale on all axes. Typically equals 1.0
    Vector          As mdrVector4   ' Direction and Magnitude of the 3D Object. ie. Which way is the object moving, and how fast?
    Parts()         As mdr3DPart    ' This object is made up from Parts.
    
    CastShadows     As Boolean      ' Render Option:

    LifeTime        As Single       ' Particle behaviour.
    
End Type


Public Type mdrPlayer
    WorldPosition   As mdrVector4   ' Position of the Player.
    VPN             As mdrVector4   ' Which direction is the player point.
    XZPlane         As Single       ' (for internal use: See GetMouseInput)
    XYPlane         As Single       ' (for internal use: See GetMouseInput)
    LeftRightVector As mdrVector4
End Type


''''' ===========================================================================================
''''' The ViewFrustum defines the 3D view-volume and 2D window through which we see the 3D world.
''''' ===========================================================================================
''''Public Type mdrViewPort
''''    Caption As String                   '   Left-Eye, Right-Eye, Main Window, Picture-in-Picture, Rear View Mirror, etc. (Optional)
''''    Description As String               '   A Caption should always have a Description. (Optional)
''''    Umin As Double                      '   The UV coordinate system coincides with the screen's XY coordinates.
''''    Umax As Double                      '       "
''''    Vmin As Double                      '       "
''''    Vmax As Double                      '       "
''''    ClipFar As Double                   '   Don't draw Vertices further away than this value. Any value higher than 0.
''''    ClipNear As Double                  '   Don't draw Vertices that are this close to us (or behind us). Typically 0, but can be higher.
''''
''''    FOV As Single                       '   Field Of View (FOV). "90 degree FOV" = "1x Zoom". If you update FOV, don't forget to update Zoom.
''''    Zoom As Single                      '   (The Zoom value is calculated from the FOV. Normally you define one of them, then calculate the other.)
''''
''''    OffsetU As Double                   '   2D Screen offset coordinate. (Optional)
''''    OffsetV As Double                   '   2D Screen offset coordinate. (Optional)
''''End Type


' ============================================
' This is our Virtual 3D Target Camera object.
' ============================================
Public Type mdr3DTargetCamera
    ID              As String       ' General purpose: Reference number or string.
    Class           As String       ' Class of Object
    Title           As String       '
    
    Visible         As Boolean      ' General purpose: Visible or Hidden from GUI.
    Caption         As String       ' Camera1, Director's Chair, Birds-eye View, etc. (Optional)
    Description     As String       ' A Caption should always have a Description.     (Optional)
    
    WorldPosition   As mdrVector4   ' VRP - Position of the Camera in World Coordinates.
    LookAtPoint     As mdrVector4   ' This is where the Camera is looking at in World Coordinates.
    VPN             As mdrVector4   ' This is an alternative to LookAtPoint, but the programmer is responsible for additional calculations.
    FreeCamera      As Boolean      ' If this value is TRUE, then use the LookTowards value.
    VUP             As mdrVector4   ' Which way is UP?
    PRP             As mdrVector4   ' Projection Reference Point (PRP). Used for perspective distortion & stereopsis.
    
    Umin            As Single       ' The UV coordinate system coincides with the screen's XY coordinates. See VPXmin etc. below.
    Umax            As Single       ' Typically, the UV coordinates will be between -1 and +1.
    Vmin            As Single       '
    Vmax            As Single       '
    UPan            As Single       ' The UV/Pan values are only useful for multiple-monitor situations.
    VPan            As Single       ' Leave these pan values at zero.
     
    ClipFar         As Single       ' Specified relative to VRP. Positive distance in the direction of VPN. This value is usually positive.
    ClipNear        As Single       ' Specified relative to VRP. Positive distance in the direction of VPN. This value is usually negative.
    
    VPXmin          As Single       ' ViewPort Xmin. This value should be in Pixels.
    VPXmax          As Single       ' ViewPort Xmax. This value should be in Pixels.
    VPYmin          As Single       ' ViewPort Ymin. This value should be in Pixels.
    VPYmax          As Single       ' ViewPort Ymax. This value should be in Pixels.
    
    RollAngle       As Single       '
    FOV             As Single       ' Field Of View (FOV). "90 degree FOV" = "1x Zoom". If you update FOV, don't forget to update Zoom.
    Zoom            As Single       ' (The Zoom value is calculated from the FOV. Normally you define one of them, then calculate the other.)
    
    ViewMatrix      As mdrMatrix4   ' View Matrix.
End Type

