Attribute VB_Name = "mParticles"
Option Explicit

Private m_objIcosahedron As mdr3DObject
Private m_objSmoke As mdr3DObject

Private Sub Init()

    'm_objIcosahedron = mDirectXParser2.LoadXFile(App.Path & "\xFiles\icosahedron2.X")
    'm_objIcosahedron = mDirectXParser2.LoadXFile(App.Path & "\xFiles\tetrahedron.X")
    'm_objIcosahedron = mDirectXParser2.LoadXFile(App.Path & "\xFiles\rock5.X")
    m_objIcosahedron = mDirectXParser2.LoadXFile(App.Path & "\xFiles\arrow3.X")
    m_objIcosahedron.Class = "icosahedron"
    
    m_objSmoke = mDirectXParser2.LoadXFile(App.Path & "\xFiles\bomb3.X")
    m_objSmoke.Class = "smoke"

End Sub

Public Function CreateParticle(p_strParticleClass As String, p_Particles() As mdr3DObject, p_PlayerOne As mdrPlayer, p_WorldPosition As mdrVector4) As Integer
    
    ' Returns the Index number of the particle created (if any).
    CreateParticle = 0
    
    If m_objIcosahedron.Class <> "icosahedron" Then Call Init
    
    Dim intIndex As Integer
    
    For intIndex = LBound(p_Particles) To UBound(p_Particles)
        With p_Particles(intIndex)
            If .Enabled = False Then
            
                ' Unused/disabled particle found, so use this.
                
                Select Case p_strParticleClass
                    Case "projectile"
                        p_Particles(intIndex) = m_objIcosahedron
                        ' Note: The LoadXFile routine resets all parameters for the object!
                        .Class = p_strParticleClass
                        .Enabled = True
                        .WorldPosition = p_PlayerOne.WorldPosition
                        .UniformScale = 6
                        
                        .Vector = p_PlayerOne.VPN
                        .Vector = VectorMultiplyByScalar(.Vector, -16)
                        .LifeTime = 1
                
                    Case "smoke"
                        p_Particles(intIndex) = m_objSmoke
                        .Class = p_strParticleClass
                        .Enabled = True
                        .WorldPosition = p_WorldPosition
                        .UniformScale = 0.5
                        .LifeTime = 1
                        .Yaw = 20
                        
                End Select
                
                CreateParticle = intIndex
                Exit For
            End If
        End With
    Next intIndex

End Function

