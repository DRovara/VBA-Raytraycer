Implements WorldSpaceShape

Private position_ As New Vector3
Private size_ As New Vector3
Private colour_ As Long
Private specularReflection_ As Double
Private diffuseReflection_ As Double
Private ambientReflection_ As Double
Private shininess_ As Double

Public Function Init(position_value As Vector3, size_value As Vector3, colour_value As Long, spec_value As Double, diff_value As Double, ambient_value As Double, shininess_value As Double)
    Set position_ = position_value
    Set size_ = size_value
    colour_ = colour_value
    specularReflection_ = spec_value
    diffuseReflection_ = diff_value
    ambientReflection_ = ambient_value
    shininess_ = shininess_value
End Function

Private Sub Class_Initialize()
    Set position_ = New Vector3
    Call position_.Init(0, 0, 0)
    Set size_ = New Vector3
    Call size_.Init(10, 10, 10)
    colour_ = RGB(0, 0, 255)
    specularReflection_ = 1
    diffuseReflection_ = 1
    ambientReflection_ = 1
    shininess_ = 1
End Sub

Property Get Position() As Vector3
    Set Position = position_
End Property

Property Get size() As Vector3
    Set size = size_
End Property

Property Get WorldSpaceShape_colour() As Long
    WorldSpaceShape_colour = colour_
End Property

Property Let WorldSpaceShape_colour(ByVal value As Long)
    colour_ = value
End Property

Property Get WorldSpaceShape_specularReflection() As Double
    WorldSpaceShape_specularReflection = specularReflection_
End Property

Property Let WorldSpaceShape_specularReflection(ByVal value As Double)
    specularReflection_ = value
End Property

Property Get WorldSpaceShape_diffuseReflection() As Double
    WorldSpaceShape_diffuseReflection = diffuseReflection_
End Property

Property Let WorldSpaceShape_diffuseReflection(ByVal value As Double)
    diffuseReflection_ = value
End Property

Property Get WorldSpaceShape_ambientReflection() As Double
    WorldSpaceShape_ambientReflection = ambientReflection_
End Property

Property Let WorldSpaceShape_ambientReflection(ByVal value As Double)
    ambientReflection_ = value
End Property

Property Get WorldSpaceShape_shininess() As Double
    WorldSpaceShape_shininess = shininess_
End Property

Property Let WorldSpaceShape_shininess(ByVal value As Double)
    shininess_ = value
End Property

Public Function WorldSpaceShape_Distance(p As Vector3) As Double
    Dim difference As Vector3
    Set difference = DistanceVector(p, Position)
    Dim dx As Double
    Dim dy As Double
    Dim dz As Double
    dx = Abs(difference.x) - size.x / 2
    dy = Abs(difference.y) - size.y / 2
    dz = Abs(difference.z) - size.z / 2
    
    WorldSpaceShape_Distance = IIf(dx > dy, IIf(dx > dz, dx, dz), IIf(dy > dz, dy, dz))
End Function
