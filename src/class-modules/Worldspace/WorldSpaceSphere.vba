Implements WorldSpaceShape

Private position_ As Vector3
Private radius_ As Double
Private colour_ As Long
Private specularReflection_ As Double
Private diffuseReflection_ As Double
Private ambientReflection_ As Double
Private shininess_ As Double

Public Function Init(position_value As Vector3, radius_value As Double, colour_value As Long, spec_value As Double, diff_value As Double, ambient_value As Double, shininess_value As Double)
    Set position_ = position_value
    radius_ = radius_value
    colour_ = colour_value
    specularReflection_ = spec_value
    diffuseReflection_ = diff_value
    ambientReflection_ = ambient_value
    shininess_ = shininess_value
End Function

Private Sub Class_Initialize()
    Set position_ = New Vector3
    Call position_.Init(0, 0, 0)
    radius_ = 5
    specularReflection_ = 1
    diffuseReflection_ = 1
    ambientReflection_ = 1
    shininess_ = 1
    colour_ = RGB(0, 0, 255)
End Sub

Property Get Position() As Vector3
    Set Position = position_
End Property

Property Get radius() As Double
    radius = radius_
End Property

Property Get WorldSpaceShape_colour() As Long
    WorldSpaceShape_colour = colour_
End Property

Property Let WorldSpaceShape_colour(ByVal value As Long)
    colour_ = value
End Property

Public Function WorldSpaceShape_Distance(p As Vector3) As Double
    WorldSpaceShape_Distance = DistanceVector(p, Position).Magnitude - radius
End Function

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
