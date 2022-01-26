'----WorldSpaceSphere class

Implements WorldSpaceShape

Private position_ As Vector3
Private radius_ As Double

Public Function Init(position_value As Vector3, radius_value As Double)
    position_ = position_value
    radius_ = radius_value
End Function

Private Sub Class_Initialize()
    Set position_ = New Vector3
    Call position_.Init(0, 0, 0)
    radius_ = 5
End Sub

Property Get position() As Vector3
    Set position = position_
End Property

Property Get radius() As Double
    radius = radius_
End Property

Public Function WorldSpaceShape_Distance(p As Vector3) As Double
    WorldSpaceShape_Distance = DistanceVector(p, position).Magnitude - radius
End Function

