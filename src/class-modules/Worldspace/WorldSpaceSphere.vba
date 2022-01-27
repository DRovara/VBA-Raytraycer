Implements WorldSpaceShape

Private position_ As Vector3
Private radius_ As Double
Private colour_ As Long

Public Function Init(position_value As Vector3, radius_value As Double, colour_value As Long)
    Set position_ = position_value
    radius_ = radius_value
    colour_ = colour_value
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

Property Get WorldSpaceShape_colour() As Long
    WorldSpaceShape_colour = colour_
End Property

Property Let WorldSpaceShape_colour(ByVal value As Long)
    colour_ = value
End Property

Public Function WorldSpaceShape_Distance(p As Vector3) As Double
    WorldSpaceShape_Distance = DistanceVector(p, position).Magnitude - radius
End Function

