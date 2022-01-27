Implements WorldSpaceShape

Private position_ As New Vector3
Private size_ As New Vector3
Private colour_ As Long

Public Function Init(position_value As Vector3, size_value As Vector3, colour_value As Long)
    Set position_ = position_value
    Set size_ = size_value
    colour_ = colour_value
End Function

Private Sub Class_Initialize()
    Set position_ = New Vector3
    Call position_.Init(0, 0, 0)
    Set size_ = New Vector3
    Call size_.Init(10, 10, 10)
    colour_ = RGB(0, 0, 255)
End Sub

Property Get position() As Vector3
    Set position = position_
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

Public Function WorldSpaceShape_Distance(p As Vector3) As Double
    Dim difference As Vector3
    Set difference = DistanceVector(p, position)
    Dim dx As Double
    Dim dy As Double
    Dim dz As Double
    dx = Abs(difference.x) - size.x / 2
    dy = Abs(difference.y) - size.y / 2
    dz = Abs(difference.z) - size.z / 2
    
    WorldSpaceShape_Distance = IIf(dx > dy, IIf(dx > dz, dx, dz), IIf(dy > dz, dy, dz))
End Function
