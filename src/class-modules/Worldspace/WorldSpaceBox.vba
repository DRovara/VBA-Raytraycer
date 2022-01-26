'----WorldSpaceBox class

Implements WorldSpaceShape

Private position_ As New Vector3
Private size_ As New Vector3

Public Function Init(position_value As Vector3, size_value As Vector3)
    position_ = position_value
    size_ = size_value
End Function

Private Sub Class_Initialize()
    Set position_ = New Vector3
    Call position_.Init(-5, -5, -10)
    Set size_ = New Vector3
    Call size_.Init(10, 10, 10)
End Sub

Property Get position() As Vector3
    Set position = position_
End Property

Property Get size() As Vector3
    Set size = size_
End Property

Public Function WorldSpaceShape_Distance(p As Vector3) As Double
    Dim difference As Vector3
    Set difference = DistanceVector(p, position)
    Dim dx As Double
    Dim dy As Double
    Dim dz As Double
    dx = Abs(difference.x) - size.x
    dy = Abs(difference.y) - size.y
    dz = Abs(difference.z) - size.z
    
    WorldSpaceShape_Distance = IIf(dx > dy, IIf(dx > dz, dx, dz), IIf(dy > dz, dy, dz))
End Function
