'members
Private position_ As Vector3
Private direction_ As Vector3
Private up_ As Vector3



'initializer
Private Sub Class_Initialize()
    Set position_ = New Vector3
    Call position_.Init(0, 0, 20)
    Set direction_ = New Vector3
    Call direction_.Init(0, 0, -1)
    
    Set up_ = New Vector3
    Call up_.Init(0, 1, 0)
End Sub

Public Function Init(position_value As Vector3, direction_value As Vector3, up_value As Vector3)
    Set position_ = position_value
    Set direction_ = direction_value.Normalize
    Set up_ = up_value.Normalize
    
    Dim dot As Double
    dot = DotProduct(direction_, up_)
    If dot <> 0 Then
        Set up_ = CrossProduct(direction_, up_).Normalize
    End If
End Function



'properties
Property Get position() As Vector3
    Set position = position_
End Property

Property Get direction() As Vector3
    Set direction = direction_
End Property

Property Get up() As Vector3
    Set up = up_
End Property



'methods
