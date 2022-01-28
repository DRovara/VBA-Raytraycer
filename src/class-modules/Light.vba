'members
Private position_ As Vector3
Private specularColour_ As Vector3
Private diffuseColour_ As Vector3
Private ambientColour_ As Vector3

'properties
Public Property Get Position() As Vector3
    Set Position = position_
End Property

Public Property Get SpecularColour() As Vector3
    Set SpecularColour = specularColour_
End Property

Public Property Get DiffuseColour() As Vector3
    Set DiffuseColour = diffuseColour_
End Property

Public Property Get AmbientColour() As Vector3
    Set AmbientColour = ambientColour_
End Property

'initializer
Private Sub Class_Initialize()
    Set position_ = New Vector3
    Set specularColour_ = New Vector3
    Set diffuseColour_ = New Vector3
    Set ambientColour_ = New Vector3
End Sub

Public Sub Init(position_value As Vector3, specular As Long, diffuse As Long, ambient As Long)
    Set position_ = position_value
    Set diffuseColour_ = ColourToVector(diffuse)
    Set specularColour_ = ColourToVector(specular)
    Set ambientColour_ = ColourToVector(ambient)
End Sub
