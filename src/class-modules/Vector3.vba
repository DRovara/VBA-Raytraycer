'members
Private x_ As Double
Private y_ As Double
Private z_ As Double



'constructor
Public Function Init(x_val As Double, y_val As Double, z_val As Double) As Vector3
    x_ = x_val
    y_ = y_val
    z_ = z_val
    Set Init = Me
End Function



'properties
Property Get x() As Double
    x = x_
End Property

Property Get y() As Double
    y = y_
End Property

Property Get z() As Double
    z = z_
End Property

Property Get SqrMagnitude() As Double
    SqrMagnitude = x * x + y * y + z * z
End Property

Property Get Magnitude() As Double
    Magnitude = Sqr(SqrMagnitude)
End Property



'methods
Public Function Add(V As Vector3) As Vector3
    Dim v2 As Vector3
    Set v2 = New Vector3
    Set Add = v2.Init(x + V.x, y + V.y, z + V.z)
End Function

Public Function Multiply(scalar As Double) As Vector3
    Dim v2 As Vector3
    Set v2 = New Vector3
    Set Multiply = v2.Init(x * scalar, y * scalar, z * scalar)
End Function

Public Function Normalize() As Vector3
    Dim v2 As New Vector3
    Dim length As Double
    length = Magnitude
    Set Normalize = v2.Init(x / length, y / length, z / length)
End Function

Public Sub Translate(V As Vector3)
    x_ = x_ + V.x
    y_ = y_ + V.y
    z_ = z_ + V.z
End Sub
