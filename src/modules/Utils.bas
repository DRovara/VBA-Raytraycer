Public Function NewVector3(x As Double, y As Double, z As Double) As Vector3
    Dim V As New Vector3
    Set NewVector3 = V.Init(x, y, z)
End Function

Public Function DistanceVector(fromPoint As Vector3, toPoint As Vector3) As Vector3
    Dim V As New Vector3
    Set DistanceVector = V.Init(toPoint.x - fromPoint.x, toPoint.y - fromPoint.y, toPoint.z - fromPoint.z)
End Function

Public Function CrossProduct(u As Vector3, V As Vector3) As Vector3
    Dim w As New Vector3
    Set CrossProduct = w.Init(u.y * V.z - u.z * V.y, u.z * V.x - u.x * V.z, u.x * V.y - u.y * V.x)
End Function

Public Function CreateRaycaster(cam As ViewerCamera, world As WorldSpace, planeDistance As Double, far As Double, stepSize As Double, planeWidth As Double, planeHeight As Double, pixelWidth As Integer, pixelHeight As Integer) As Raycaster
    Dim R As New Raycaster
    R.cam = cam
    R.world = word
    R.planeDistance = planeDistance
    R.far = far
    R.stepSize = stepSize
    R.planeWidth = planeWidth
    R.planeHeight = planeHeight
    R.pixelWidth = pixelWidth
    R.pixelHeight = pixelHeight
    CreateRaycaster = R
End Function

Public Function DotProduct(u As Vector3, V As Vector3) As Double
    DotProduct = u.x * V.x + u.y * V.y + u.z * V.z
End Function

Public Function NormalAt(world As WorldSpace, pt As Vector3) As Vector3
    Dim result As New Vector3
    Dim x As Double
    Dim y As Double
    Dim z As Double
    
    Dim epsilon As Double
    epsilon = 0.000000000001
    
    Dim vXP As New Vector3
    Dim vYP As New Vector3
    Dim vZP As New Vector3
    Dim vXN As New Vector3
    Dim vYN As New Vector3
    Dim vZN As New Vector3
    Call vXP.Init(pt.x + epsilon, pt.y, pt.z)
    Call vYP.Init(pt.x, pt.y + epsilon, pt.z)
    Call vZP.Init(pt.x, pt.y, pt.z + epsilon)
    Call vXN.Init(pt.x - epsilon, pt.y, pt.z)
    Call vYN.Init(pt.x, pt.y - epsilon, pt.z)
    Call vZN.Init(pt.x, pt.y, pt.z - epsilon)
    
    x = world.Distance(vXP) - world.Distance(vXN)
    y = world.Distance(vYP) - world.Distance(vYN)
    z = world.Distance(vZP) - world.Distance(vZN)
    
    
    Call result.Init(x, y, z)
    
    If result.SqrMagnitude < 1E-30 Then
        Call result.Init(0, 0, 1)
    End If
    Set NormalAt = result.Normalize
    
End Function

Public Function Floor(val As Double) As Long
    Dim result As Long
    result = Int(val)
    Floor = result
End Function

Public Function ColourToVector(col As Long) As Vector3
    Dim result As New Vector3
    Dim R As Long
    Dim b As Long
    Dim g As Long
    
    b = Floor(col / 65536#)
    g = Floor((col - (b * 65536)) / 256#)
    R = col - b * 65536 - g * 256
    Call result.Init(R / 255#, g / 255#, b / 255#)
    Set ColourToVector = result
End Function

Public Function VectorToColour(vec As Vector3) As Long
    Dim result As Long
    Dim clamped As New Vector3
    Call clamped.Init(IIf(vec.x > 1, 1, vec.x), IIf(vec.y > 1, 1, vec.y), IIf(vec.z > 1, 1, vec.z))
    
    result = Int(clamped.z * 255) * 256
    result = result + Int(clamped.y * 255)
    result = result * 256
    result = result + Int(clamped.x * 255)
    
    VectorToColour = result
End Function
