Public Function NewVector3(x As Double, y As Double, z As Double) As Vector3
    Dim v As New Vector3
    Set NewVector3 = v.Init(x, y, z)
End Function

Public Function DistanceVector(fromPoint As Vector3, toPoint As Vector3) As Vector3
    Dim v As New Vector3
    Set DistanceVector = v.Init(toPoint.x - fromPoint.x, toPoint.y - fromPoint.y, toPoint.z - fromPoint.z)
End Function

Public Function CrossProduct(u As Vector3, v As Vector3) As Vector3
    Dim w As New Vector3
    Set CrossProduct = w.Init(u.y * v.z - u.z * v.y, u.z * v.x - u.x * v.z, u.x * v.y - u.y * v.x)
End Function

Public Function CreateRaycaster(cam As ViewerCamera, world As WorldSpace, planeDistance As Double, far As Double, stepSize As Double, planeWidth As Double, planeHeight As Double, pixelWidth As Integer, pixelHeight As Integer) As Raycaster
    Dim r As New Raycaster
    r.cam = cam
    r.world = word
    r.planeDistance = planeDistance
    r.far = far
    r.stepSize = stepSize
    r.planeWidth = planeWidth
    r.planeHeight = planeHeight
    r.pixelWidth = pixelWidth
    r.pixelHeight = pixelHeight
    CreateRaycaster = r
End Function

Public Function DotProduct(u As Vector3, v As Vector3) As Double
    DotProduct = u.x * v.x + u.y * v.y + u.z * v.z
End Function

Public Function NormalAt(world As WorldSpace, pt As Vector3) As Vector3
    Dim result As New Vector3
    Dim x As Double
    Dim y As Double
    Dim z As Double
    
    Dim epsilon As Double
    epsilon = 0.000001
    
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
    Set NormalAt = result.Normalize
    
End Function

Public Function ColourToVector(col As Long) As Vector3
    Dim result As New Vector3
    Dim r As Long
    Dim b As Long
    Dim g As Long
    
    b = col / 65536
    g = (col - b * 65536) / 256
    r = col - b * 65536 - g * 256
    Call result.Init(r / 255#, g / 255#, b / 255#)
    Set ColourToVector = result
End Function
