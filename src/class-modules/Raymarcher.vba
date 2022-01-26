'----Raymarcher class

'members
Private cam_ As ViewerCamera
Private planeDistance_ As Double
Private far_ As Double
Private world_ As WorldSpace
Private stepSize_ As Double

Private pixelWidth_ As Integer
Private pixelHeight_ As Integer

Private planeWidth_ As Double
Private planeHeight_ As Double



'initializer
Private Sub Class_Initialize()
    Set cam_ = New ViewerCamera
    planeDistance_ = 10
    far_ = 50
    Set world_ = New WorldSpace
    stepSize_ = 1
    
    pixelWidth_ = 200
    pixelHeight_ = 150
    
    planeWidth_ = 20
    planeHeight = 15
End Sub



'properties
Property Get cam() As ViewerCamera
    Set cam = cam_
End Property

Property Get planeDistance() As Double
    planeDistance = planeDistance_
End Property

Property Get far() As Double
    far = far_
End Property

Property Get world() As WorldSpace
    Set world = world_
End Property

Property Get stepSize() As Double
    stepSize = stepSize_
End Property

Property Get pixelWidth() As Integer
    pixelWidth = pixelWidth_
End Property

Property Get pixelHeight() As Integer
    pixelHeight = pixelHeight_
End Property

Property Get planeWidth() As Double
    planeWidth = planeWidth_
End Property

Property Get planeHeight() As Double
    planeHeight = planeHeight_
End Property



Property Set cam(value As ViewerCamera)
    Set cam_ = value
End Property

Property Let planeDistance(value As Double)
    planeDistance_ = value
End Property

Property Let far(value As Double)
    far_ = value
End Property

Property Set world(value As WorldSpace)
    Set world_ = value
End Property

Property Let stepSize(value As Double)
    stepSize_ = value
End Property

Property Let pixelWidth(value As Integer)
    pixelWidth_ = value
End Property

Property Let pixelHeight(value As Integer)
    pixelHeight_ = value
End Property

Property Let planeWidth(value As Double)
    planeWidth_ = value
End Property

Property Let planeHeight(value As Double)
    planeHeight_ = value
End Property



'methods
Public Function run() As Boolean()
    Application.DisplayStatusBar = True
    Dim planeCenter As Vector3
    Set planeCenter = cam.position.Add(cam.direction.Multiply(planeDistance))
    
    Dim pW As Double
    Dim pH As Double
    pW = planeWidth / pixelWidth
    pH = planeHeight / pixelHeight
    
    Dim stepX As New Vector3
    Dim stepY As New Vector3
    Call stepY.Init(cam.up.x, cam.up.y, cam.up.z)
    Set stepX = CrossProduct(cam.direction, cam.up).Normalize
    
    Dim result() As Boolean
    ReDim result(pixelHeight, pixelWidth)
    
    Dim y As Integer
    Dim x As Integer
    For y = 0 To pixelHeight - 1
        Dim dy As Double
        dy = -planeHeight / 2 + y * pH
        For x = 0 To pixelWidth - 1
            If x Mod 50 = 0 Then
                Dim msg As String
                msg = "Pixel " & (x) & ", " & (y)
                Sheet1.Cells(1, 1).value = msg
            End If
            Dim dx As Double
            dx = -planeWidth / 2 + x * pW
            Dim planePoint As Vector3
            Set planePoint = planeCenter.Add(stepY.Multiply(dy)).Add(stepX.Multiply(dx))
            result(y, x) = Cast(cam.position, planePoint)
            If result(y, x) Then
                Sheet1.Cells(2, 1).value = "New: " & (y) & " " & (x)
            Else
                Sheet1.Cells(2, 1).value = "No:  " & (y) & " " & (x)
            End If
        Next x
    Next y
    
    run = result
End Function

Public Function Cast(startPosition As Vector3, toPosition As Vector3) As Boolean
    Dim direction As Vector3
    Set direction = DistanceVector(startPosition, toPosition).Normalize.Multiply(stepSize)
    Dim current As New Vector3
    Call current.Init(startPosition.x, startPosition.y, startPosition.z)
    
    Dim totalDistance As Double
    totalDistance = 0
    
    While totalDistance < far
        DoEvents
        If world.Intersects(current) Then
            Cast = True
            Exit Function
        End If
        current.Translate direction
        totalDistance = totalDistance + stepSize
    Wend
    Cast = world.Intersects(current)
End Function
