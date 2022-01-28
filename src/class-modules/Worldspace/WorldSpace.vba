'members
Private shapeList_ As New Collection
Private lightList_ As New Collection



'initializer
Private Sub Class_Initialize()
    
End Sub



'methods
Public Sub AddLight(l As Light)
    lightList_.Add l
End Sub


Public Sub AddShape(shape As WorldSpaceShape)
    shapeList_.Add shape
End Sub

Public Function Distance(p As Vector3) As Double
    Dim minDistance As Double
    minDistance = 999999999
    
    Dim i As Long
    For i = 1 To shapeList_.Count
        Dim dist As Double
        dist = shapeList_(i).Distance(p)
        If dist < minDistance Then
            minDistance = dist
        End If
    Next i
    Distance = minDistance
End Function

Public Function ColourAt(p As Vector3) As Long
    Dim minDistance As Double
    minDistance = 999999999
    Dim colour As Long
    
    Dim i As Long
    For i = 1 To shapeList_.Count
        Dim dist As Double
        dist = shapeList_(i).Distance(p)
        If Abs(dist) < minDistance Then
            minDistance = Abs(dist)
            colour = shapeList_(i).colour
        End If
    Next i
    ColourAt = colour
End Function
