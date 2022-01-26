'----WorldSpace class

'members
Private shapeList_ As New Collection



'initializer
Private Sub Class_Initialize()
    
End Sub



'methods
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
