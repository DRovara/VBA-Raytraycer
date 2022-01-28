'members
Private shapeList_ As New Collection
Private lightList_ As New Collection



'initializer
Private Sub Class_Initialize()
    
End Sub



'methods
Public Sub AddLight(L As Light)
    lightList_.Add L
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

Public Function ColourAt(p As Vector3, cam As ViewerCamera) As Long
    Dim minDistance As Double
    minDistance = 999999999
    Dim bestShape As WorldSpaceShape
    
    Dim i As Long
    For i = 1 To shapeList_.Count
        Dim dist As Double
        dist = shapeList_(i).Distance(p)
        If Abs(dist) < minDistance Then
            minDistance = Abs(dist)
            Set bestShape = shapeList_(i)
        End If
    Next i
    
    Dim N As Vector3
    Dim V As Vector3
    Dim R As Vector3
    Dim L As Vector3
    
    Set N = NormalAt(Me, p).Normalize
    Set V = DistanceVector(p, cam.Position).Normalize
    
    Dim colourSum As Vector3
    Set colourSum = ColourToVector(bestShape.colour).Multiply(bestShape.ambientReflection)
    
    For i = 1 To lightList_.Count
        Dim lgt As Light
        Set lgt = lightList_(i)
        Set L = DistanceVector(p, lgt.Position).Normalize
        Set R = DistanceVector(L, N.Multiply(2).Multiply(DotProduct(L, N))).Normalize
        
        Dim specular_part As Double
        Dim diffuse_part As Double
        specular_part = DotProduct(R, V)
        If specular_part < 0 Then
            specular_part = 0
        End If
        specular_part = (specular_part ^ bestShape.shininess) * bestShape.specularReflection
        diffuse_part = bestShape.diffuseReflection * DotProduct(L, N)
        If diffuse_part < 0 Then
            diffuse_part = 0
        End If
        
        Dim lightComponent As Vector3
        Set lightComponent = lgt.SpecularColour.Multiply(specular_part).Add(lgt.DiffuseColour.Multiply(diffuse_part))
        colourSum.Translate lightComponent
        
    Next i
    ColourAt = VectorToColour(colourSum)
End Function
