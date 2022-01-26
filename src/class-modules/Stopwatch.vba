'----Stopwatch class

Private lastTime As Double
Private measurements As New Collection
Private measurementIndex As Integer

Public Property Get size() As Integer
    size = measurements.Count
End Property

Public Sub Start()
    lastTime = Timer
    measurementIndex = 1
End Sub

Public Sub Reset()
    lastTime = Timer
    measurementIndex = 1
    While measurements.Count > 0
        measurements.Remove 1
    End Sub
End Sub

Public Function Time()
    Dim t As Double
    t = Timer - lastTime
    lastTime = Timer
    If measurementIndex > measurements.Count Then
        Dim newCol As New Collection
        measurements.Add newCol
    End If
    measurements.Item(measurementIndex).Add t
    measurementIndex = measurementIndex + 1
    Time = t
End Function

Public Sub Lap()
    measurementIndex = 1
End Sub

Public Function Sums() As Double()
    Dim sms() As Double
    ReDim sms(measurements.Count)
    Dim i As Integer
    Dim j As Integer
    
    For i = 1 To measurements.Count
        Dim s As Double
        s = 0
        For j = 1 To measurements.Item(i).Count
            s = s + measurements.Item(i).Item(j)
        Next j
        sms(i - 1) = s
    Next i
    Sums = sms
End Function
