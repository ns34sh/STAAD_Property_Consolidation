Sub Main()
Dim objOpenSTAAD As Object
Dim lBeamCnt As Long
Dim BeamNumberArray() As Long
'Get the application object

Set objOpenSTAAD = GetObject(, "StaadPro.OpenSTAAD")
'Get Beam Numbers

lBeamCnt = objOpenSTAAD.Geometry.GetMemberCount
ReDim BeamNumberArray(0 To (lBeamCnt - 1)) As Long

Dim lBeamSectionName() As String
ReDim lBeamSectionName(0 To (lBeamCnt - 1)) As String
Dim lBSPropertyTypeNo() As Integer
ReDim lBSPropertyTypeNo(0 To (lBeamCnt - 1)) As Integer

'Get Beam list
objOpenSTAAD.Geometry.GetBeamList BeamNumberArray
For i = 0 To lBeamCnt - 1
    lBeamSectionName(i) = objOpenSTAAD.Property.GetBeamSectionName(BeamNumberArray(i))
    lBSPropertyTypeNo(i) = objOpenSTAAD.Property.GetBeamSectionPropertyRefNo(BeamNumberArray(i))
Next i

Dim d As Object
Set d = CreateObject("Scripting.Dictionary")
'Set d = New Scripting.Dictionary

For i = 0 To lBeamCnt - 1
    d(lBeamSectionName(i)) = 1
Next i

Dim TempBeamArray() As Integer
'ReDim TempBeamArray(0 To (lBeamCnt - 1)) As Integer

Dim MinPropRefNo As Integer

Dim v As Variant
'For Each v In d.Keys()
For j = LBound(d.keys()) To UBound(d.keys())
    Erase TempBeamArray

    MinPropRefNo = 0
    Counter = 0
    If CStr(d.keys()(j)) = "" Then GoTo ContinueFor:
    For i = 0 To lBeamCnt - 1
        If CStr(d.keys()(j)) = lBeamSectionName(i) Then
            Counter = Counter + 1
            ReDim Preserve TempBeamArray(0 To (Counter - 1)) As Integer
            TempBeamArray(Counter - 1) = BeamNumberArray(i)
            If Counter = 1 Then
                MinPropRefNo = lBSPropertyTypeNo(i)
            ElseIf lBSPropertyTypeNo(i) < MinPropRefNo Then
                MinPropRefNo = lBSPropertyTypeNo(i)
            End If
        End If
    Next i
    If MinPropRefNo > 0 Then

        objOpenSTAAD.Property.AssignBeamProperty TempBeamArray, MinPropRefNo
    Else
        MinPropRefNo = 0
    End If
ContinueFor:
Next j
objOpenSTAAD.UpdateStructure
MsgBox "Section properties consolidated."
End Sub
