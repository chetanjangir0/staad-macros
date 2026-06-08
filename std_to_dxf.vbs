Option Explicit

Private Const PI As Double = 3.14159265358979
Private Const MIN_SECTION_HALF_WIDTH As Double = 0.05
Private Const SECTION_WIDTH_SCALE As Double = 1#
Private Const LABEL_HEIGHT_FACTOR As Double = 0.035

Sub Main()

    Dim objOpenSTAAD As Object
    Dim StaadFile As String
    Dim DXFFile As String
    Dim f As Integer

    Set objOpenSTAAD = GetObject(, "StaadPro.OpenSTAAD")

    objOpenSTAAD.GetSTAADFile StaadFile, True

    If StaadFile = "" Then
        MsgBox "No STAAD model open.", vbExclamation
        Exit Sub
    End If

    Dim defaultDXF As String
    defaultDXF = Left$(StaadFile, Len(StaadFile) - 4) & "_Geometry_Sections.dxf"
    DXFFile = PickOutputDXFFile(defaultDXF)
    If DXFFile = "" Then
        MsgBox "Export cancelled.", vbInformation
        Exit Sub
    End If

    f = FreeFile
    Open DXFFile For Output As #f

    WriteDXFHeader f
    ExportMembersToDXF objOpenSTAAD, f
    WriteDXFFooter f

    Close #f

    MsgBox "DXF created successfully:" & vbCrLf & DXFFile, vbInformation

    Set objOpenSTAAD = Nothing

End Sub

Private Function PickOutputDXFFile(defaultPath As String) As String
    Dim dlg As Object
    Dim result As String
    result = ""
    On Error Resume Next
    Set dlg = CreateObject("UserAccounts.CommonDialog")
    If Err.Number <> 0 Then
        Err.Clear
        result = InputBox("Enter full output path for DXF file:", "Save DXF As", defaultPath)
        PickOutputDXFFile = result
        Exit Function
    End If
    On Error GoTo 0
    dlg.Filter = "DXF Files (*.dxf)|*.dxf|All Files (*.*)|*.*"
    dlg.FilterIndex = 1
    dlg.InitDir = Left$(defaultPath, InStrRev(defaultPath, "\"))
    dlg.FileName = Mid$(defaultPath, InStrRev(defaultPath, "\") + 1)
    dlg.Flags = &H800
    If dlg.ShowSave Then
        result = dlg.FileName
        If LCase$(Right$(result, 4)) <> ".dxf" Then
            result = result & ".dxf"
        End If
    End If
    PickOutputDXFFile = result
End Function

Private Sub ExportMembersToDXF(os As Object, f As Integer)

    Dim nMembers As Long
    Dim BeamNos() As Long
    Dim i As Long
    Dim retVal As Long
    Dim startNode As Long
    Dim endNode As Long
    Dim x1 As Double, y1 As Double, z1 As Double
    Dim x2 As Double, y2 As Double, z2 As Double
    Dim length As Double
    Dim sectionName As String
    Dim propertyType As Long
    Dim startHalfWidth As Double
    Dim endHalfWidth As Double
    Dim labelText As String

    nMembers = CLng(os.Geometry.GetMemberCount())
    If nMembers <= 0 Then
        Exit Sub
    End If

    ReDim BeamNos(nMembers - 1)
    os.Geometry.GetBeamList BeamNos

    For i = 0 To nMembers - 1

        retVal = os.Geometry.GetMemberIncidence(BeamNos(i), startNode, endNode)

        If retVal = 0 Then

            os.Geometry.GetNodeCoordinates startNode, x1, y1, z1
            os.Geometry.GetNodeCoordinates endNode, x2, y2, z2

            length = GetMemberLength(os, BeamNos(i), x1, y1, z1, x2, y2, z2)
            sectionName = GetMemberSectionDisplayName(os, BeamNos(i))
            GetMemberVisualHalfWidths os, BeamNos(i), length, startHalfWidth, endHalfWidth, propertyType

            labelText = "M" & CStr(BeamNos(i))
            If Len(sectionName) > 0 Then
                labelText = labelText & " | " & sectionName
            End If
            labelText = labelText & " | L=" & FormatNumberSafe(length)

            WriteDXFLine f, "MEMBER_CENTERLINE", x1, y1, z1, x2, y2, z2, "DASHED"
            WriteMemberEnvelope f, x1, y1, z1, x2, y2, z2, startHalfWidth, endHalfWidth, propertyType, sectionName
            WriteMemberLabel f, x1, y1, z1, x2, y2, z2, labelText, MaxD(startHalfWidth, endHalfWidth)

        End If

    Next i

End Sub

Private Sub WriteMemberEnvelope( _
    f As Integer, _
    x1 As Double, y1 As Double, z1 As Double, _
    x2 As Double, y2 As Double, z2 As Double, _
    startHalfWidth As Double, _
    endHalfWidth As Double, _
    propertyType As Long, _
    sectionName As String)

    Dim ox1 As Double, oy1 As Double, oz1 As Double
    Dim ox2 As Double, oy2 As Double, oz2 As Double
    Dim layerName As String

    GetViewOffsetVector x1, y1, z1, x2, y2, z2, ox1, oy1, oz1

    ox2 = ox1 * endHalfWidth
    oy2 = oy1 * endHalfWidth
    oz2 = oz1 * endHalfWidth

    ox1 = ox1 * startHalfWidth
    oy1 = oy1 * startHalfWidth
    oz1 = oz1 * startHalfWidth

    layerName = "MEMBER_SECTION"
    If IsTubeOrPipeSection(propertyType, sectionName) Then
        layerName = "TUBE_PIPE_SECTION"
    End If

    If IsTaperedSection(propertyType, sectionName, startHalfWidth, endHalfWidth) Then
        layerName = "TAPERED_SECTION"
    End If

    WriteDXFLine f, layerName, x1 + ox1, y1 + oy1, z1 + oz1, x2 + ox2, y2 + oy2, z2 + oz2, "CONTINUOUS"
    WriteDXFLine f, layerName, x1 - ox1, y1 - oy1, z1 - oz1, x2 - ox2, y2 - oy2, z2 - oz2, "CONTINUOUS"
    WriteDXFLine f, layerName, x1 + ox1, y1 + oy1, z1 + oz1, x1 - ox1, y1 - oy1, z1 - oz1, "CONTINUOUS"
    WriteDXFLine f, layerName, x2 + ox2, y2 + oy2, z2 + oz2, x2 - ox2, y2 - oy2, z2 - oz2, "CONTINUOUS"

    If IsTubeOrPipeSection(propertyType, sectionName) Then
        WriteDXFLine f, "TUBE_PIPE_SECTION", x1 + ox1 * 0.65, y1 + oy1 * 0.65, z1 + oz1 * 0.65, x2 + ox2 * 0.65, y2 + oy2 * 0.65, z2 + oz2 * 0.65, "HIDDEN"
        WriteDXFLine f, "TUBE_PIPE_SECTION", x1 - ox1 * 0.65, y1 - oy1 * 0.65, z1 - oz1 * 0.65, x2 - ox2 * 0.65, y2 - oy2 * 0.65, z2 - oz2 * 0.65, "HIDDEN"
    End If

End Sub

Private Sub WriteMemberLabel( _
    f As Integer, _
    x1 As Double, y1 As Double, z1 As Double, _
    x2 As Double, y2 As Double, z2 As Double, _
    labelText As String, _
    halfWidth As Double)

    Dim mx As Double, my As Double, mz As Double
    Dim ox As Double, oy As Double, oz As Double
    Dim memberLength As Double
    Dim textHeight As Double
    Dim rotationDeg As Double

    mx = (x1 + x2) / 2#
    my = (y1 + y2) / 2#
    mz = (z1 + z2) / 2#

    memberLength = Distance3D(x1, y1, z1, x2, y2, z2)
    textHeight = MaxD(memberLength * LABEL_HEIGHT_FACTOR, halfWidth * 0.45)
    textHeight = MaxD(textHeight, 0.1)

    GetViewOffsetVector x1, y1, z1, x2, y2, z2, ox, oy, oz
    mx = mx + ox * (halfWidth + textHeight * 1.5)
    my = my + oy * (halfWidth + textHeight * 1.5)
    mz = mz + oz * (halfWidth + textHeight * 1.5)

    rotationDeg = Atn2(y2 - y1, x2 - x1) * 180# / PI

    WriteDXFText f, "MEMBER_LABELS", mx, my, mz, textHeight, rotationDeg, labelText

End Sub

Private Function GetMemberLength( _
    os As Object, _
    beamNo As Long, _
    x1 As Double, y1 As Double, z1 As Double, _
    x2 As Double, y2 As Double, z2 As Double) As Double

    On Error GoTo FallbackLength
    GetMemberLength = CDbl(os.Geometry.GetBeamLength(beamNo))
    If GetMemberLength > 0# Then
        Exit Function
    End If

FallbackLength:
    Err.Clear
    GetMemberLength = Distance3D(x1, y1, z1, x2, y2, z2)
End Function

Private Function GetMemberSectionDisplayName(os As Object, beamNo As Long) As String
    On Error Resume Next
    GetMemberSectionDisplayName = CStr(os.Property.GetBeamSectionDisplayName(beamNo))
    If Err.Number <> 0 Or Len(GetMemberSectionDisplayName) = 0 Then
        Err.Clear
        GetMemberSectionDisplayName = CStr(os.Property.GetBeamSectionName(beamNo))
    End If
    If Err.Number <> 0 Then
        Err.Clear
        GetMemberSectionDisplayName = "NO SECTION"
    End If
    On Error GoTo 0
End Function

Private Sub GetMemberVisualHalfWidths( _
    os As Object, _
    beamNo As Long, _
    memberLength As Double, _
    ByRef startHalfWidth As Double, _
    ByRef endHalfWidth As Double, _
    ByRef propertyType As Long)

    Dim width As Double, depth As Double
    Dim ax As Double, ay As Double, az As Double
    Dim ix As Double, iy As Double, iz As Double
    Dim tf As Double, tw As Double
    Dim propValues(23) As Double
    Dim candidate As Double

    startHalfWidth = MaxD(memberLength * 0.0125, MIN_SECTION_HALF_WIDTH)
    endHalfWidth = startHalfWidth
    propertyType = 0

    On Error Resume Next
    os.Property.GetBeamPropertyAll beamNo, width, depth, ax, ay, az, ix, iy, iz, tf, tw
    If Err.Number = 0 Then
        candidate = MaxD(width, depth) * SECTION_WIDTH_SCALE / 2#
        If candidate > 0# Then
            startHalfWidth = MaxD(candidate, MIN_SECTION_HALF_WIDTH)
            endHalfWidth = startHalfWidth
        End If
    Else
        Err.Clear
    End If

    os.Property.GetBeamSectionPropertyValuesEx beamNo, propertyType, propValues
    If Err.Number = 0 Then
        GetEnvelopeFromPropertyValues propertyType, propValues, startHalfWidth, endHalfWidth
    Else
        Err.Clear
    End If
    On Error GoTo 0

End Sub

Private Sub GetEnvelopeFromPropertyValues( _
    propertyType As Long, _
    propValues() As Double, _
    ByRef startHalfWidth As Double, _
    ByRef endHalfWidth As Double)

    Dim d1 As Double
    Dim d2 As Double
    Dim bf As Double
    Dim maxDim As Double

    On Error GoTo Done

    Select Case propertyType
        Case 610, 611, 612, 613, 614, 615, 630, 631, 633, 620, 656
            d1 = Abs(propValues(1))
            bf = Abs(propValues(2))
            maxDim = MaxD(d1, bf)

        Case 616
            d1 = Abs(propValues(0))
            bf = Abs(propValues(1))
            maxDim = MaxD(d1, bf)

        Case 640, 641, 642, 643, 644, 645, 646, 650, 654, 662, 663, 664, 666
            d1 = Abs(propValues(1))
            bf = Abs(propValues(2))
            maxDim = MaxD(d1, bf)

        Case 660, 655, 668, 695
            maxDim = Abs(propValues(1))

        Case 675
            d1 = Abs(propValues(5))
            d2 = Abs(propValues(4))
            maxDim = MaxD(d1, d2)
            If d1 > 0# And d2 > 0# Then
                startHalfWidth = MaxD(d1 / 2#, MIN_SECTION_HALF_WIDTH)
                endHalfWidth = MaxD(d2 / 2#, MIN_SECTION_HALF_WIDTH)
                Exit Sub
            End If

        Case 680
            maxDim = MaxFirstValues(propValues, 0, 6)
            If Abs(propValues(0)) > 0# And Abs(propValues(1)) > 0# Then
                startHalfWidth = MaxD(Abs(propValues(1)) / 2#, MIN_SECTION_HALF_WIDTH)
                endHalfWidth = MaxD(Abs(propValues(0)) / 2#, MIN_SECTION_HALF_WIDTH)
                Exit Sub
            End If

        Case 671
            maxDim = Abs(propValues(4))

        Case 672, 674, 673, 699
            d1 = Abs(propValues(4))
            d2 = Abs(propValues(5))
            maxDim = MaxD(d1, d2)

        Case 676
            d1 = Abs(propValues(6))
            d2 = Abs(propValues(7))
            maxDim = MaxD(d1, d2)

        Case 690, 691, 694, 696, 697
            d1 = Abs(propValues(1))
            bf = Abs(propValues(3))
            maxDim = MaxD(d1, bf)

        Case 692, 693
            d1 = Abs(propValues(1))
            bf = Abs(propValues(2))
            maxDim = MaxD(d1, bf)

        Case 698
            maxDim = MaxFirstValues(propValues, 0, 6)
    End Select

    If maxDim > 0# Then
        startHalfWidth = MaxD(maxDim / 2#, MIN_SECTION_HALF_WIDTH)
        endHalfWidth = startHalfWidth
    End If

Done:
    Err.Clear
End Sub

Private Function MaxFirstValues(propValues() As Double, firstIndex As Long, lastIndex As Long) As Double
    Dim i As Long
    For i = firstIndex To lastIndex
        MaxFirstValues = MaxD(MaxFirstValues, Abs(propValues(i)))
    Next i
End Function

Private Function IsTubeOrPipeSection(propertyType As Long, sectionName As String) As Boolean
    Dim s As String
    s = UCase$(sectionName)

    IsTubeOrPipeSection = False

    If InStr(s, "TUBE") > 0 Then
        IsTubeOrPipeSection = True
    End If

    If InStr(s, "PIPE") > 0 Then
        IsTubeOrPipeSection = True
    End If

    If InStr(s, "CHS") > 0 Then
        IsTubeOrPipeSection = True
    End If

    If InStr(s, "RHS") > 0 Then
        IsTubeOrPipeSection = True
    End If

    If InStr(s, "SHS") > 0 Then
        IsTubeOrPipeSection = True
    End If

    If propertyType = 650 Then
        IsTubeOrPipeSection = True
    End If

    If propertyType = 654 Then
        IsTubeOrPipeSection = True
    End If

    If propertyType = 655 Then
        IsTubeOrPipeSection = True
    End If

    If propertyType = 660 Then
        IsTubeOrPipeSection = True
    End If

    If propertyType = 675 Then
        IsTubeOrPipeSection = True
    End If

    If propertyType = 695 Then
        IsTubeOrPipeSection = True
    End If

    If propertyType = 696 Then
        IsTubeOrPipeSection = True
    End If
End Function

Private Function IsTaperedSection(propertyType As Long, sectionName As String, startHalfWidth As Double, endHalfWidth As Double) As Boolean
    Dim s As String

    s = UCase$(sectionName)
    IsTaperedSection = False

    If InStr(s, "TAPER") > 0 Then
        IsTaperedSection = True
    End If

    If Abs(startHalfWidth - endHalfWidth) > MaxD(startHalfWidth, endHalfWidth) * 0.1 Then
        IsTaperedSection = True
    End If

    If propertyType = 675 Then
        IsTaperedSection = True
    End If

    If propertyType = 680 Then
        IsTaperedSection = True
    End If
End Function

Private Sub GetViewOffsetVector( _
    x1 As Double, y1 As Double, z1 As Double, _
    x2 As Double, y2 As Double, z2 As Double, _
    ByRef ox As Double, ByRef oy As Double, ByRef oz As Double)

    Dim dx As Double, dy As Double, dz As Double
    Dim lengthXY As Double

    dx = x2 - x1
    dy = y2 - y1
    dz = z2 - z1

    lengthXY = Sqr(dx * dx + dy * dy)

    If lengthXY > 0.000001 Then
        ox = -dy / lengthXY
        oy = dx / lengthXY
        oz = 0#
    Else
        ox = 1#
        oy = 0#
        oz = 0#
    End If

End Sub

Private Sub WriteDXFHeader(f As Integer)

    Print #f, "0"
    Print #f, "SECTION"
    Print #f, "2"
    Print #f, "HEADER"
    Print #f, "9"
    Print #f, "$ACADVER"
    Print #f, "1"
    Print #f, "AC1009"
    Print #f, "0"
    Print #f, "ENDSEC"

    Print #f, "0"
    Print #f, "SECTION"
    Print #f, "2"
    Print #f, "TABLES"

    WriteLinetypeTable f
    WriteLayerTable f

    Print #f, "0"
    Print #f, "ENDSEC"

    Print #f, "0"
    Print #f, "SECTION"
    Print #f, "2"
    Print #f, "ENTITIES"

End Sub

Private Sub WriteDXFFooter(f As Integer)
    Print #f, "0"
    Print #f, "ENDSEC"
    Print #f, "0"
    Print #f, "EOF"
End Sub

Private Sub WriteLinetypeTable(f As Integer)
    Print #f, "0"
    Print #f, "TABLE"
    Print #f, "2"
    Print #f, "LTYPE"
    Print #f, "70"
    Print #f, "3"

    WriteLinetype f, "CONTINUOUS", "Solid line", 0, 0#
    WriteLinetype f, "DASHED", "Dashed centerline", 2, 0.75
    Print #f, "49"
    Print #f, "0.50"
    Print #f, "49"
    Print #f, "-0.25"

    WriteLinetype f, "HIDDEN", "Hidden inner tube line", 2, 0.45
    Print #f, "49"
    Print #f, "0.25"
    Print #f, "49"
    Print #f, "-0.20"

    Print #f, "0"
    Print #f, "ENDTAB"
End Sub

Private Sub WriteLinetype(f As Integer, name As String, description As String, elementCount As Long, patternLength As Double)
    Print #f, "0"
    Print #f, "LTYPE"
    Print #f, "2"
    Print #f, name
    Print #f, "70"
    Print #f, "0"
    Print #f, "3"
    Print #f, description
    Print #f, "72"
    Print #f, "65"
    Print #f, "73"
    Print #f, CStr(elementCount)
    Print #f, "40"
    Print #f, FormatDXF(patternLength)
End Sub

Private Sub WriteLayerTable(f As Integer)
    Print #f, "0"
    Print #f, "TABLE"
    Print #f, "2"
    Print #f, "LAYER"
    Print #f, "70"
    Print #f, "5"

    WriteLayer f, "MEMBER_CENTERLINE", 8, "DASHED"
    WriteLayer f, "MEMBER_SECTION", 3, "CONTINUOUS"
    WriteLayer f, "TAPERED_SECTION", 1, "CONTINUOUS"
    WriteLayer f, "TUBE_PIPE_SECTION", 5, "CONTINUOUS"
    WriteLayer f, "MEMBER_LABELS", 7, "CONTINUOUS"

    Print #f, "0"
    Print #f, "ENDTAB"
End Sub

Private Sub WriteLayer(f As Integer, name As String, colorNo As Long, lineTypeName As String)
    Print #f, "0"
    Print #f, "LAYER"
    Print #f, "2"
    Print #f, name
    Print #f, "70"
    Print #f, "0"
    Print #f, "62"
    Print #f, CStr(colorNo)
    Print #f, "6"
    Print #f, lineTypeName
End Sub

Private Sub WriteDXFLine( _
    f As Integer, _
    layerName As String, _
    x1 As Double, y1 As Double, z1 As Double, _
    x2 As Double, y2 As Double, z2 As Double, _
    lineTypeName As String)

    Print #f, "0"
    Print #f, "LINE"
    Print #f, "8"
    Print #f, layerName
    Print #f, "6"
    Print #f, lineTypeName
    Print #f, "10"
    Print #f, FormatDXF(x1)
    Print #f, "20"
    Print #f, FormatDXF(y1)
    Print #f, "30"
    Print #f, FormatDXF(z1)
    Print #f, "11"
    Print #f, FormatDXF(x2)
    Print #f, "21"
    Print #f, FormatDXF(y2)
    Print #f, "31"
    Print #f, FormatDXF(z2)

End Sub

Private Sub WriteDXFText( _
    f As Integer, _
    layerName As String, _
    x As Double, y As Double, z As Double, _
    height As Double, _
    rotationDeg As Double, _
    value As String)

    Print #f, "0"
    Print #f, "TEXT"
    Print #f, "8"
    Print #f, layerName
    Print #f, "10"
    Print #f, FormatDXF(x)
    Print #f, "20"
    Print #f, FormatDXF(y)
    Print #f, "30"
    Print #f, FormatDXF(z)
    Print #f, "40"
    Print #f, FormatDXF(height)
    Print #f, "1"
    Print #f, CleanDXFText(value)
    Print #f, "50"
    Print #f, FormatDXF(rotationDeg)
    Print #f, "72"
    Print #f, "1"
    Print #f, "11"
    Print #f, FormatDXF(x)
    Print #f, "21"
    Print #f, FormatDXF(y)
    Print #f, "31"
    Print #f, FormatDXF(z)

End Sub

Private Function Distance3D( _
    x1 As Double, y1 As Double, z1 As Double, _
    x2 As Double, y2 As Double, z2 As Double) As Double

    Distance3D = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2 + (z2 - z1) ^ 2)
End Function

Private Function Atn2(y As Double, x As Double) As Double
    If Abs(x) < 0.0000001 Then
        If y >= 0# Then
            Atn2 = PI / 2#
        Else
            Atn2 = -PI / 2#
        End If
    Else
        Atn2 = Atn(y / x)
        If x < 0# Then
            Atn2 = Atn2 + PI
        End If
    End If
End Function

Private Function MaxD(a As Double, b As Double) As Double
    If a > b Then
        MaxD = a
    Else
        MaxD = b
    End If
End Function

Private Function FormatDXF(value As Double) As String
    FormatDXF = Replace$(Format$(value, "0.000000"), ",", ".")
End Function

Private Function FormatNumberSafe(value As Double) As String
    FormatNumberSafe = Replace$(Format$(value, "0.000"), ",", ".")
End Function

Private Function CleanDXFText(value As String) As String
    CleanDXFText = Replace$(Replace$(value, vbCr, " "), vbLf, " ")
End Function
