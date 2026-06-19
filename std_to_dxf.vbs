Option Explicit

Private Const PI As Double = 3.14159265358979
Private Const MIN_SECTION_HALF_WIDTH As Double = 0.05
Private Const SECTION_WIDTH_SCALE As Double = 1#
Private Const LABEL_HEIGHT_FACTOR As Double = 0.035
Private Const MIN_LABEL_HEIGHT As Double = 0.05
Private Const SHORT_MEMBER_LENGTH As Double = 2#
Private Const LABEL_WIDTH_FACTOR As Double = 0.65
Private Const LABEL_MAX_SPAN_FACTOR As Double = 0.8
Private Const LABEL_SECTION_COLOR As Long = 4
Private Const LABEL_FLANGE_COLOR As Long = 6
Private Const LABEL_LENGTH_COLOR As Long = 2

Private gViewPlane As String
Private gLabelTextScale As Double
Private gWriteLabels As Boolean

Sub Main()

    Dim objOpenSTAAD As Object
    Dim StaadFile As String
    Dim DXFFile As String
    Dim f As Integer
    Dim exportedMembers As Long

    Set objOpenSTAAD = GetObject(, "StaadPro.OpenSTAAD")

    objOpenSTAAD.GetSTAADFile StaadFile, True

    If StaadFile = "" Then
        MsgBox "No STAAD model open.", vbExclamation
        Exit Sub
    End If

    Dim defaultDXF As String
    defaultDXF = Left$(StaadFile, Len(StaadFile) - 4) & "_Geometry_Sections.dxf"
    If Not GetExportSettings(defaultDXF, DXFFile, gViewPlane) Then
        MsgBox "Export cancelled.", vbInformation
        Exit Sub
    End If

    f = FreeFile
    Open DXFFile For Output As #f

    WriteDXFHeader f
    exportedMembers = ExportMembersToDXF(objOpenSTAAD, f)
    WriteDXFFooter f

    Close #f

    If exportedMembers > 0 Then
        MsgBox "DXF created successfully:" & vbCrLf & DXFFile & vbCrLf & _
               "Members exported: " & CStr(exportedMembers), vbInformation
    Else
        MsgBox "DXF was created, but no selected members were exported." & vbCrLf & _
               "Select one or more beam members in STAAD before running this macro." & vbCrLf & _
               DXFFile, vbExclamation
    End If

    Set objOpenSTAAD = Nothing

End Sub

Private Function GetExportSettings(defaultPath As String, ByRef outputPath As String, ByRef viewPlane As String) As Boolean
    Dim fso As Object
    Dim shell As Object
    Dim defaultDir As String
    Dim defaultName As String
    Dim tempDir As String
    Dim htaPath As String
    Dim resultPath As String
    Dim ts As Object
    Dim line As String
    Dim p As Long
    Dim key As String
    Dim value As String
    Dim cancelled As Boolean
    Dim folder As String
    Dim fileName As String

    defaultDir = Left$(defaultPath, InStrRev(defaultPath, "\"))
    defaultName = Mid$(defaultPath, InStrRev(defaultPath, "\") + 1)

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set shell = CreateObject("WScript.Shell")
    tempDir = shell.ExpandEnvironmentStrings("%TEMP%")
    htaPath = fso.BuildPath(tempDir, "staad_dxf_export_settings.hta")
    resultPath = fso.BuildPath(tempDir, "staad_dxf_export_settings.txt")

    If fso.FileExists(resultPath) Then
        fso.DeleteFile resultPath, True
    End If

    WriteSettingsHTA htaPath, resultPath, defaultDir, defaultName
    shell.Run "mshta.exe " & QuoteArg(htaPath), 1, True
    On Error GoTo 0

    If fso Is Nothing Then
        GetExportSettings = False
        Exit Function
    End If

    If Not fso.FileExists(resultPath) Then
        GetExportSettings = False
        Exit Function
    End If

    cancelled = True
    folder = defaultDir
    fileName = defaultName
    viewPlane = "XY"
    gLabelTextScale = 1#
    gWriteLabels = True

    Set ts = fso.OpenTextFile(resultPath, 1, False)
    Do Until ts.AtEndOfStream
        line = ts.ReadLine
        p = InStr(line, "=")
        If p > 0 Then
            key = LCase$(Left$(line, p - 1))
            value = Mid$(line, p + 1)
            Select Case key
                Case "cancelled"
                    cancelled = (value <> "0")
                Case "folder"
                    folder = value
                Case "filename"
                    fileName = value
                Case "plane"
                    viewPlane = UCase$(value)
                Case "textscale"
                    If IsNumeric(value) Then
                        gLabelTextScale = CDbl(value)
                    End If
                Case "labels"
                    gWriteLabels = (value <> "0")
            End Select
        End If
    Loop
    ts.Close

    If cancelled Or Len(Trim$(folder)) = 0 Or Len(Trim$(fileName)) = 0 Then
        GetExportSettings = False
        Exit Function
    End If

    If LCase$(Right$(fileName, 4)) <> ".dxf" Then
        fileName = fileName & ".dxf"
    End If

    If Right$(folder, 1) = "\" Then
        outputPath = folder & fileName
    Else
        outputPath = folder & "\" & fileName
    End If

    If viewPlane <> "XY" And viewPlane <> "YZ" And viewPlane <> "ZX" Then
        viewPlane = "XY"
    End If

    If gLabelTextScale < 0.1 Or gLabelTextScale > 10# Then
        gLabelTextScale = 1#
    End If

    GetExportSettings = True
End Function

Private Sub WriteSettingsHTA(htaPath As String, resultPath As String, defaultDir As String, defaultName As String)
    Dim f As Integer

    f = FreeFile
    Open htaPath For Output As #f

    Print #f, "<html>"
    Print #f, "<head>"
    Print #f, "<title>Export STAAD Members to DXF</title>"
    Print #f, "<HTA:APPLICATION ID=""DXFExportSettings"" APPLICATIONNAME=""STAAD DXF Export"" BORDER=""thin"" CAPTION=""yes"" SHOWINTASKBAR=""yes"" SINGLEINSTANCE=""yes"" SYSMENU=""yes"" WINDOWSTATE=""normal"">"
    Print #f, "<style>"
    Print #f, "body{font-family:Segoe UI,Arial,sans-serif;font-size:13px;margin:18px;color:#202020;background:#f6f6f6;} label{display:block;margin:12px 0 5px;} input[type=text]{width:360px;padding:6px;} button{padding:6px 14px;margin-left:6px;} .row{display:flex;align-items:center;} .planes label{display:inline-block;margin-right:18px;} .actions{text-align:right;margin-top:18px;}"
    Print #f, "</style>"
    Print #f, "<script language=""VBScript"">"
    Print #f, "Option Explicit"
    Print #f, "Const RESULT_PATH = """ & EscapeVBString(resultPath) & """"
    Print #f, "Sub Window_OnLoad"
    Print #f, "  document.getElementById(""folder"").Value = """ & EscapeVBString(RemoveTrailingBackslash(defaultDir)) & """"
    Print #f, "  document.getElementById(""filename"").Value = """ & EscapeVBString(defaultName) & """"
    Print #f, "  document.getElementById(""textscale"").Value = ""1.00"""
    Print #f, "  window.resizeTo 560, 435"
    Print #f, "End Sub"
    Print #f, "Sub btnBrowse_OnClick"
    Print #f, "  Dim sh, fld, startFolder"
    Print #f, "  startFolder = document.getElementById(""folder"").Value"
    Print #f, "  Set sh = CreateObject(""Shell.Application"")"
    Print #f, "  Set fld = sh.BrowseForFolder(0, ""Select output folder for DXF file:"", &H1, startFolder)"
    Print #f, "  If Not fld Is Nothing Then document.getElementById(""folder"").Value = fld.Self.Path"
    Print #f, "End Sub"
    Print #f, "Sub btnExport_OnClick"
    Print #f, "  Dim fso, ts, plane, textScale, labels"
    Print #f, "  plane = ""XY"""
    Print #f, "  If document.getElementById(""planeYZ"").Checked Then plane = ""YZ"""
    Print #f, "  If document.getElementById(""planeZX"").Checked Then plane = ""ZX"""
    Print #f, "  textScale = document.getElementById(""textscale"").Value"
    Print #f, "  labels = ""0"""
    Print #f, "  If document.getElementById(""labels"").Checked Then labels = ""1"""
    Print #f, "  Set fso = CreateObject(""Scripting.FileSystemObject"")"
    Print #f, "  Set ts = fso.CreateTextFile(RESULT_PATH, True)"
    Print #f, "  ts.WriteLine ""cancelled=0"""
    Print #f, "  ts.WriteLine ""folder="" & document.getElementById(""folder"").Value"
    Print #f, "  ts.WriteLine ""filename="" & document.getElementById(""filename"").Value"
    Print #f, "  ts.WriteLine ""plane="" & plane"
    Print #f, "  ts.WriteLine ""textscale="" & textScale"
    Print #f, "  ts.WriteLine ""labels="" & labels"
    Print #f, "  ts.Close"
    Print #f, "  window.close"
    Print #f, "End Sub"
    Print #f, "Sub btnCancel_OnClick"
    Print #f, "  Dim fso, ts"
    Print #f, "  Set fso = CreateObject(""Scripting.FileSystemObject"")"
    Print #f, "  Set ts = fso.CreateTextFile(RESULT_PATH, True)"
    Print #f, "  ts.WriteLine ""cancelled=1"""
    Print #f, "  ts.Close"
    Print #f, "  window.close"
    Print #f, "End Sub"
    Print #f, "</script>"
    Print #f, "</head>"
    Print #f, "<body>"
    Print #f, "<h3>Export Selected Beams to DXF</h3>"
    Print #f, "<label for=""filename"">File name</label>"
    Print #f, "<input id=""filename"" type=""text"">"
    Print #f, "<label for=""folder"">Output folder</label>"
    Print #f, "<div class=""row""><input id=""folder"" type=""text""><button id=""btnBrowse"">Browse...</button></div>"
    Print #f, "<label>View plane mapped to DXF X-Y</label>"
    Print #f, "<div class=""planes"">"
    Print #f, "<label><input id=""planeXY"" name=""plane"" type=""radio"" checked> X-Y</label>"
    Print #f, "<label><input id=""planeYZ"" name=""plane"" type=""radio""> Y-Z</label>"
    Print #f, "<label><input id=""planeZX"" name=""plane"" type=""radio""> Z-X</label>"
    Print #f, "</div>"
    Print #f, "<label><input id=""labels"" type=""checkbox"" checked> Text labels</label>"
    Print #f, "<label for=""textscale"">Text size scale</label>"
    Print #f, "<input id=""textscale"" type=""text"">"
    Print #f, "<div class=""actions""><button id=""btnCancel"">Cancel</button><button id=""btnExport"">Export</button></div>"
    Print #f, "</body>"
    Print #f, "</html>"

    Close #f
End Sub

Private Function EscapeVBString(value As String) As String
    EscapeVBString = Replace$(value, """", """""")
End Function

Private Function RemoveTrailingBackslash(value As String) As String
    If Len(value) > 3 And Right$(value, 1) = "\" Then
        RemoveTrailingBackslash = Left$(value, Len(value) - 1)
    Else
        RemoveTrailingBackslash = value
    End If
End Function

Private Function QuoteArg(value As String) As String
    QuoteArg = """" & Replace$(value, """", """""") & """"
End Function

Private Function ExportMembersToDXF(os As Object, f As Integer) As Long

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

    nMembers = CLng(os.Geometry.GetNoOfSelectedBeams())
    If nMembers <= 0 Then
        Exit Function
    End If

    ReDim BeamNos(nMembers - 1)
    os.Geometry.GetSelectedBeams BeamNos, 1

    For i = 0 To nMembers - 1

        startNode = 0
        endNode = 0
        retVal = os.Geometry.GetMemberIncidence(BeamNos(i), startNode, endNode)

        If HasValidMemberIncidence(startNode, endNode) Then

            os.Geometry.GetNodeCoordinates startNode, x1, y1, z1
            os.Geometry.GetNodeCoordinates endNode, x2, y2, z2

            length = GetMemberLength(os, BeamNos(i), x1, y1, z1, x2, y2, z2)
            MapPointToDXFPlane x1, y1, z1
            MapPointToDXFPlane x2, y2, z2

            sectionName = GetMemberSectionDisplayName(os, BeamNos(i))
            GetMemberVisualHalfWidths os, BeamNos(i), length, startHalfWidth, endHalfWidth, propertyType
            sectionName = FormatTubePipeSectionName(os, BeamNos(i), propertyType, sectionName)

            Dim taperedLabel As String
            Dim pv(23) As Double
            Dim pt As Long
            taperedLabel = ""
            If propertyType = 675 Or propertyType = 680 Then
                os.Property.GetBeamSectionPropertyValuesEx BeamNos(i), pt, pv
                taperedLabel = FormatTaperedILabel(BeamNos(i), os, length, pt, pv)
            End If

            If Len(taperedLabel) > 0 Then
                If length < SHORT_MEMBER_LENGTH Then
                    labelText = CompactTaperedLabel(taperedLabel)
                Else
                    labelText = taperedLabel
                End If
            Else
                labelText = FormatMemberLabel(sectionName, length)
            End If

            WriteDXFLine f, "MEMBER_CENTERLINE", x1, y1, z1, x2, y2, z2, "DASHED"
            WriteMemberEnvelope f, x1, y1, z1, x2, y2, z2, startHalfWidth, endHalfWidth, propertyType, sectionName
            If gWriteLabels Then
                WriteMemberLabel f, x1, y1, z1, x2, y2, z2, labelText, MaxD(startHalfWidth, endHalfWidth)
            End If
            ExportMembersToDXF = ExportMembersToDXF + 1

        End If

    Next i

End Function

Private Sub MapPointToDXFPlane(ByRef x As Double, ByRef y As Double, ByRef z As Double)
    Dim tx As Double
    Dim ty As Double

    Select Case gViewPlane
        Case "YZ"
            tx = z
            ty = y
        Case "ZX"
            tx = z
            ty = x
        Case Else
            tx = x
            ty = y
    End Select

    x = tx
    y = ty
    z = 0#
End Sub

Private Function FormatMemberLabel(sectionName As String, memberLength As Double) As String
    If memberLength < SHORT_MEMBER_LENGTH And Len(sectionName) > 0 Then
        FormatMemberLabel = sectionName
    ElseIf Len(sectionName) > 0 Then
        FormatMemberLabel = sectionName & " (" & FormatNumberSafe(memberLength) & "M)"
    Else
        FormatMemberLabel = "(" & FormatNumberSafe(memberLength) & "M)"
    End If
End Function

Private Function CompactTaperedLabel(labelText As String) As String
    Dim p As Long
    p = InStr(labelText, " (")
    If p > 0 Then
        CompactTaperedLabel = Left$(labelText, p - 1)
    Else
        CompactTaperedLabel = labelText
    End If
End Function

Private Function HasValidMemberIncidence(startNode As Long, endNode As Long) As Boolean
    HasValidMemberIncidence = (startNode > 0 And endNode > 0)
End Function

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

Private Function FormatTaperedILabel(beamNo As Long, os As Object, memberLength As Double, propertyType As Long, propValues() As Double) As String
    Dim d1 As Double, d2 As Double
    Dim bf As Double, tf As Double, tw As Double

    If propertyType = 675 Then
        d1 = Abs(propValues(4))
        d2 = Abs(propValues(5))
        tw = Abs(propValues(1))
        bf = Abs(propValues(6))
        tf = Abs(propValues(7))
    ElseIf propertyType = 680 Then
        d1 = Abs(propValues(0))
        d2 = Abs(propValues(2))
        tw = Abs(propValues(1))
        bf = MaxD(Abs(propValues(3)), Abs(propValues(5)))
        tf = MaxD(Abs(propValues(4)), Abs(propValues(6)))
    Else
        FormatTaperedILabel = ""
        Exit Function
    End If

    d1 = d1 * 1000#
    d2 = d2 * 1000#
    tw = tw * 1000#
    bf = bf * 1000#
    tf = tf * 1000#

    FormatTaperedILabel = "W(" & Format$(d2, "0") & "~" & Format$(d1, "0") & "x" & Format$(tw, "0") & ")\P" & _
                          "2F(" & Format$(bf, "0") & "x" & Format$(tf, "0") & "); (" & FormatNumberSafe(memberLength) & "M)"
End Function

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
    Dim maxTextHeight As Double
    Dim rotationDeg As Double
    Dim lx1 As Double, ly1 As Double, lz1 As Double
    Dim lx2 As Double, ly2 As Double, lz2 As Double

    If x2 < x1 Then
        lx1 = x2 : ly1 = y2 : lz1 = z2
        lx2 = x1 : ly2 = y1 : lz2 = z1
    Else
        lx1 = x1 : ly1 = y1 : lz1 = z1
        lx2 = x2 : ly2 = y2 : lz2 = z2
    End If

    mx = (lx1 + lx2) / 2#
    my = (ly1 + ly2) / 2#
    mz = (lz1 + lz2) / 2#

    memberLength = Distance3D(lx1, ly1, lz1, lx2, ly2, lz2)
    textHeight = MaxD(memberLength * LABEL_HEIGHT_FACTOR, halfWidth * 0.45)
    textHeight = textHeight * gLabelTextScale
    textHeight = MaxD(textHeight, MIN_LABEL_HEIGHT)
    maxTextHeight = GetMaxLabelHeightForSpan(labelText, memberLength)
    If maxTextHeight > 0# Then
        textHeight = MinD(textHeight, maxTextHeight)
        textHeight = MaxD(textHeight, MIN_LABEL_HEIGHT)
    End If

    GetViewOffsetVector lx1, ly1, lz1, lx2, ly2, lz2, ox, oy, oz

    If oy < 0# Then
        ox = -ox
        oy = -oy
        oz = -oz
    End If

    mx = mx + ox * (halfWidth + textHeight * 1.5)
    my = my + oy * (halfWidth + textHeight * 1.5)
    mz = mz + oz * (halfWidth + textHeight * 1.5)

    rotationDeg = Atn2(ly2 - ly1, lx2 - lx1) * 180# / PI

    WriteDXFLabelText f, "MEMBER_LABELS", mx, my, mz, textHeight, rotationDeg, labelText, ox, oy, oz

End Sub

Private Function GetMaxLabelHeightForSpan(labelText As String, memberLength As Double) As Double
    Dim textChars As Long

    textChars = Len(CleanDXFText(labelText))
    If textChars <= 0 Or memberLength <= 0# Then
        GetMaxLabelHeightForSpan = 0#
        Exit Function
    End If

    GetMaxLabelHeightForSpan = (memberLength * LABEL_MAX_SPAN_FACTOR) / (CDbl(textChars) * LABEL_WIDTH_FACTOR)
End Function

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

Private Function FormatTubePipeSectionName(os As Object, beamNo As Long, propertyType As Long, sectionName As String) As String
    Dim propValues(23) As Double
    Dim pt As Long
    Dim formattedName As String
    Dim hasExtendedValues As Boolean

    FormatTubePipeSectionName = sectionName

    If Not IsTubeOrPipeSection(propertyType, sectionName) Then
        Exit Function
    End If

    On Error Resume Next
    os.Property.GetBeamSectionPropertyValuesEx beamNo, pt, propValues
    If Err.Number <> 0 Then
        Err.Clear
        hasExtendedValues = False
    Else
        hasExtendedValues = True
    End If
    On Error GoTo 0

    If hasExtendedValues Then
        Select Case pt
            Case 2
                If IsPipeSectionName(sectionName) Then
                    formattedName = FormatPipeSectionName(propValues(0), propValues(1))
                ElseIf IsTubeSectionName(sectionName) Then
                    formattedName = FormatTubeSectionName(propValues(2), propValues(1), propValues(0))
                End If
            Case 650, 654, 696
                formattedName = FormatTubeSectionName(propValues(1), propValues(2), propValues(3))
            Case 660, 655
                formattedName = FormatPipeSectionName(propValues(1), MaxD(propValues(1) - 2# * propValues(2), 0#))
            Case 695
                formattedName = FormatPipeSectionName(propValues(1), propValues(2))
        End Select
    End If

    If Len(formattedName) = 0 Then
        formattedName = FormatTubePipeFromBeamPropertyAll(os, beamNo, sectionName)
    End If

    If Len(formattedName) > 0 Then
        FormatTubePipeSectionName = formattedName
    End If
End Function

Private Function FormatTubePipeFromBeamPropertyAll(os As Object, beamNo As Long, sectionName As String) As String
    Dim width As Double, depth As Double
    Dim ax As Double, ay As Double, az As Double
    Dim ix As Double, iy As Double, iz As Double
    Dim tf As Double, tw As Double
    Dim thickness As Double
    Dim outerDiameter As Double

    FormatTubePipeFromBeamPropertyAll = ""

    On Error Resume Next
    os.Property.GetBeamPropertyAll beamNo, width, depth, ax, ay, az, ix, iy, iz, tf, tw
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    If IsTubeSectionName(sectionName) Then
        thickness = MaxD(Abs(tf), Abs(tw))
        FormatTubePipeFromBeamPropertyAll = FormatTubeSectionName(Abs(depth), Abs(width), thickness)
    ElseIf IsPipeSectionName(sectionName) Then
        thickness = MaxD(Abs(tf), Abs(tw))
        outerDiameter = MaxD(Abs(depth), Abs(width))
        FormatTubePipeFromBeamPropertyAll = FormatPipeSectionName(outerDiameter, MaxD(outerDiameter - 2# * thickness, 0#))
    End If
End Function

Private Function IsTubeSectionName(sectionName As String) As Boolean
    IsTubeSectionName = (InStr(UCase$(sectionName), "TUBE") > 0 Or _
                         InStr(UCase$(sectionName), "RHS") > 0 Or _
                         InStr(UCase$(sectionName), "SHS") > 0)
End Function

Private Function IsPipeSectionName(sectionName As String) As Boolean
    IsPipeSectionName = (InStr(UCase$(sectionName), "PIPE") > 0 Or _
                         InStr(UCase$(sectionName), "CHS") > 0)
End Function

Private Function FormatTubeSectionName(depth As Double, width As Double, thickness As Double) As String
    If depth <= 0# Or width <= 0# Or thickness <= 0# Then
        FormatTubeSectionName = ""
    Else
        FormatTubeSectionName = "TUBE (" & FormatMillimeters(depth) & "x" & _
                                FormatMillimeters(width) & "x" & _
                                FormatMillimeters(thickness) & ")"
    End If
End Function

Private Function FormatPipeSectionName(outerDiameter As Double, innerDiameter As Double) As String
    If outerDiameter <= 0# Then
        FormatPipeSectionName = ""
    Else
        FormatPipeSectionName = "PIPE (OD " & FormatMillimeters(outerDiameter) & _
                                " ID " & FormatMillimeters(MaxD(innerDiameter, 0#)) & ")"
    End If
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
            d1 = Abs(propValues(0))
            d2 = Abs(propValues(2))
            maxDim = MaxD(d1, d2)
            If d1 > 0# And d2 > 0# Then
                startHalfWidth = MaxD(d1 / 2#, MIN_SECTION_HALF_WIDTH)
                endHalfWidth = MaxD(d2 / 2#, MIN_SECTION_HALF_WIDTH)
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
    Print #f, "62"
    Print #f, "7"
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

Private Sub WriteDXFLabelText( _
    f As Integer, _
    layerName As String, _
    x As Double, y As Double, z As Double, _
    height As Double, _
    rotationDeg As Double, _
    value As String, _
    ox As Double, oy As Double, oz As Double)

    Dim lines() As String
    Dim i As Long
    Dim lineCount As Long
    Dim lineOffset As Double

    If InStr(value, "\P") = 0 Then
        WriteDXFColoredLabelLine f, layerName, x, y, z, height, rotationDeg, value
        Exit Sub
    End If

    lines = Split(value, "\P")
    lineCount = UBound(lines) - LBound(lines) + 1

    For i = LBound(lines) To UBound(lines)
        lineOffset = ((CDbl(lineCount - 1) / 2#) - CDbl(i - LBound(lines))) * height * 1.25
        WriteDXFColoredLabelLine f, layerName, x + ox * lineOffset, y + oy * lineOffset, z + oz * lineOffset, height, rotationDeg, lines(i)
    Next i

End Sub

Private Sub WriteDXFColoredLabelLine( _
    f As Integer, _
    layerName As String, _
    x As Double, y As Double, z As Double, _
    height As Double, _
    rotationDeg As Double, _
    value As String)

    Dim p As Long

    p = InStr(value, "; (")
    If p > 0 Then
        WriteDXFColoredTextPart f, layerName, x, y, z, height, rotationDeg, value, Left$(value, p), 1, LABEL_FLANGE_COLOR
        WriteDXFColoredTextPart f, layerName, x, y, z, height, rotationDeg, value, Mid$(value, p + 1), p + 1, LABEL_LENGTH_COLOR
        Exit Sub
    End If

    p = InStrRev(value, " (")
    If p > 0 Then
        WriteDXFColoredTextPart f, layerName, x, y, z, height, rotationDeg, value, Left$(value, p - 1), 1, LABEL_SECTION_COLOR
        WriteDXFColoredTextPart f, layerName, x, y, z, height, rotationDeg, value, Mid$(value, p), p, LABEL_LENGTH_COLOR
    ElseIf Left$(value, 2) = "2F" Then
        WriteDXFColoredTextPart f, layerName, x, y, z, height, rotationDeg, value, value, 1, LABEL_FLANGE_COLOR
    ElseIf Left$(value, 1) = "(" Then
        WriteDXFColoredTextPart f, layerName, x, y, z, height, rotationDeg, value, value, 1, LABEL_LENGTH_COLOR
    Else
        WriteDXFColoredTextPart f, layerName, x, y, z, height, rotationDeg, value, value, 1, LABEL_SECTION_COLOR
    End If

End Sub

Private Sub WriteDXFColoredTextPart( _
    f As Integer, _
    layerName As String, _
    x As Double, y As Double, z As Double, _
    height As Double, _
    rotationDeg As Double, _
    fullText As String, _
    partText As String, _
    startChar As Long, _
    colorNo As Long)

    Dim cleanFull As String
    Dim cleanPart As String
    Dim offsetChars As Double
    Dim offsetDistance As Double
    Dim angleRad As Double

    cleanFull = CleanDXFText(fullText)
    cleanPart = CleanDXFText(partText)
    If Len(cleanPart) = 0 Then
        Exit Sub
    End If

    offsetChars = (CDbl(startChar - 1) + CDbl(Len(cleanPart)) / 2#) - (CDbl(Len(cleanFull)) / 2#)
    offsetDistance = offsetChars * height * LABEL_WIDTH_FACTOR
    angleRad = rotationDeg * PI / 180#

    WriteDXFTextColor f, layerName, x + Cos(angleRad) * offsetDistance, y + Sin(angleRad) * offsetDistance, z, height, rotationDeg, cleanPart, colorNo

End Sub

Private Sub WriteDXFTextColor( _
    f As Integer, _
    layerName As String, _
    x As Double, y As Double, z As Double, _
    height As Double, _
    rotationDeg As Double, _
    value As String, _
    colorNo As Long)

    Print #f, "0"
    Print #f, "TEXT"
    Print #f, "8"
    Print #f, layerName
    Print #f, "62"
    Print #f, CStr(colorNo)
    Print #f, "10"
    Print #f, FormatDXF(x)
    Print #f, "20"
    Print #f, FormatDXF(y)
    Print #f, "30"
    Print #f, FormatDXF(z)
    Print #f, "40"
    Print #f, FormatDXF(height)
    Print #f, "1"
    Print #f, value
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

Private Function MinD(a As Double, b As Double) As Double
    If a < b Then
        MinD = a
    Else
        MinD = b
    End If
End Function

Private Function FormatDXF(value As Double) As String
    FormatDXF = Replace$(Format$(value, "0.000000"), ",", ".")
End Function

Private Function FormatNumberSafe(value As Double) As String
    FormatNumberSafe = Replace$(Format$(value, "0.000"), ",", ".")
End Function

Private Function FormatMillimeters(value As Double) As String
    FormatMillimeters = Replace$(Format$(Abs(value) * 1000#, "0.###"), ",", ".")
End Function

Private Function CleanDXFText(value As String) As String
    CleanDXFText = Replace$(Replace$(value, vbCr, " "), vbLf, " ")
End Function
