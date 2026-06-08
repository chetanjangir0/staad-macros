Option Explicit

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

    DXFFile = Left$(StaadFile, Len(StaadFile) - 4) & "_Geometry.dxf"

    f = FreeFile

    Open DXFFile For Output As #f

    '--------------------------------------------------
    ' DXF HEADER
    '--------------------------------------------------
    Print #f, "0"
    Print #f, "SECTION"
    Print #f, "2"
    Print #f, "ENTITIES"

    ExportMembersToDXF objOpenSTAAD, f

    Print #f, "0"
    Print #f, "ENDSEC"

    Print #f, "0"
    Print #f, "EOF"

    Close #f

    MsgBox "DXF created successfully:" & vbCrLf & DXFFile, vbInformation

    Set objOpenSTAAD = Nothing

End Sub

Sub ExportMembersToDXF(os As Object, f As Integer)

    Dim nMembers As Long
    Dim BeamNos() As Long

    Dim i As Long
    Dim RetVal As Long

    Dim StartNode As Long
    Dim EndNode As Long

    Dim x1 As Double
    Dim y1 As Double
    Dim z1 As Double

    Dim x2 As Double
    Dim y2 As Double
    Dim z2 As Double

    nMembers = os.Geometry.GetMemberCount()

    If nMembers <= 0 Then Exit Sub

    ReDim BeamNos(nMembers - 1)

    os.Geometry.GetBeamList BeamNos

    For i = 0 To nMembers - 1

        RetVal = os.Geometry.GetMemberIncidence( _
                    BeamNos(i), _
                    StartNode, _
                    EndNode)

        If RetVal = 0 Then

            os.Geometry.GetNodeCoordinates _
                StartNode, _
                x1, y1, z1

            os.Geometry.GetNodeCoordinates _
                EndNode, _
                x2, y2, z2

            WriteDXFLine _
                f, _
                x1, y1, z1, _
                x2, y2, z2

        End If

    Next i

End Sub

Sub WriteDXFLine( _
    f As Integer, _
    x1 As Double, _
    y1 As Double, _
    z1 As Double, _
    x2 As Double, _
    y2 As Double, _
    z2 As Double)

    Print #f, "0"
    Print #f, "LINE"

    Print #f, "8"
    Print #f, "MEMBERS"

    Print #f, "10"
    Print #f, Format$(x1, "0.000")

    Print #f, "20"
    Print #f, Format$(y1, "0.000")

    Print #f, "30"
    Print #f, Format$(z1, "0.000")

    Print #f, "11"
    Print #f, Format$(x2, "0.000")

    Print #f, "21"
    Print #f, Format$(y2, "0.000")

    Print #f, "31"
    Print #f, Format$(z2, "0.000")

End Sub
