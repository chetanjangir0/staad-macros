Sub Main()
'DESCRIPTION: Create 3D Portal Frame Structure with Main Frames and End Frames
'             Frames repeat in Z direction; connecting beams at column top (eave) level
'             User inputs: frame geometry, number of bays, bay spacings, support types

    '-------------------------------------------------------------------------
    ' DIALOG DEFINITION
    '-------------------------------------------------------------------------
    Begin Dialog UserDialog 660, 420, "3D Portal Frame Generator"
        ' --- Frame Geometry ---
        Text 20, 15, 200, 15, "--- Frame Geometry ---", .lblGeom
        Text 20, 35, 180, 15, "Frame Width (m):", .lblWidth
        TextBox 210, 33, 100, 15, .width

        Text 20, 57, 180, 15, "Eave Height (m):", .lblHt
        TextBox 210, 55, 100, 15, .ht

        Text 20, 79, 180, 15, "Roof Slope 1:x  (x=):", .lblSlope
        TextBox 210, 77, 100, 15, .slope

        ' --- Bay Configuration ---
        Text 20, 105, 200, 15, "--- Bay Configuration ---", .lblBay
        Text 20, 125, 180, 15, "Total Number of Bays:", .lblNMain
        TextBox 210, 123, 100, 15, .nMainBays

        Text 20, 147, 180, 15, "Main Bay Spacing (m):", .lblMainSp
        TextBox 210, 145, 100, 15, .mainSpacing

        Text 20, 169, 180, 15, "End Bay Spacing (m):", .lblEndSp
        TextBox 210, 167, 100, 15, .endSpacing

        ' --- Support Types ---
        Text 20, 200, 200, 15, "--- Support Conditions ---", .lblSupp
        Text 20, 220, 180, 15, "Left Column Support:", .lblS1
        OptionGroup .sprt1
            OptionButton 210, 218, 80, 15, "Fixed", .ob1
            OptionButton 310, 218, 80, 15, "Pinned", .ob2

        Text 20, 242, 180, 15, "Right Column Support:", .lblS2
        OptionGroup .sprt2
            OptionButton 210, 240, 80, 15, "Fixed", .ob3
            OptionButton 310, 240, 80, 15, "Pinned", .ob4

        ' --- Options ---
        Text 20, 270, 200, 15, "--- Options ---", .lblOpt
        CheckBox 20, 290, 300, 15, "Add Connecting Beams at Eave Level (purlins/ties)", .addConnBeams
        CheckBox 20, 310, 300, 15, "Add Ridge Connecting Beams", .addRidgeBeams

        ' --- Buttons ---
        OKButton     450, 380, 90, 25
        CancelButton 555, 380, 90, 25

        ' --- Info ---
        Text 20, 345, 620, 15, "Total frames = totalBays + 1  |  2 end bays + (totalBays-2) main bays", .lblInfo
        Text 20, 362, 620, 15, "Total Z length = 2*endSpacing + (totalBays-2)*mainSpacing", .lblInfo2
    End Dialog

    Dim dlg As UserDialog
    Dim dlgResult As Integer

    ' --- Default Values ---
    dlg.width       = "20"
    dlg.ht          = "7"
    dlg.slope       = "5"
    dlg.nMainBays   = "6"
    dlg.mainSpacing = "6"
    dlg.endSpacing  = "3"
    dlg.addConnBeams = 1
    dlg.addRidgeBeams = 1

    dlgResult = Dialog(dlg)
    Debug.Clear

    If dlgResult = -1 Then  ' OK pressed

        '--- Read Inputs ---
        Dim fw       As Double  ' frame width (X direction)
        Dim eh       As Double  ' eave height
        Dim rs       As Double  ' roof slope denominator (1:rs)
        Dim totalBays As Long    ' total bays including 2 end bays
        Dim nMain    As Long    ' derived: main bays = totalBays - 2
        Dim mSp      As Double  ' main bay spacing (Z)
        Dim eSp      As Double  ' end bay spacing (Z)
        Dim sp1      As String  ' left support
        Dim sp2      As String  ' right support
        Dim bConn    As Boolean ' add eave connecting beams
        Dim bRidge   As Boolean ' add ridge connecting beams

        fw    = Abs(CDbl(dlg.width))
        eh    = Abs(CDbl(dlg.ht))
        rs    = Abs(CDbl(dlg.slope))
        totalBays = Abs(CLng(dlg.nMainBays))
        nMain = totalBays - 2
        mSp   = Abs(CDbl(dlg.mainSpacing))
        eSp   = Abs(CDbl(dlg.endSpacing))
        sp1   = CStr(dlg.sprt1)   ' "0"=Fixed, "1"=Pinned
        sp2   = CStr(dlg.sprt2)
        bConn  = (dlg.addConnBeams = 1)
        bRidge = (dlg.addRidgeBeams = 1)

        If rs = 0 Then
            MsgBox "Roof slope cannot be zero.", vbOKOnly, "Input Error"
            Exit Sub
        End If
        If totalBays < 2 Then
            MsgBox "Total bays must be at least 2 (one end bay on each side).", vbOKOnly, "Input Error"
            Exit Sub
        End If

        '--- Geometry Calculations ---
        Dim ridgeHt As Double
        ridgeHt = eh + (fw / (2# * rs))   ' ridge height

        ' Total number of frames = totalBays + 1
        '   Frame 0            = end frame  (z = 0)
        '   Frames 1..nMain    = main frames (z = eSp, eSp+mSp, ..., eSp+(nMain-1)*mSp)
        '   Frame nMain+1      = main frame  (z = eSp + nMain*mSp)
        '   Frame nMain+2      = end frame   (z = eSp + nMain*mSp + eSp)
        Dim nFrames As Long
        nFrames = totalBays + 1   ' bays + 1 = frames (e.g. 6 bays -> 7 frames)

        ' Z coordinates of each frame
        Dim zCoord() As Double
        ReDim zCoord(0 To nFrames - 1)
        Dim i As Long
        ' Frame indices:
        '   0              -> End Frame 1   z = 0
        '   1..nFrames-2   -> Main Frames   z = eSp, eSp+mSp, ..., eSp+(totalBays-2)*mSp
        '   nFrames-1      -> End Frame 2   z = last main frame z + eSp
        '
        ' Example: totalBays=6, eSp=3, mSp=6
        '   Frame 0:z=0, 1:z=3, 2:z=9, 3:z=15, 4:z=21, 5:z=27, 6:z=30  (6 bays, 7 frames)
        zCoord(0) = 0#                          ' End Frame 1
        Dim zCur As Double
        zCur = eSp
        For i = 1 To nFrames - 2               ' all interior (main) frames
            zCoord(i) = zCur
            zCur = zCur + mSp
        Next i
        zCoord(nFrames - 1) = zCoord(nFrames - 2) + eSp  ' End Frame 2

        ' Each frame has 5 nodes:
        '   Local 1: base left    (0,   0,      z)
        '   Local 2: base right   (fw,  0,      z)
        '   Local 3: eave left    (0,   eh,     z)
        '   Local 4: eave right   (fw,  eh,     z)
        '   Local 5: ridge        (fw/2, ridgeHt, z)

        '--- Connect to STAAD ---
        Dim objOpenSTAAD As Object
        Set objOpenSTAAD = GetObject(, "StaadPro.OpenSTAAD")
        objOpenSTAAD.SetInputUnits 4, 5  ' Meters, KiloNewtons

        Dim geometry As Object
        Set geometry = objOpenSTAAD.Geometry

        '=========================================================
        ' STEP 1: ADD ALL NODES
        '=========================================================
        ' Node numbering: frame i (0-based) has nodes (i*5+1) to (i*5+5)
        '   node i*5+1 = base left
        '   node i*5+2 = base right
        '   node i*5+3 = eave left
        '   node i*5+4 = eave right
        '   node i*5+5 = ridge

        Dim nf As Long
        For nf = 0 To nFrames - 1
            Dim z As Double
            z = zCoord(nf)
            geometry.AddNode 0,      0,        z   ' base left
            geometry.AddNode fw,     0,        z   ' base right
            geometry.AddNode 0,      eh,       z   ' eave left
            geometry.AddNode fw,     eh,       z   ' eave right
            geometry.AddNode fw / 2, ridgeHt,  z   ' ridge
        Next nf

        Debug.Print "Nodes added: "; nFrames * 5

        '=========================================================
        ' STEP 2: ADD FRAME MEMBERS (columns + rafters per frame)
        '=========================================================
        ' Member numbering within each frame (4 members per frame):
        '   m*4+1: left column   (base left -> eave left)
        '   m*4+2: right column  (base right -> eave right)
        '   m*4+3: left rafter   (eave left -> ridge)
        '   m*4+4: right rafter  (eave right -> ridge)

        For nf = 0 To nFrames - 1
            Dim baseNode As Long
            baseNode = nf * 5 + 1
            ' Nodes: baseNode=basL, +1=basR, +2=eaveL, +3=eaveR, +4=ridge
            geometry.AddBeam baseNode,     baseNode + 2   ' left column
            geometry.AddBeam baseNode + 1, baseNode + 3   ' right column
            geometry.AddBeam baseNode + 2, baseNode + 4   ' left rafter
            geometry.AddBeam baseNode + 3, baseNode + 4   ' right rafter
        Next nf

        Debug.Print "Frame members added: "; nFrames * 4

        '=========================================================
        ' STEP 3: ADD CONNECTING BEAMS BETWEEN FRAMES (purlins/ties)
        '=========================================================
        ' Connect adjacent frames at:
        '   - Eave left  (node offset +2)
        '   - Eave right (node offset +3)
        '   - Ridge      (node offset +4)  if bRidge=True
        ' And optionally:
        '   - Base left  (node offset 0) - usually not needed but can add
        '   - Base right (node offset 1) - usually not needed

        If bConn Or bRidge Then
            For nf = 0 To nFrames - 2
                Dim n1Base As Long, n2Base As Long
                n1Base = nf * 5 + 1        ' first frame base node
                n2Base = (nf + 1) * 5 + 1  ' next frame base node

                If bConn Then
                    ' Eave left tie beam
                    geometry.AddBeam n1Base + 2, n2Base + 2
                    ' Eave right tie beam
                    geometry.AddBeam n1Base + 3, n2Base + 3
                End If

                If bRidge Then
                    ' Ridge purlin
                    geometry.AddBeam n1Base + 4, n2Base + 4
                End If
            Next nf

            Dim connCount As Long
            connCount = 0
            If bConn Then connCount = connCount + (nFrames - 1) * 2
            If bRidge Then connCount = connCount + (nFrames - 1)
            Debug.Print "Connecting beams added: "; connCount
        End If

        '=========================================================
        ' STEP 4: ASSIGN SUPPORTS
        '=========================================================
        Dim support As Object
        Set support = objOpenSTAAD.Support

        Dim s1 As Long, s2 As Long

        If sp1 = "0" Then
            s1 = support.CreateSupportFixed()
        ElseIf sp1 = "1" Then
            s1 = support.CreateSupportPinned()
        Else
            MsgBox "Invalid left support type.", vbOKOnly, "Error"
            Exit Sub
        End If

        If sp2 = "0" Then
            s2 = support.CreateSupportFixed()
        ElseIf sp2 = "1" Then
            s2 = support.CreateSupportPinned()
        Else
            MsgBox "Invalid right support type.", vbOKOnly, "Error"
            Exit Sub
        End If

        ' Assign supports to base nodes of every frame
        For nf = 0 To nFrames - 1
            Dim bL As Long, bR As Long
            bL = nf * 5 + 1   ' base left node
            bR = nf * 5 + 2   ' base right node
            support.AssignSupportToNode bL, s1
            support.AssignSupportToNode bR, s2
        Next nf

        Debug.Print "Supports assigned to "; nFrames * 2; " base nodes"

        '=========================================================
        ' SUMMARY
        '=========================================================
        Dim totalNodes   As Long
        Dim totalMembers As Long
        totalNodes   = nFrames * 5
        totalMembers = nFrames * 4
        If bConn  Then totalMembers = totalMembers + (nFrames - 1) * 2
        If bRidge Then totalMembers = totalMembers + (nFrames - 1)

        Dim zTotal As Double
        zTotal = eSp + nMain * mSp + eSp

        Debug.Print "================================================"
        Debug.Print "3D Frame Generation Complete"
        Debug.Print "Frame Width         = "; fw; " m"
        Debug.Print "Eave Height         = "; eh; " m"
        Debug.Print "Ridge Height        = "; ridgeHt; " m"
        Debug.Print "Roof Slope          = 1:"; rs
        Debug.Print "Total Z Length      = "; zTotal; " m"
        Debug.Print "End Bay Spacing     = "; eSp; " m"
        Debug.Print "Main Bay Spacing    = "; mSp; " m"
        Debug.Print "Total Bays (input)  = "; totalBays; " (2 end + "; nMain; " main)"
        Debug.Print "Number of Frames    = "; nFrames
        Debug.Print "  - End Frames      = 2 (frames 1 and "; nFrames; ")"
        Debug.Print "  - Main Frames     = "; nMain + 1
        Debug.Print "Total Nodes         = "; totalNodes
        Debug.Print "Total Members       = "; totalMembers
        Debug.Print "================================================"

        MsgBox "3D Frame generated successfully!" & vbCrLf & _
               "Frames: " & nFrames & "  |  Nodes: " & totalNodes & "  |  Members: " & totalMembers & vbCrLf & _
               "Total Building Length (Z): " & zTotal & " m", _
               vbInformation, "Done"

    ElseIf dlgResult = 0 Then
        Debug.Print "Cancelled."
    End If

End Sub
