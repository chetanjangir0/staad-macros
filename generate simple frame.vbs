Sub Main()
'DESCRIPTION: Create simple 2D PEB frame with optional intermediate columns
'             Interior columns can be defined by COUNT or by SPACING list
'             Optional brick wall height adds nodes on outer columns

    Begin Dialog UserDialog 640,335,"2D Frame Model" ' %GRID:5,5,1,1
        Text 20,20,200,15,"Width (m):",.Text1
        Text 20,45,200,15,"Eave Height (m):",.Text2
        Text 20,70,200,15,"Roof Slope x (1:x):",.Text3
        Text 20,95,200,15,"Brick Wall Height (m):",.Text4
        TextBox 230,20,140,15,.width
        TextBox 230,45,140,15,.ht
        TextBox 230,70,140,15,.slope
        TextBox 230,92,140,15,.brickHt

        GroupBox 15,115,610,105,"Interior Column Definition",.grp1
        OptionGroup .colMode
            OptionButton 30,135,200,15,"By Count (evenly spaced)",.optCount
            OptionButton 30,160,230,15,"By Spacing (cumulative, e.g. 4,5,6)",.optSpacing
        Text 30,185,200,15,"Count  OR  Spacing list:",.TextMode
        TextBox 230,182,380,15,.colInput

        Text 20,235,200,15,"Left Support Type:",.Text5
        Text 20,270,200,15,"Right Support Type:",.Text6
        OptionGroup .sprt1
            OptionButton 230,232,90,15,"Fixed",.OptionButton1
            OptionButton 380,232,90,15,"Pinned",.OptionButton2
        OptionGroup .sprt2
            OptionButton 230,267,90,15,"Fixed",.OptionButton3
            OptionButton 380,267,90,15,"Pinned",.OptionButton4

        OKButton     390,303,90,20
        CancelButton 520,303,90,20
    End Dialog

    Dim dlg As UserDialog
    Dim dlgResult As Integer

    dlg.width    = "20"
    dlg.ht       = "7"
    dlg.slope    = "5"
    dlg.brickHt  = "0"
    dlg.colMode  = 0
    dlg.colInput = "0"

    dlgResult = Dialog(dlg)
    Debug.Clear

    If dlgResult = -1 Then

        Dim fw   As Double
        Dim eh   As Double
        Dim rs   As Double
        Dim bwh  As Double
        Dim sp1  As String
        Dim sp2  As String

        fw  = Abs(CDbl(dlg.width))
        eh  = Abs(CDbl(dlg.ht))
        rs  = Abs(CDbl(dlg.slope))
        bwh = Abs(CDbl(dlg.brickHt))
        sp1 = CStr(dlg.sprt1)
        sp2 = CStr(dlg.sprt2)

        ' ── Validate brick wall height ───────────────────────────────────
        If bwh >= eh Then
            MsgBox "Brick wall height (" & CStr(bwh) & " m) must be less than eave height (" & CStr(eh) & " m).", _
                   vbOKOnly, "Error"
            Exit Sub
        End If

        ' ── Parse interior column X-positions ───────────────────────────
        Dim xPositions() As Double
        Dim nIC As Integer

        If dlg.colMode = 0 Then
            Dim countVal As Integer
            countVal = CInt(Abs(CDbl(dlg.colInput)))
            nIC = countVal
            If nIC > 0 Then
                ReDim xPositions(1 To nIC)
                Dim k As Integer
                For k = 1 To nIC
                    xPositions(k) = CDbl(k) * fw / CDbl(nIC + 1)
                Next k
            End If
            Debug.Print "Mode: Count = "; nIC

        ElseIf dlg.colMode = 1 Then
            Dim rawInput As String
            Dim tokens() As String
            rawInput = Trim(dlg.colInput)

            If rawInput = "" Or rawInput = "0" Then
                nIC = 0
            Else
                tokens = Split(rawInput, ",")
                nIC = UBound(tokens) - LBound(tokens) + 1
                ReDim xPositions(1 To nIC)
                Dim cumX    As Double
                Dim t       As Integer
                Dim spacVal As Double
                cumX = 0
                For t = LBound(tokens) To UBound(tokens)
                    spacVal = Abs(CDbl(Trim(tokens(t))))
                    cumX = cumX + spacVal
                    If cumX >= fw Then
                        MsgBox "Spacing error: cumulative distance " & CStr(cumX) & _
                               " m exceeds or equals frame width " & CStr(fw) & " m.", _
                               vbOKOnly, "Error"
                        Exit Sub
                    End If
                    xPositions(t - LBound(tokens) + 1) = cumX
                Next t
            End If
            Debug.Print "Mode: Spacing"
        End If

        ' ── Connect to STAAD ────────────────────────────────────────────
        Dim objOpenSTAAD As Object
        Set objOpenSTAAD = GetObject(,"StaadPro.OpenSTAAD")
        objOpenSTAAD.SetInputUnits 4, 5

        Dim geometry As Object
        Set geometry = objOpenSTAAD.Geometry

        Dim ridgeH As Double
        ridgeH = fw / (2 * rs)

        ' ── STEP 1: Main corner nodes (1-5) ─────────────────────────────
        geometry.AddNode 0,    0,           0   ' Node 1 - base left
        geometry.AddNode fw,   0,           0   ' Node 2 - base right
        geometry.AddNode 0,    eh,          0   ' Node 3 - eave left
        geometry.AddNode fw,   eh,          0   ' Node 4 - eave right
        geometry.AddNode fw/2, eh + ridgeH, 0   ' Node 5 - ridge
        Debug.Print "Main nodes added (1-5)"

        ' ── STEP 2: Brick wall nodes on outer columns (optional) ─────────
        ' If bwh > 0, add one node per outer column at x=0 and x=fw
        ' at height bwh. These will be used to split the column beams.
        Dim bwNodeL As Long   ' brick wall node - left column
        Dim bwNodeR As Long   ' brick wall node - right column
        Dim hasBW   As Boolean
        hasBW = (bwh > 0.0001)

        Dim nextNodeNum As Long
        nextNodeNum = 6   ' nodes 1-5 already used

        If hasBW Then
            geometry.AddNode 0,  bwh, 0   ' left brick wall node
            bwNodeL = nextNodeNum
            nextNodeNum = nextNodeNum + 1

            geometry.AddNode fw, bwh, 0   ' right brick wall node
            bwNodeR = nextNodeNum
            nextNodeNum = nextNodeNum + 1

            Debug.Print "Brick wall nodes: Left=Node "; bwNodeL; " Right=Node "; bwNodeR; " at h="; bwh
        End If

        ' ── STEP 3: Interior column nodes ───────────────────────────────
        Dim i       As Integer
        Dim xPos    As Double
        Dim rafterY As Double

        Dim atRidge()  As Boolean
        Dim topNode()  As Long
        Dim baseNode() As Long

        If nIC > 0 Then
            ReDim atRidge(1 To nIC)
            ReDim topNode(1 To nIC)
            ReDim baseNode(1 To nIC)
        End If

        For i = 1 To nIC
            xPos = xPositions(i)

            geometry.AddNode xPos, 0, 0
            baseNode(i) = nextNodeNum
            nextNodeNum = nextNodeNum + 1

            If Abs(xPos - fw / 2) < 0.0001 Then
                atRidge(i) = True
                topNode(i) = 5
                Debug.Print "Int. col "; i; " x="; xPos; " -> AT RIDGE, top=Node 5, base=Node "; baseNode(i)
            Else
                atRidge(i) = False
                If xPos < fw / 2 Then
                    rafterY = eh + xPos / rs
                Else
                    rafterY = eh + (fw - xPos) / rs
                End If
                geometry.AddNode xPos, rafterY, 0
                topNode(i) = nextNodeNum
                nextNodeNum = nextNodeNum + 1
                Debug.Print "Int. col "; i; " x="; xPos; " rafterY="; rafterY; _
                            " base=Node "; baseNode(i); " top=Node "; topNode(i)
            End If
        Next i

        ' ── STEP 4: Beams ───────────────────────────────────────────────
        ' 4a. Main outer columns (split at brick wall node if bwh > 0)
        If hasBW Then
            ' Left column: base -> bw node -> eave
            geometry.AddBeam 1, bwNodeL
            geometry.AddBeam bwNodeL, 3
            ' Right column: base -> bw node -> eave
            geometry.AddBeam 2, bwNodeR
            geometry.AddBeam bwNodeR, 4
            Debug.Print "Outer columns split at brick wall nodes"
        Else
            geometry.AddBeam 1, 3
            geometry.AddBeam 2, 4
        End If

        ' 4b. Interior columns (base -> top node)
        For i = 1 To nIC
            geometry.AddBeam baseNode(i), topNode(i)
            Debug.Print "Int. col beam "; i; ": Node "; baseNode(i); " -> Node "; topNode(i)
        Next i

        ' 4c. Left rafter: Node 3 -> [left-half tops, excl. ridge] -> Node 5
        Dim prevNode As Long
        prevNode = 3
        For i = 1 To nIC
            If xPositions(i) < fw / 2 Then
                geometry.AddBeam prevNode, topNode(i)
                prevNode = topNode(i)
            End If
        Next i
        geometry.AddBeam prevNode, 5
        Debug.Print "Left rafter segments added"

        ' 4d. Right rafter: Node 5 -> [right-half tops, excl. ridge] -> Node 4
        prevNode = 5
        For i = 1 To nIC
            If xPositions(i) > fw / 2 Then
                geometry.AddBeam prevNode, topNode(i)
                prevNode = topNode(i)
            End If
        Next i
        geometry.AddBeam prevNode, 4
        Debug.Print "Right rafter segments added"

        ' ── STEP 5: Supports ────────────────────────────────────────────
        Dim support As Object
        Set support = objOpenSTAAD.Support

        Dim s1  As Long
        Dim s2  As Long
        Dim sIC As Long

        If sp1 = "0" Then
            s1 = support.CreateSupportFixed()
        ElseIf sp1 = "1" Then
            s1 = support.CreateSupportPinned()
        Else
            MsgBox "Select proper support type for Left Support", vbOKOnly, "Error"
            Exit Sub
        End If

        If sp2 = "0" Then
            s2 = support.CreateSupportFixed()
        ElseIf sp2 = "1" Then
            s2 = support.CreateSupportPinned()
        Else
            MsgBox "Select proper support type for Right Support", vbOKOnly, "Error"
            Exit Sub
        End If

        support.AssignSupportToNode 1, s1
        support.AssignSupportToNode 2, s2

        If nIC > 0 Then
            sIC = s1
            For i = 1 To nIC
                support.AssignSupportToNode baseNode(i), sIC
                Debug.Print "Support assigned to base Node "; baseNode(i)
            Next i
        End If

        Debug.Print "Script completed successfully"

    ElseIf dlgResult = 0 Then
        Debug.Print "Cancel button pressed"
    End If

End Sub
