'/*--------------------------------------------------------------------------------------+
'|  Load Combination Generator for STAAD.Pro for asce 7 16 ASD
'+--------------------------------------------------------------------------------------*/
Option Explicit

'==============================================================================
' ENTRY POINT
'==============================================================================

Dim gNextCombNum As Long 'numbering tracker for load combos

Sub Main()
    Dim staadObj As Object
    Dim strFileName As String
    Dim bIncludePath As Boolean

    Set staadObj = GetObject(, "StaadPro.OpenSTAAD")

    bIncludePath = True
    staadObj.GetSTAADFile strFileName, bIncludePath

    If strFileName = "" Then
        MsgBox "Error - Please open a STAAD.Pro model before running this macro.", vbOkOnly
        End
    End If

    ShowCategoryDialog staadObj

    Set staadObj = Nothing
End Sub


'==============================================================================
' STEP 1 : Show dialog so user can assign categories to each load case
'==============================================================================
Sub ShowCategoryDialog(staad As Object)

    '--------------------------------------------------------------------------
    ' Read all primary load cases from the model
    '--------------------------------------------------------------------------
    Dim nPrimary As Long
    nPrimary = staad.Load.GetPrimaryLoadCaseCount()

    If nPrimary < 1 Then
        MsgBox "No primary load cases found in the model.", vbOkOnly
        End
    End If

    Dim nLCNums() As Long
    ReDim nLCNums(nPrimary - 1)
    staad.Load.GetPrimaryLoadCaseNumbers nLCNums()

    '--------------------------------------------------------------------------
    ' Auto-detect categories
    '--------------------------------------------------------------------------
    Dim nCatChoice() As Integer
    ReDim nCatChoice(nPrimary - 1)
    Dim i As Integer

    For i = 0 To nPrimary - 1
        Dim nType As Integer
        nType = staad.Load.GetLoadType(nLCNums(i))
        If nType > 100 Then nType = nType \ 101
        Select Case nType
            Case 0  : nCatChoice(i) = 1   ' Dead
            Case 1  : nCatChoice(i) = 2   ' Live
            Case 3  : nCatChoice(i) = 3   ' Wind
            Case 4  : nCatChoice(i) = 4   ' Seismic
            Case 2  : nCatChoice(i) = 5   ' Roof Live
            Case Else : nCatChoice(i) = 0
        End Select
    Next i

    '--------------------------------------------------------------------------
    ' Show detected load cases summary
    '--------------------------------------------------------------------------
    Dim sInfo As String
    sInfo = "Load cases detected:" & Chr(13)
    For i = 0 To nPrimary - 1
        Dim sCat As String
        Select Case nCatChoice(i)
            Case 1 : sCat = "Dead Load"
            Case 2 : sCat = "Live Load"
            Case 3 : sCat = "Wind Load"
            Case 4 : sCat = "Seismic"
            Case 5 : sCat = "Roof Live"
            Case Else : sCat = "(Skip)"
        End Select
        sInfo = sInfo & "  LC" & nLCNums(i) & "  ->  " & sCat & Chr(13)
    Next i
    MsgBox sInfo, vbOkOnly, "Auto-Detected Categories"

    '--------------------------------------------------------------------------
    ' Simple dialog - start number only
    '--------------------------------------------------------------------------
    Begin Dialog UserDialog 300, 95, "Load Combination Generator"
        Text    20, 14, 180, 14, "Start Combination Number:", .LblStart
        TextBox 210, 11, 70, 21,                              .TxtStart
        OKButton     60, 55, 80, 21
        CancelButton 160, 55, 80, 21
    End Dialog

    Dim dlg As UserDialog
    dlg.TxtStart = "101"

    Dim iBtn As Integer
    iBtn = Dialog(dlg)
    If iBtn = 0 Then End

    Dim nStartComb As Integer
    nStartComb = CInt(Val(dlg.TxtStart))
    If nStartComb < 1 Then nStartComb = 101

    '--------------------------------------------------------------------------
    ' Sort load cases into category buckets
    '--------------------------------------------------------------------------
    Dim DL_LC() As Long
    ReDim DL_LC(nPrimary)
    Dim LL_LC() As Long
    ReDim LL_LC(nPrimary)
    Dim WL_LC() As Long
    ReDim WL_LC(nPrimary)
    Dim EQ_LC() As Long
    ReDim EQ_LC(nPrimary)
    Dim WL_Lbl() As String
    ReDim WL_Lbl(nPrimary)
    Dim EQ_Lbl() As String
    ReDim EQ_Lbl(nPrimary)
    Dim RL_LC() As Long
    ReDim RL_LC(nPrimary)

    Dim nDL As Integer, nLL As Integer, nWL As Integer, nEQ As Integer, nRL As Integer
    nDL = 0 : nLL = 0 : nWL = 0 : nEQ = 0 : nRL = 0

    For i = 0 To nPrimary - 1
        Select Case nCatChoice(i)
            Case 1
                DL_LC(nDL) = nLCNums(i)
                nDL = nDL + 1
            Case 2
                LL_LC(nLL) = nLCNums(i)
                nLL = nLL + 1
            Case 3
                WL_LC(nWL) = nLCNums(i)
                WL_Lbl(nWL) = "LC" & nLCNums(i)
                nWL = nWL + 1
            Case 4
                EQ_LC(nEQ) = nLCNums(i)
                EQ_Lbl(nEQ) = "LC" & nLCNums(i)
                nEQ = nEQ + 1
            Case 5
                RL_LC(nRL) = nLCNums(i)
                nRL = nRL + 1
        End Select
    Next i

    If nDL = 0 Then
        MsgBox "No Dead Load case detected. Please check load type assignments.", vbOkOnly
        Exit Sub
    End If

    gNextCombNum = nStartComb
    GenerateCombinations staad, _
        nDL, DL_LC(), _
        nLL, LL_LC(), _
        nWL, WL_LC(), WL_Lbl(), _
        nEQ, EQ_LC(), EQ_Lbl(), _
        nRL, RL_LC(), _
        nStartComb

End Sub

'==============================================================================
' STEP 2 : Create all load combinations in STAAD.Pro
'
'  1.  1 DL
'  2.  1 DL + 1 LL
'  3.  1 DL + 1 RL
'  4.  1 DL + 0.75 LL + 0.75 RL
'  5.  1 DL + 0.6 WL              (looped over each WL case)
'  6.  1 DL + 0.75 LL + 0.75 RL + 0.45 WL  (looped over each WL case)
'  7.  0.6 DL + 0.6 WL            (looped over each WL case)
'  8.  1 DL + 0.7 EL              (looped over each EL case)
'  9.  1 DL + 0.75 LL + 0.525 EL  (looped over each EL case)
' 10.  0.6 DL + 0.7 EL            (looped over each EL case)
'==============================================================================
Sub GenerateCombinations(staad As Object, _
    nDL As Integer, DL_LC() As Long, _
    nLL As Integer, LL_LC() As Long, _
    nWL As Integer, WL_LC() As Long, WL_Lbl() As String, _
    nEQ As Integer, EQ_LC() As Long, EQ_Lbl() As String, _
    nRL As Integer, RL_LC() As Long, _
    nStart As Integer)

    Dim newComb As Long
    Dim iDL As Integer, iLL As Integer, iWL As Integer, iEQ As Integer, iRL As Integer
    Dim CombName As String

    '--------------------------------------------------------------------------
    ' Helper: build the "DL" part of a label, using CL tag when nDL > 1
    '--------------------------------------------------------------------------
    ' (inline where needed)

    '==========================================================================
    ' C1 : 1 DL
    '==========================================================================
    newComb = NextComb(staad)
    CombName = "1 DL"
    If nDL > 1 Then CombName = "1 DL + 1 CL"
    staad.Load.CreateNewLoadCombination CombName, newComb
    For iDL = 0 To nDL - 1
        staad.Load.AddLoadAndFactorToCombination newComb, DL_LC(iDL), 1
    Next iDL

    '==========================================================================
    ' C2 : 1 DL + 1 LL
    '==========================================================================
    If nLL > 0 Then
        newComb = NextComb(staad)
        CombName = "1 DL"
        If nDL > 1 Then CombName = "1 DL + 1 CL"
        CombName = CombName & " + 1 LL"
        staad.Load.CreateNewLoadCombination CombName, newComb
        For iDL = 0 To nDL - 1
            staad.Load.AddLoadAndFactorToCombination newComb, DL_LC(iDL), 1
        Next iDL
        For iLL = 0 To nLL - 1
            staad.Load.AddLoadAndFactorToCombination newComb, LL_LC(iLL), 1
        Next iLL
    End If

    '==========================================================================
    ' C3 : 1 DL + 1 RL
    '==========================================================================
    If nRL > 0 Then
        newComb = NextComb(staad)
        CombName = "1 DL"
        If nDL > 1 Then CombName = "1 DL + 1 CL"
        CombName = CombName & " + 1 RL"
        staad.Load.CreateNewLoadCombination CombName, newComb
        For iDL = 0 To nDL - 1
            staad.Load.AddLoadAndFactorToCombination newComb, DL_LC(iDL), 1
        Next iDL
        For iRL = 0 To nRL - 1
            staad.Load.AddLoadAndFactorToCombination newComb, RL_LC(iRL), 1
        Next iRL
    End If

    '==========================================================================
    ' C4 : 1 DL + 0.75 LL + 0.75 RL
    '==========================================================================
    If nLL > 0 Or nRL > 0 Then
        newComb = NextComb(staad)
        CombName = "1 DL"
        If nDL > 1 Then CombName = "1 DL + 1 CL"
        If nLL > 0 Then
            CombName = CombName & " + 0.75 LL"
        End If
        If nRL > 0 Then
            CombName = CombName & " + 0.75 RL"
        End If
        staad.Load.CreateNewLoadCombination CombName, newComb
        For iDL = 0 To nDL - 1
            staad.Load.AddLoadAndFactorToCombination newComb, DL_LC(iDL), 1
        Next iDL
        For iLL = 0 To nLL - 1
            staad.Load.AddLoadAndFactorToCombination newComb, LL_LC(iLL), 0.75
        Next iLL
        For iRL = 0 To nRL - 1
            staad.Load.AddLoadAndFactorToCombination newComb, RL_LC(iRL), 0.75
        Next iRL
    End If

    '==========================================================================
    ' C5 : 1 DL + 0.6 WL   (looped over each WL direction)
    '==========================================================================
    For iWL = 0 To nWL - 1
        newComb = NextComb(staad)
        CombName = "1 DL"
        If nDL > 1 Then CombName = "1 DL + 1 CL"
        CombName = CombName & " + 0.6 WL" & (iWL + 1)
        staad.Load.CreateNewLoadCombination CombName, newComb
        For iDL = 0 To nDL - 1
            staad.Load.AddLoadAndFactorToCombination newComb, DL_LC(iDL), 1
        Next iDL
        staad.Load.AddLoadAndFactorToCombination newComb, WL_LC(iWL), 0.6
    Next iWL

    '==========================================================================
    ' C6 : 1 DL + 0.75 LL + 0.75 RL + 0.45 WL   (looped over each WL direction)
    '==========================================================================

        For iWL = 0 To nWL - 1
            newComb = NextComb(staad)
            CombName = "1 DL"
            If nDL > 1 Then CombName = "1 DL + 1 CL"
            If nLL > 0 Then
                CombName = CombName & " + 0.75 LL"
            End If
            If nRL > 0 Then
                CombName = CombName & " + 0.75 RL"
            End If
            CombName = CombName & " + 0.45 WL" & (iWL + 1)
            staad.Load.CreateNewLoadCombination CombName, newComb
            For iDL = 0 To nDL - 1
                staad.Load.AddLoadAndFactorToCombination newComb, DL_LC(iDL), 1
            Next iDL
            For iLL = 0 To nLL - 1
                staad.Load.AddLoadAndFactorToCombination newComb, LL_LC(iLL), 0.75
            Next iLL
            For iRL = 0 To nRL - 1
                staad.Load.AddLoadAndFactorToCombination newComb, RL_LC(iRL), 0.75
            Next iRL
            staad.Load.AddLoadAndFactorToCombination newComb, WL_LC(iWL), 0.45
        Next iWL

    '==========================================================================
    ' C7 : 0.6 DL + 0.6 WL   (looped over each WL direction)
    '==========================================================================
    For iWL = 0 To nWL - 1
        newComb = NextComb(staad)
        CombName = "0.6 DL"
        If nDL > 1 Then CombName = "0.6 DL + 0.6 CL"
        CombName = CombName & " + 0.6 WL" & (iWL + 1)
        staad.Load.CreateNewLoadCombination CombName, newComb
        For iDL = 0 To nDL - 1
            staad.Load.AddLoadAndFactorToCombination newComb, DL_LC(iDL), 0.6
        Next iDL
        staad.Load.AddLoadAndFactorToCombination newComb, WL_LC(iWL), 0.6
    Next iWL

    '==========================================================================
    ' C8 : 1 DL + 0.7 EL   (looped over each EL case)
    '==========================================================================
    For iEQ = 0 To nEQ - 1
        newComb = NextComb(staad)
        CombName = "1 DL"
        If nDL > 1 Then CombName = "1 DL + 1 CL"
        CombName = CombName & " + 0.7 EL" & (iEQ + 1)
        staad.Load.CreateNewLoadCombination CombName, newComb
        For iDL = 0 To nDL - 1
            staad.Load.AddLoadAndFactorToCombination newComb, DL_LC(iDL), 1
        Next iDL
        staad.Load.AddLoadAndFactorToCombination newComb, EQ_LC(iEQ), 0.7
    Next iEQ

    '==========================================================================
    ' C9 : 1 DL + 0.75 LL + 0.525 EL   (looped over each EL case)
    '==========================================================================

        For iEQ = 0 To nEQ - 1
            newComb = NextComb(staad)
            CombName = "1 DL"
            If nDL > 1 Then CombName = "1 DL + 1 CL"
            If nLL > 0 Then
                CombName = CombName & " + 0.75 LL"
            End If
            CombName = CombName & " + 0.525 EL" & (iEQ + 1)
            staad.Load.CreateNewLoadCombination CombName, newComb
            For iDL = 0 To nDL - 1
                staad.Load.AddLoadAndFactorToCombination newComb, DL_LC(iDL), 1
            Next iDL
            For iLL = 0 To nLL - 1
                staad.Load.AddLoadAndFactorToCombination newComb, LL_LC(iLL), 0.75
            Next iLL
            staad.Load.AddLoadAndFactorToCombination newComb, EQ_LC(iEQ), 0.525
        Next iEQ

    '==========================================================================
    ' C10 : 0.6 DL + 0.7 EL   (looped over each EL case)
    '==========================================================================
    For iEQ = 0 To nEQ - 1
        newComb = NextComb(staad)
        CombName = "0.6 DL"
        If nDL > 1 Then CombName = "0.6 DL + 0.6 CL"
        CombName = CombName & " + 0.7 EL" & (iEQ + 1)
        staad.Load.CreateNewLoadCombination CombName, newComb
        For iDL = 0 To nDL - 1
            staad.Load.AddLoadAndFactorToCombination newComb, DL_LC(iDL), 0.6
        Next iDL
        staad.Load.AddLoadAndFactorToCombination newComb, EQ_LC(iEQ), 0.7
    Next iEQ

    '==========================================================================
    ' Done
    '==========================================================================
    MsgBox "Load combination(s) generated successfully." & Chr(13) & _
           "Combinations : " & nStart & " to " & (gNextCombNum - 1) & Chr(13) & Chr(13) & _
           "Please verify the combinations in your STAAD.Pro model.", vbOkOnly

End Sub

Function NextComb(staad As Object) As Long
    NextComb = gNextCombNum
    gNextCombNum = gNextCombNum + 1
End Function
