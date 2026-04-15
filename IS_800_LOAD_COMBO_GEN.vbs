'/*--------------------------------------------------------------------------------------+
'|  IS 800:2007 Load Combination Generator for STAAD.Pro
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
            Case 4  : nCatChoice(i) = 4   ' Seismic H
            Case 2  : nCatChoice(i) = 5   ' Roof Live
            Case 19  : nCatChoice(i) = 6   ' Crane hook
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
            Case 5 : sCat = "Roof live"
            Case 6 : sCat = "Crane"
            Case Else : sCat = "(Skip)"
        End Select
        sInfo = sInfo & "  LC" & nLCNums(i) & "  ->  " & sCat & Chr(13)
    Next i
    MsgBox sInfo, vbOkOnly, "Auto-Detected Categories"

    '--------------------------------------------------------------------------
    ' Simple dialog - start number only
    '--------------------------------------------------------------------------
    Begin Dialog UserDialog 300, 120, "IS 800:2007 Combination Generator"
        Text    20, 14, 180, 14, "Start Strength Combination Number:", .LblStart
        TextBox 210, 11, 70, 21,                                  .TxtStart
        Text    20, 44, 180, 14, "Start Serviceability Combination Number:", .LblSLS
        TextBox 210, 41, 70, 21,                                  .TxtSLS
        OKButton     60, 80, 80, 21
        CancelButton 160, 80, 80, 21
    End Dialog
    
    Dim dlg As UserDialog
    dlg.TxtStart = "101"
    dlg.TxtSLS   = "201"

    Dim iBtn As Integer
    iBtn = Dialog(dlg)
    If iBtn = 0 Then End

    Dim nStartComb As Integer
    Dim nStartSLS  As Integer
    nStartComb = CInt(Val(dlg.TxtStart))
    nStartSLS  = CInt(Val(dlg.TxtSLS))
    If nStartComb < 1 Then nStartComb = 101
    If nStartSLS  < 1 Then nStartSLS  = 201

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
    Dim CRL_LC() As Long
    ReDim CRL_LC(nPrimary)
    Dim CRL_Lbl() As String
    ReDim CRL_Lbl(nPrimary)

    Dim nDL As Integer, nLL As Integer, nWL As Integer, nEQ As Integer, nRL As Integer, nCRL As Integer
    nDL = 0 : nLL = 0 : nWL = 0 : nEQ = 0 : nRL = 0 : nCRL = 0

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
            Case 6
                CRL_LC(nCRL) = nLCNums(i)
                CRL_Lbl(nCRL) = "LC" & nLCNums(i)
                nCRL = nCRL + 1
        End Select
    Next i

    If nDL = 0 Then
        MsgBox "No Dead Load case detected. Please check load type assignments.", vbOkOnly
        Exit Sub
    End If

    gNextCombNum = nStartComb   ' <-- seed ULS counter
    GenerateCombinations staad, _
        nDL, DL_LC(), _
        nLL, LL_LC(), _
        nWL, WL_LC(), WL_Lbl(), _
        nEQ, EQ_LC(), EQ_Lbl(), _
        nRL, RL_LC(), _
        nCRL, CRL_LC(), CRL_Lbl(), _
        nStartComb, _
        nStartSLS

End Sub

'==============================================================================
' STEP 2 : Create all IS 800:2007 Table 4 combinations in STAAD.Pro
'==============================================================================
Sub GenerateCombinations(staad As Object, _
    nDL As Integer, DL_LC() As Long, _
    nLL As Integer, LL_LC() As Long, _
    nWL As Integer, WL_LC() As Long, WL_Lbl() As String, _
    nEQ As Integer, EQ_LC() As Long, EQ_Lbl() As String, _
    nRL As Integer, RL_LC() As Long, _
    nCRL As Integer, CRL_LC() As Long, CRL_Lbl() As String, _
    nStart As Integer, _
    nStartSLS As Integer)
    
    Dim newComb As Long

    '--------------------------------------------------------------------------
    ' Helper variables
    '--------------------------------------------------------------------------

    Dim iDL As Integer, iLL As Integer, iWL As Integer, iEQ As Integer
    Dim CombName As String

    Dim iRL As Integer
    Dim iCRL As Integer
    Dim iLead As Integer

    '==========================================================================
' C1 : 1.5 DL + Leading Live (LL / RL / CRL)
'      Implements IS800 leading / accompanying rule
'==========================================================================

For iLead = 1 To 3

    'Skip if that load type does not exist
    If iLead = 1 And nLL = 0 Then GoTo SkipLead
    If iLead = 2 And nRL = 0 Then GoTo SkipLead
    If iLead = 3 And nCRL = 0 Then GoTo SkipLead

    CombName = "1.5DL"
    If nDL > 1 Then
        CombName="1.5DL + 1.5CL"
    End If

    If nLL > 0 Then
        If iLead = 1 Then
            CombName = CombName & " + 1.5LL"
        Else
            CombName = CombName & " + 1.05LL"
        End If
    End If

    If nRL > 0 Then
        If iLead = 2 Then
            CombName = CombName & " + 1.5RL"
        Else
            CombName = CombName & " + 1.05RL"
        End If
    End If

    '-------------------------------------------------
    ' IF Crane load
    '-------------------------------------------------
    If nCRL > 0 Then
    For iCRL = 0 To nCRL - 1
        newComb = NextComb(staad)
        If iLead = 3 Then
            staad.Load.CreateNewLoadCombination CombName & " + 1.5CRL" & iCRL + 1 & "", newComb
        Else
            staad.Load.CreateNewLoadCombination CombName & " + 1.05CRL" & iCRL + 1 & "", newComb
        End If


        'DL
        For iDL = 0 To nDL - 1
            staad.Load.AddLoadAndFactorToCombination newComb, DL_LC(iDL), 1.5
        Next iDL

        'LL
        For iLL = 0 To nLL - 1
            If iLead = 1 Then
                staad.Load.AddLoadAndFactorToCombination newComb, LL_LC(iLL), 1.5
            Else
                staad.Load.AddLoadAndFactorToCombination newComb, LL_LC(iLL), 1.05
            End If
        Next iLL

        'RL
        For iRL = 0 To nRL - 1
            If iLead = 2 Then
                staad.Load.AddLoadAndFactorToCombination newComb, RL_LC(iRL), 1.5
            Else
                staad.Load.AddLoadAndFactorToCombination newComb, RL_LC(iRL), 1.05
            End If
        Next iRL

        'Single crane load
        If iLead = 3 Then
            staad.Load.AddLoadAndFactorToCombination newComb, CRL_LC(iCRL), 1.5
        Else
            staad.Load.AddLoadAndFactorToCombination newComb, CRL_LC(iCRL), 1.05
        End If

    Next iCRL
Else
    newComb = NextComb(staad)
    staad.Load.CreateNewLoadCombination CombName, newComb

    '-------------------------------------------------
    ' Dead loads
    '-------------------------------------------------
    For iDL = 0 To nDL - 1
        staad.Load.AddLoadAndFactorToCombination newComb, DL_LC(iDL), 1.5
    Next iDL

    '-------------------------------------------------
    ' Live load
    '-------------------------------------------------
    For iLL = 0 To nLL - 1
        If iLead = 1 Then
            staad.Load.AddLoadAndFactorToCombination newComb, LL_LC(iLL), 1.5
        Else
            staad.Load.AddLoadAndFactorToCombination newComb, LL_LC(iLL), 1.05
        End If
    Next iLL

    '-------------------------------------------------
    ' Roof live
    '-------------------------------------------------
    For iRL = 0 To nRL - 1
        If iLead = 2 Then
            staad.Load.AddLoadAndFactorToCombination newComb, RL_LC(iRL), 1.5
        Else
            staad.Load.AddLoadAndFactorToCombination newComb, RL_LC(iRL), 1.05
        End If
    Next iRL
End If

SkipLead:
Next iLead

        ''==========================================================================
    '' C2 : 1.2 DL  +  1.2 leading + 1.05 accomp + 0.6 WL/EL
    ''      All DL cases combined with all LL cases  (one combined combination)
    ''==========================================================================


Call GenerateLateralCombos(staad, _
    nDL, nLL, nRL, nCRL, nWL, _
    DL_LC, LL_LC, RL_LC, CRL_LC, WL_LC, "WL", _
    1.2, _
    1.2, 1.05, _
    1.2, 1.05, _
    1.2, 1.05, _
    0.6)
' ← factors

Call GenerateLateralCombos(staad, _
    nDL, nLL, nRL, nCRL, nEQ, _
    DL_LC, LL_LC, RL_LC, CRL_LC, EQ_LC, "EL", _
    1.2, _
    1.2, 1.05, _
    1.2, 1.05, _
    1.2, 1.05, _
    0.6)

        ''==========================================================================
    '' C3 : 1.2 DL  +  1.2 leading + 0.53 accomp + 1.2 WL/EL
    ''      All DL cases combined with all LL cases  (one combined combination)
    ''==========================================================================

Call GenerateLateralCombos(staad, _
    nDL, nLL, nRL, nCRL, nWL, _
    DL_LC, LL_LC, RL_LC, CRL_LC, WL_LC, "WL", _
    1.2, _
    1.2, 0.53, _
    1.2, 0.53, _
    1.2, 0.53, _
    1.2)
' ← factors

Call GenerateLateralCombos(staad, _
    nDL, nLL, nRL, nCRL, nEQ, _
    DL_LC, LL_LC, RL_LC, CRL_LC, EQ_LC, "EL", _
    1.2, _
    1.2, 0.53, _
    1.2, 0.53, _
    1.2, 0.53, _
    1.2)


    '==========================================================================
    ' C4 : 1.5 DL  +  1.5 WL   (one combination per wind direction)
    '==========================================================================
    For iWL = 0 To nWL - 1
        newComb = NextComb(staad)
        CombName = "1.5 DL + 1.5 WL" & iWL+1 & ""
        If nDL > 1 Then
            CombName = "1.5 DL + 1.5 CL + 1.5 WL" & iWL+1 & ""
        End If
        staad.Load.CreateNewLoadCombination CombName, newComb
        For iDL = 0 To nDL - 1
            staad.Load.AddLoadAndFactorToCombination newComb, DL_LC(iDL), 1.5
        Next iDL
        staad.Load.AddLoadAndFactorToCombination newComb, WL_LC(iWL), 1.5
    Next iWL

    '==========================================================================
    ' C5 : 0.9 DL  +  1.5 WL   (one combination per wind direction)
    '==========================================================================
    For iWL = 0 To nWL - 1
        newComb = NextComb(staad)
        CombName = "0.9 DL + 1.5 WL" & iWL+1 & ""
        If nDL >1 Then
            CombName = "0.9 DL + 0.9 CL + 1.5 WL" & iWL+1 & ""
        End If
        staad.Load.CreateNewLoadCombination CombName, newComb
        For iDL = 0 To nDL - 1
            staad.Load.AddLoadAndFactorToCombination newComb, DL_LC(iDL), 0.9
        Next iDL
        staad.Load.AddLoadAndFactorToCombination newComb, WL_LC(iWL), 1.5
    Next iWL

        '==========================================================================
    ' C6 : 1.5 DL  +  1.5 EL   (one combination per wind direction)
    '==========================================================================
    For iEQ = 0 To nEQ - 1
        newComb = NextComb(staad)
        CombName = "1.5 DL + 1.5 EL" & iEQ+1 & ""
        If nDL >1 Then
            CombName = "1.5 DL + 1.5 CL + 1.5 EL" & iEQ+1 & ""
        End If
        staad.Load.CreateNewLoadCombination CombName, newComb
        For iDL = 0 To nDL - 1
            staad.Load.AddLoadAndFactorToCombination newComb, DL_LC(iDL), 1.5
        Next iDL
        staad.Load.AddLoadAndFactorToCombination newComb, EQ_LC(iEQ), 1.5
    Next iEQ

    '==========================================================================
    ' C7 : 0.9 DL  +  1.5 EL   (one combination per wind direction)
    '==========================================================================
    For iEQ = 0 To nEQ - 1
        newComb = NextComb(staad)
        CombName = "0.9 DL + 1.5 EL" & iEQ+1 & ""
        If nDL >1 Then
            CombName = "0.9 DL + 0.9 CL+ 1.5 EL" & iEQ+1 & ""
        End If
        staad.Load.CreateNewLoadCombination CombName, newComb
        For iDL = 0 To nDL - 1
            staad.Load.AddLoadAndFactorToCombination newComb, DL_LC(iDL), 0.9
        Next iDL
        staad.Load.AddLoadAndFactorToCombination newComb, EQ_LC(iEQ), 1.5
    Next iEQ


    '''''''''''''''''''''''''''''''''servicability'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim nEndULS As Long ' captures the last number of strength combo
    nEndULS = gNextCombNum - 1
    gNextCombNum = nStartSLS   ' <-- reset counter for SLS loads

 '==========================================================================
    ' C8 : 1 DL  +  1 LL
    '      All DL cases combined with all LL cases  (one combined combination)
    '==========================================================================
    CombName = "1DL"
    If nDL >1 Then
            CombName = "1DL + 1CL"
    End If

    If nLL > 0 Then
        CombName = CombName & " + 1LL"
    End If

    If nRL > 0 Then
        CombName = CombName & " + 1RL"
    End If

    If nCRL > 0 Then
    For iCRL = 0 To nCRL - 1

        newComb = NextComb(staad)
        staad.Load.CreateNewLoadCombination CombName & " + 1CRL" & iCRL + 1 & "", newComb

        'DL
        For iDL = 0 To nDL - 1
            staad.Load.AddLoadAndFactorToCombination newComb, DL_LC(iDL), 1
        Next iDL

        'LL
        For iLL = 0 To nLL - 1
            staad.Load.AddLoadAndFactorToCombination newComb, LL_LC(iLL), 1
        Next iLL

        'RL
        For iRL = 0 To nRL - 1
            staad.Load.AddLoadAndFactorToCombination newComb, RL_LC(iRL), 1
        Next iRL

        'Single crane load
        staad.Load.AddLoadAndFactorToCombination newComb, CRL_LC(iCRL), 1

    Next iCRL
Else
        newComb = NextComb(staad)
    staad.Load.CreateNewLoadCombination CombName, newComb

    '-------------------------------------------------
    ' Dead loads
    '-------------------------------------------------
    For iDL = 0 To nDL - 1
        staad.Load.AddLoadAndFactorToCombination newComb, DL_LC(iDL), 1
    Next iDL

    '-------------------------------------------------
    ' Live load
    '-------------------------------------------------
    For iLL = 0 To nLL - 1
            staad.Load.AddLoadAndFactorToCombination newComb, LL_LC(iLL), 1
    Next iLL

    '-------------------------------------------------
    ' Roof live
    '-------------------------------------------------
    For iRL = 0 To nRL - 1
            staad.Load.AddLoadAndFactorToCombination newComb, RL_LC(iRL), 1
    Next iRL
End If


        '==========================================================================
    ' C9 : 1 DL  + 0.8 LL + 0.8 WL
    '      All DL cases combined with all WL cases  (one combined combination)
    '==========================================================================
Call GenerateLateralCombosSLS(staad,  _
    nDL, nLL, nRL, nCRL, nWL, _
    DL_LC, LL_LC, RL_LC, CRL_LC, WL_LC, "WL")

Call GenerateLateralCombosSLS(staad, _
    nDL, nLL, nRL, nCRL, nEQ, _
    DL_LC, LL_LC, RL_LC, CRL_LC, EQ_LC, "EL")

        '==========================================================================
    ' C10 : 1 DL  +  1 WL   (one combination per wind direction)
    '==========================================================================
    For iWL = 0 To nWL - 1
        newComb = NextComb(staad)
        CombName = "1 DL + 1 WL" & iWL+1 & ""
        If nDL >1 Then
            CombName = "1 DL + 1CL + 1 WL" & iWL+1 & ""
        End If
        staad.Load.CreateNewLoadCombination CombName, newComb
        For iDL = 0 To nDL - 1
            staad.Load.AddLoadAndFactorToCombination newComb, DL_LC(iDL), 1
        Next iDL
        staad.Load.AddLoadAndFactorToCombination newComb, WL_LC(iWL), 1
    Next iWL

    '==========================================================================
    ' Done
    '==========================================================================
    MsgBox "Load combination(s) generated successfully." & Chr(13) & _
   "Strength    : " & nStart & " to " & nEndULS & Chr(13) & _
   "Serviceability : " & nStartSLS & " to " & (gNextCombNum - 1) & Chr(13) & Chr(13) & _
   "Please verify the combinations in your STAAD.Pro model.", vbOkOnly

End Sub

' case with LL + WL/EL in strength
Sub GenerateLateralCombos(staad As Object, _
                           nDL As Integer, nLL As Integer, nRL As Integer, nCRL As Integer, _
                           nLatLC As Integer, _
                           DL_LC() As Long, LL_LC() As Long, RL_LC() As Long, CRL_LC() As Long, _
                           LatLC() As Long, sLatTag As String, _
                           fDL As Double, _          ' DL factor (always 1.2 typically)
                           fLL_lead As Double, _     ' LL factor when LL is leading  (1.2)
                           fLL_acc As Double, _      ' LL factor when LL is not leading (1.05)
                           fRL_lead As Double, _     ' RL factor when RL is leading
                           fRL_acc As Double, _      ' RL factor when RL is not leading
                           fCRL_lead As Double, _    ' CRL factor when CRL is leading
                           fCRL_acc As Double, _     ' CRL factor when CRL is not leading
                           fLat As Double)           ' Lateral (WL/EL) factor

    Dim iLead As Integer, iWL As Integer, iCRL As Integer
    Dim iDL As Integer, iLL As Integer, iRL As Integer
    Dim CombName As String
    Dim newComb As Long

    For iLead = 1 To 3

        If iLead = 1 And nLL  = 0 Then GoTo SkipLead
        If iLead = 2 And nRL  = 0 Then GoTo SkipLead
        If iLead = 3 And nCRL = 0 Then GoTo SkipLead

        'nCurr = nCurr + 1

        ' ── Build base name ───────────────────────────────────────────
        CombName = fDL & "DL"

        If nDL >1 Then
            CombName = fDL & "DL +" & fDL & "CL"
        End If

        If nLL > 0 Then
            CombName = CombName & IIf(iLead = 1, " + " & fLL_lead & "LL", " + " & fLL_acc & "LL")
        End If
        If nRL > 0 Then
            CombName = CombName & IIf(iLead = 2, " + " & fRL_lead & "RL", " + " & fRL_acc & "RL")
        End If

        ' ── Loop over lateral cases ───────────────────────────────────
        If nCRL > 0 Then
            For iCRL = 0 To nCRL - 1
                For iWL = 0 To nLatLC - 1
                     newComb = NextComb(staad)
                    Dim crlTag As String
                    crlTag = IIf(iLead = 3, " + " & fCRL_lead & "CRL", " + " & fCRL_acc & "CRL")
                    staad.Load.CreateNewLoadCombination _
                        CombName & crlTag & (iCRL + 1) & " + " & fLat & sLatTag & (iWL + 1), newComb

                    For iDL = 0 To nDL - 1
                        staad.Load.AddLoadAndFactorToCombination newComb, DL_LC(iDL), fDL
                    Next iDL
                    For iLL = 0 To nLL - 1
                        staad.Load.AddLoadAndFactorToCombination newComb, LL_LC(iLL), _
                            IIf(iLead = 1, fLL_lead, fLL_acc)
                    Next iLL
                    For iRL = 0 To nRL - 1
                        staad.Load.AddLoadAndFactorToCombination newComb, RL_LC(iRL), _
                            IIf(iLead = 2, fRL_lead, fRL_acc)
                    Next iRL
                    staad.Load.AddLoadAndFactorToCombination newComb, CRL_LC(iCRL), _
                        IIf(iLead = 3, fCRL_lead, fCRL_acc)
                    staad.Load.AddLoadAndFactorToCombination newComb, LatLC(iWL), fLat
                Next iWL
            Next iCRL
        Else
            For iWL = 0 To nLatLC - 1
                newComb = NextComb(staad)
                staad.Load.CreateNewLoadCombination _
                    CombName & " + " & fLat & sLatTag & (iWL + 1), newComb

                For iDL = 0 To nDL - 1
                    staad.Load.AddLoadAndFactorToCombination newComb, DL_LC(iDL), fDL
                Next iDL
                For iLL = 0 To nLL - 1
                    staad.Load.AddLoadAndFactorToCombination newComb, LL_LC(iLL), _
                        IIf(iLead = 1, fLL_lead, fLL_acc)
                Next iLL
                For iRL = 0 To nRL - 1
                    staad.Load.AddLoadAndFactorToCombination newComb, RL_LC(iRL), _
                        IIf(iLead = 2, fRL_lead, fRL_acc)
                Next iRL
                staad.Load.AddLoadAndFactorToCombination newComb, LatLC(iWL), fLat
            Next iWL
        End If

SkipLead:
    Next iLead

End Sub

' combo with 0.8WL/EL in serviciability
Sub GenerateLateralCombosSLS(staad As Object, _
                           nDL As Integer, nLL As Integer, nRL As Integer, nCRL As Integer, _
                           nLatLC As Integer, _
                           DL_LC() As Long, LL_LC() As Long, RL_LC() As Long, CRL_LC() As Long, _
                           LatLC() As Long, sLatTag As String)

    Dim iWL As Integer, iCRL As Integer
    Dim iDL As Integer, iLL As Integer, iRL As Integer
    Dim CombName As String
    Dim newComb As Long

    CombName = "1DL"
     If nDL >1 Then
        CombName = "1DL + 1CL"
    End If

    If nLL > 0 Then
        CombName = CombName & " + 0.8LL"
    End If

    If nRL > 0 Then
        CombName = CombName & " + 0.8RL"
    End If

    '-------------------------------------------------
    ' IF Crane load
    '-------------------------------------------------
    If nCRL > 0 Then
    For iCRL = 0 To nCRL - 1
    For iWL = 0 To nLatLC - 1

        newComb = NextComb(staad)
        staad.Load.CreateNewLoadCombination CombName & " + 0.8CRL" & iCRL + 1 & " + 0.8" & sLatTag & iWL+1, newComb


        'DL
        For iDL = 0 To nDL - 1
            staad.Load.AddLoadAndFactorToCombination newComb, DL_LC(iDL), 1
        Next iDL

        'LL
        For iLL = 0 To nLL - 1
                staad.Load.AddLoadAndFactorToCombination newComb, LL_LC(iLL), 0.8
        Next iLL

        'RL
        For iRL = 0 To nRL - 1
                staad.Load.AddLoadAndFactorToCombination newComb, RL_LC(iRL), 0.8
        Next iRL

        'Single crane load
            staad.Load.AddLoadAndFactorToCombination newComb, CRL_LC(iCRL), 0.8

        'Wind Load
        staad.Load.AddLoadAndFactorToCombination newComb, LatLC(iWL), 0.8
    Next iWL
    Next iCRL
Else
    For iWL = 0 To nLatLC - 1
        newComb = NextComb(staad)
        staad.Load.CreateNewLoadCombination CombName & " + 0.8" & sLatTag & iWL+1, newComb
    
        '-------------------------------------------------
        ' Dead loads
        '-------------------------------------------------
        For iDL = 0 To nDL - 1
            staad.Load.AddLoadAndFactorToCombination newComb, DL_LC(iDL), 1
        Next iDL
    
        '-------------------------------------------------
        ' Live load
        '-------------------------------------------------
        For iLL = 0 To nLL - 1
                staad.Load.AddLoadAndFactorToCombination newComb, LL_LC(iLL), 0.8
        Next iLL
    
        '-------------------------------------------------
        ' Roof live
        '-------------------------------------------------
        For iRL = 0 To nRL - 1
                staad.Load.AddLoadAndFactorToCombination newComb, RL_LC(iRL), 0.8
        Next iRL
        staad.Load.AddLoadAndFactorToCombination newComb, LatLC(iWL), 0.8
    Next iWL
End If
End Sub

Function NextComb(staad As Object) As Long
    NextComb = gNextCombNum
    gNextCombNum = gNextCombNum + 1
End Function
