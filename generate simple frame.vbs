Sub Main()
'DESCRIPTION:Create simple 2d frame

    Begin Dialog UserDialog 600,200,"2D Frame Model" ' %GRID:5,5,1,1
        Text 20,20,190,15,"Width:",.Text1
        Text 20,45,190,15,"Eave Height",.Text2
        Text 20,70,190,15,"Roof slope x (1:x):",.Text3
        Text 20,95,190,15,"Left Support Type:",.Text4
        Text 20,130,190,15,"Right Support Type",.Text5
        TextBox 220,20,130,15,.width
        TextBox 220,40,130,15,.ht
        TextBox 220,70,130,15,.slope
        OptionGroup .sprt1
            OptionButton 220,130,90,15,"Fixed",.OptionButton1
            OptionButton 370,130,90,15,"Pinned",.OptionButton2
        OptionGroup .sprt2
            OptionButton 220,140,90,15,"Fixed",.OptionButton3
            OptionButton 370,140,90,15,"Pinned",.OptionButton4
        OKButton 360,165,90,20
        CancelButton 490,165,90,20
    End Dialog
    Dim dlg As UserDialog

    Dim dlgResult As Integer
    Dim s1 As Long
    Dim s2 As Long
    Dim fw As Double  ' frame width
    Dim eh As Double  ' eave height
    Dim rs As Double  ' roof slope
    Dim sp1 As String ' support type 1
    Dim sp2 As String ' support type 2

    'Initialization
    dlg.width = "20"
    dlg.ht = "7"
    dlg.slope = "5"

    'Popup the dialog
    dlgResult = Dialog(dlg)
    Debug.Clear

    If dlgResult = -1 Then 'OK button pressed

        fw  = Abs(CDbl(dlg.width))
        eh  = Abs(CDbl(dlg.ht))
        rs  = Abs(CDbl(dlg.slope))
        sp1 = CStr(dlg.sprt1)
        sp2 = CStr(dlg.sprt2)

        Debug.Print "OK button pressed"
        Debug.Print "Frame Width     = "; fw
        Debug.Print "Eave Height     = "; eh
        Debug.Print "Roof Slope  1:  "; rs
        Debug.Print "Support Type 1  = "; sp1
        Debug.Print "Support Type 2  = "; sp2

        Dim objOpenSTAAD As Object
        Set objOpenSTAAD = GetObject(,"StaadPro.OpenSTAAD")

        ' Set units to Meter (4) and KiloNewton (5)
        objOpenSTAAD.SetInputUnits 4, 5

        Dim geometry As Object
        Set geometry = objOpenSTAAD.Geometry

        ' STEP 1: Add all nodes
        geometry.AddNode 0,      0,                  0   ' Node 1 - base left
        geometry.AddNode fw,     0,                  0   ' Node 2 - base right
        geometry.AddNode 0,      eh,                 0   ' Node 3 - eave left
        geometry.AddNode fw,     eh,                 0   ' Node 4 - eave right
        geometry.AddNode fw/2,   eh + (fw/(2*rs)),   0   ' Node 5 - ridge

        Debug.Print "Nodes added"

        ' STEP 2: Add all beams
        geometry.AddBeam 1, 3   ' Left column
        geometry.AddBeam 2, 4   ' Right column
        geometry.AddBeam 3, 5   ' Left rafter
        geometry.AddBeam 4, 5   ' Right rafter

        Debug.Print "Beams added"

        ' STEP 3: Create and assign supports
        Dim support As Object
        Set support = objOpenSTAAD.Support

        If sp1 = "0" Then
            s1 = support.CreateSupportFixed()
        ElseIf sp1 = "1" Then
            s1 = support.CreateSupportPinned()
        Else
            MsgBox "Select Proper Support Type for Left Support", vbOKOnly, "Error"
            Exit Sub
        End If

        If sp2 = "0" Then
            s2 = support.CreateSupportFixed()
        ElseIf sp2 = "1" Then
            s2 = support.CreateSupportPinned()
        Else
            MsgBox "Select Proper Support Type for Right Support", vbOKOnly, "Error"
            Exit Sub
        End If

        Debug.Print "Support 1 handle = "; s1
        Debug.Print "Support 2 handle = "; s2

        support.AssignSupportToNode 1, s1
        support.AssignSupportToNode 2, s2

        Debug.Print "Supports assigned"
        Debug.Print "Script completed successfully"

    ElseIf dlgResult = 0 Then 'Cancel button pressed
        Debug.Print "Cancel button pressed"
    End If

End Sub
