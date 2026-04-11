Sub Main()
'DESCRIPTION:Create a 2D frame with supports

    Begin Dialog UserDialog 600,200,"2D Frame Model" ' %GRID:5,5,1,1
        Text 20,20,190,15,"No. of Horizontal Bays:",.Text1
        Text 20,45,190,15,"No. of Vertical Bays",.Text2
        Text 20,70,190,15,"Vertical Distance",.Text3
        Text 20,95,190,15,"Horizontal Distance",.Text4
        Text 20,130,190,15,"Support Type",.Text5
        TextBox 220,20,130,15,.clmn
        TextBox 220,40,130,15,.row
        TextBox 220,70,130,15,.ht
        TextBox 220,95,130,15,.wdth
        OptionGroup .sprt
            OptionButton 220,130,90,15,"Fixed",.OptionButton1
            OptionButton 370,130,90,15,"Pinned",.OptionButton2
        OKButton 360,165,90,20
        CancelButton 490,165,90,20
    End Dialog
    Dim dlg As UserDialog

    Dim dlgResult As Integer
    Dim crdx As Double
    Dim crdy As Double
    Dim crdz As Double
    Dim n1 As Long
    Dim n2 As Long
    Dim i1 As Long
    Dim s1 As Long

    'Initialization
    dlg.clmn = "3"
    dlg.row = "5"
    dlg.ht = "3"
    dlg.wdth = "5"

    'Popup the dialog
    dlgResult = Dialog(dlg)
    Debug.Clear

    If dlgResult = -1 Then 'OK button pressed
        Debug.Print "OK button pressed"

        clmn = Abs( CDbl(dlg.clmn) )
        row = Abs( CDbl(dlg.row) )
        ht = Abs( CDbl(dlg.ht) )
        wdth = Abs( CDbl(dlg.wdth) )
        sprt = CStr(dlg.sprt)

        Debug.Print "No. of Horizontal Bays = ";clmn
        Debug.Print "No. of Vertical Bays = ";row
        Debug.Print "Vertical Distance = ";ht
        Debug.Print "Horizontal Distance = ";wdth
        Debug.Print "Support Type = ";sprt

        crdx = 0
        crdy = 0
        crdz = 0

        Dim objOpenSTAAD As Object
        Set objOpenSTAAD = GetObject(,"StaadPro.OpenSTAAD")

        Dim geometry As Object
        Set geometry = objOpenSTAAD.Geometry

        'Nodes
        For j = 2 To (row + 2)
            For i = 1 To (clmn + 1)
                crdx = (i - 1) * wdth
                geometry.AddNode crdx, crdy, crdz
            Next
            crdy = (j - 1) * ht
        Next

        Dim support As Object
        Set support = objOpenSTAAD.Support

        'Supports
        If sprt = "0" Then
            s1 = support.CreateSupportFixed()
        ElseIf sprt = "1" Then
            s1 = support.CreateSupportPinned()
        Else
            MsgBox("Select Proper Support Type",vbOkOnly,"Error")
            Exit Sub
        End If
        Debug.Print "Support return value = ";s1
        For i1 = 1 To (clmn + 1)
            support.AssignSupportToNode i1,s1
        Next

        'Columns
        n1 = 1
        n2 = (n1 + clmn +1)
        For k = 1 To (clmn + 1)*row
            geometry.AddBeam n1, n2
            n1 = n1 + 1
            n2 = n2 + 1
        Next

        'Beams
        n1 = 1
        For k1 = 1 To row
            n1 = k1 * (clmn + 1)+1
            n2 = n1 + 1
            For k2 = 1 To clmn
                geometry.AddBeam n1, n2
                n1 = n1 + 1
                n2 = n2 + 1
            Next
        Next

    ElseIf dlgResult = 0 Then 'Cancel button pressed
        Debug.Print "Cancel button pressed"
    End If

End Sub
