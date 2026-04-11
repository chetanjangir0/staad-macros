Sub Main()

Dim objOpenSTAAD

Set objOpenSTAAD = GetObject(, "StaadPro.OpenSTAAD")

'Run analysis
objOpenSTAAD.Analyze.RunAnalysis()

'Run steel design
'objOpenSTAAD.Design.RunSteelDesign

'Switch to post processing mode
objOpenSTAAD.Command "POSTPROCESS"

MsgBox "Analysis and Design Completed"

End Sub
