Dim oAsm As AssemblyDocument = ThisDoc.Document
Dim oSelSet As SelectSet = oAsm.SelectSet
Dim UoM As UnitsOfMeasure = oAsm.UnitsOfMeasure
Dim totMass As Double = 0
Dim oList As New List(Of ComponentOccurrence)
If ThisDoc.PathAndFileName(False) = "" Then
    MessageBox.Show("Please save the assembly first.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    Exit Sub
Else
		oWrite = System.IO.File.CreateText(ThisDoc.PathAndFileName(False) & ".txt")
    For Each oObj As Object In oSelSet
        If TypeOf (oObj) Is ComponentOccurrence Then oList.Add(oObj)
    Next
    Dim noSelection As Boolean = False
    If oList.Count = 0
        noSelection = True
        For Each oOcc As ComponentOccurrence In oAsm.ComponentDefinition.Occurrences.AllLeafOccurrences
            If oOcc.Visible Then 
                totMass = totMass + oOcc.MassProperties.Mass
                ' Show individual mass of each component
                oWrite.WriteLine("Mass of component " & oOcc.Name & ": " & UoM.GetStringFromValue(oOcc.MassProperties.Mass, UoM.MassUnits))
            End If
        Next
    Else
        For Each oOcc As ComponentOccurrence In oList
            Dim parentSelected As Boolean = False
            'Make sure something isn't selected twice (subassembly and occurrence in subassembly both selected)
            For Each oParentOcc As ComponentOccurrence In oOcc.OccurrencePath
                If oList.OfType(Of ComponentOccurrence).Where(Function(x As ComponentOccurrence) x.Definition Is oParentOcc.Definition _
                AndAlso x.Definition IsNot oOcc.Definition).Count > 0 Then parentSelected = True
            Next
            '--------------------------------------------------------------------------------------------------
            If parentSelected = False Then 
                totMass = totMass + oOcc.MassProperties.Mass
                ' Show individual mass of each component
                oWrite.WriteLine("Mass of component " & oOcc.Name & ": " & UoM.GetStringFromValue(oOcc.MassProperties.Mass, UoM.MassUnits))
            End If
        Next
    End If
    ' Show total mass
    oWrite.WriteLine(If (noSelection, "Total mass of visible components", "Total mass of selected components") _
    & vbCrLf & UoM.GetStringFromValue(totMass, UoM.MassUnits), _
    "Total mass")
    oWrite.Close()
End If
