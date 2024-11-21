'[ Dimension variables
Dim totalLength As Double = 72.0
Dim totalWidth As Double = 48.0
Dim feetHeight As Double = 3.0
Dim numShelf As Integer = 3
Dim returnShelf As Integer = 0
Dim desiredShelfGap As Double = 15.0
Dim rollerPlaneAngle As Double = 5.0
Dim desiredRollerGap As Double = 5.0
']

'[ Dimension parameters
Parameter("RACK_BASE_TUBE_Z:1", "Length") = totalLength
Parameter("RACK_BASE_TUBE_X:1", "Length") = totalWidth
Parameter("RACK_LEG_TUBE:1", "feetHeight") = feetHeight
Parameter("RACK_SHELF:1", "desiredGap") = desiredRollerGap
Parameter("RACK_SHELF:1", "totalWidth") = totalWidth
Parameter("RACK_SHELF:1", "rollerPlaneAngle") = rollerPlaneAngle
Parameter("RACK_SHELF:1", "sideSupportOffset") = desiredShelfGap / 2

Parameter(MakePath("RACK_SHELF:1", "RACK_SHELF_TUBE_X:1"), "Length") = totalWidth
Parameter(MakePath("RACK_SHELF:1", "RACK_SHELF_TUBE_X:1"), "rollerPlaneAngle") = rollerPlaneAngle
Parameter(MakePath("RACK_SHELF:1", "tempRoller:1"), "rollerPlaneAngle") = rollerPlaneAngle
Parameter(MakePath("RACK_SHELF:1", "tempRoller:1"), "totalLength") = totalLength
Parameter("RACK_SHELF:1", "rollerLength") = Parameter("tempRoller:1", "Length")
Dim rollerLength As Double = Parameter("tempRoller:1", "Length")
Dim calcShelfGap As Double = ((rollerLength - 2 * (Parameter("tempRoller:1", "Thickness")) - Parameter("RACK_SHELF_TUBE_X:1", "Diameter")) * Sin(rollerPlaneAngle * (Math.PI / 180)))

Dim shelfHeight As Double = Parameter("RACK_SHELF:1", "shelfHeight")
Dim sideSupportOffset As Double = Parameter("RACK_SHELF:1", "sideSupportOffset")
Dim totalHeight As Double
If returnShelf <> numShelf And returnShelf > 1 Then
	totalHeight = ((((numShelf * (shelfHeight)) + ((numShelf - 1) * desiredShelfGap)) + feetHeight) + 2.25) -((numShelf - 3) * calcShelfGap) + sideSupportOffset
ElseIf returnShelf = 0 Then
	totalHeight = ((((numShelf * (shelfHeight)) + ((numShelf - 1) * desiredShelfGap)) + feetHeight) + 2.25) - ((numShelf - 1) * calcShelfGap) + sideSupportOffset
Else
	totalHeight = ((((numShelf * (shelfHeight)) + ((numShelf - 1) * desiredShelfGap)) + feetHeight) + 2.25) -((numShelf - 2) * calcShelfGap) + sideSupportOffset
End If
Parameter("RACK_LEG_TUBE:1", "Length") = totalHeight
InventorVb.DocumentUpdate()
']

Dim asmDoc As AssemblyDocument = ThisApplication.ActiveDocument
Dim asmCompDefRack As AssemblyComponentDefinition = asmDoc.ComponentDefinition

' Identify the "RACK_BASE" sub-assembly occurrence
Dim subAsmOccurrence As ComponentOccurrence = asmCompDefRack.Occurrences.ItemByName("RACK_BASE:1")

' Access the "RACK_BASE" sub-assembly document
Dim subAsmDoc As AssemblyDocument = subAsmOccurrence.Definition.Document
Dim subAsmCompDef As AssemblyComponentDefinition = subAsmDoc.ComponentDefinition

'[ Delete old "RACK_SHELF:" occurrences
For i As Integer = asmCompDefRack.Occurrences.Count To 1 Step -1
	Dim occurrence As ComponentOccurrence = asmCompDefRack.Occurrences.Item(i)
	If occurrence.Name.StartsWith("RACK_SHELF:") Then
		occurrence.Delete()
	End If
Next
']

' Access the part occurrence within the sub-assembly
Dim partOccurrence As ComponentOccurrence = subAsmCompDef.Occurrences.ItemByName("RACK_LEG_TUBE:3")

' Access the part document
Dim partDoc As PartDocument = partOccurrence.Definition.Document
Dim partCompDef As PartComponentDefinition = partDoc.ComponentDefinition

' Create offset planes in the part document
Dim partPlanes As WorkPlanes = partCompDef.WorkPlanes

'[ Delete old offset planes
For j As Integer = partPlanes.Count To 1 Step -1
	Dim wp As WorkPlane
	pp = partPlanes.Item(j)
	
	If pp.Name.StartsWith("tempShelfPlane:") Then
		pp.Delete()
	End If
Next
']

Dim currentHeight As Double = feetHeight + 3

'[ Main loop
For k As Integer = 1 To numShelf
	'[ Create new offset planes
	Dim offsetDistance As Double = currentHeight * 2.54
	Dim offsetPlane As WorkPlane = partPlanes.AddByPlaneAndOffset(partCompDef.WorkPlanes.Item("XZ Plane"), offsetDistance)
	offsetPlane.Visible = True

	If returnShelf = k Then
		currentHeight += shelfHeight + desiredShelfGap
	ElseIf returnShelf = k + 1 Then
		currentHeight += shelfHeight + desiredShelfGap
	Else
		currentHeight += shelfHeight + desiredShelfGap - calcShelfGap
	End If

	' Name the offset plane
	offsetPlane.Name = "legTubeOffsetShelf:" & k
	']
	
	'[ Add Shelves
	' Create an identity matrix
	Dim identityMatrix As Matrix = ThisApplication.TransientGeometry.CreateMatrix()
	
	' Place the "RACK_SHELF" assembly into the current assembly
	Dim shelfOccurrence As ComponentOccurrence = Nothing
	shelfOccurrence = asmCompDefRack.Occurrences.Add("C:\Users\240126022\Documents\Inventor\Flow Rack\RACK_SHELF.iam", identityMatrix)
	']
	
	'[ Add Constraints & Iterators
	Dim legTubeOffsetName As String = "legTubeOffsetShelf:" & k
	Dim shelfName As String = "RACK_SHELF:" & k
	Dim flushCount As String = "Flush:" & k
	Dim mateCount_XY As String = "Mate:" & (2 * k - 1)
	Dim mateCount_YZ As String = "Mate:" & (2 * k)
	
	If returnShelf = k Then
		' Add constraints for return shelf
		Constraints.AddFlush(flushCount, {shelfName, "T-SHAPE CONNECTOR:2" }, "Work Plane2", {"RACK_BASE:1", "RACK_LEG_TUBE:3" }, legTubeOffsetName)
		Constraints.AddMate(mateCount_XY, {shelfName, "T-SHAPE CONNECTOR:2" }, "XY Plane", {"RACK_BASE:1", "RACK_LEG_TUBE:3" }, "XY Plane")
		Constraints.AddMate(mateCount_YZ, {shelfName, "T-SHAPE CONNECTOR:2" }, "YZ Plane", {"RACK_BASE:1", "RACK_LEG_TUBE:3" }, "YZ Plane")
	Else
		' Add constraints for flow shelf
		Constraints.AddFlush(flushCount, {shelfName, "T-SHAPE CONNECTOR:6" }, "Work Plane2", {"RACK_BASE:1", "RACK_LEG_TUBE:2" }, legTubeOffsetName)
		Constraints.AddMate(mateCount_XY, {shelfName, "T-SHAPE CONNECTOR:8" }, "XY Plane", {"RACK_BASE:1", "RACK_LEG_TUBE:3" }, "XY Plane")
		Constraints.AddMate(mateCount_YZ, {shelfName, "T-SHAPE CONNECTOR:8" }, "YZ Plane", {"RACK_BASE:1", "RACK_LEG_TUBE:3" }, "YZ Plane")
	End If
	']
	
	' Rename the offset plane
    offsetPlane.Name = "tempShelfPlane:" & k
Next
']

RuleParametersOutput()
iLogicVb.UpdateWhenDone = True