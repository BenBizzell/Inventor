Try
	'[ Dimension variables
	Dim availableSpace As Double = totalWidth - 1.5
	Dim numRoller As Integer = Math.Floor(availableSpace / desiredGap)
	Dim rollerOffset As Double = availableSpace / numRoller
	']
	
	' Access the active document
	Dim asmDoc As AssemblyDocument = ThisApplication.ActiveDocument
	Dim asmCompDefShelf As AssemblyComponentDefinition = asmDoc.ComponentDefinition
	
	'[ Delete old "TEMP_ROLLER" occurrences
	For i As Integer = asmCompDefShelf.Occurrences.Count To 1 Step -1
	 	Dim occurrence As ComponentOccurrence = asmCompDefShelf.Occurrences.Item(i)
		If occurrence.Name.StartsWith("tempRoller:") Then
	 		occurrence.Delete()
	 	End If
	Next
	']
	
	' Access the part occurrence
	Dim partOccurrenceRoller As ComponentOccurrence = asmCompDefShelf.Occurrences.ItemByName("RACK_SHELF_TUBE_X:3")
	
	' Access the part document
	Dim partDoc As PartDocument = partOccurrenceRoller.Definition.Document
	Dim partCompDef As PartComponentDefinition = partDoc.ComponentDefinition
	
	' Create offset planes in the part document
	Dim partPlanes As WorkPlanes = partCompDef.WorkPlanes
	
	'[ Delete old offset planes
	For j As Integer = partPlanes.Count To 1 Step -1
	    pp = partPlanes.Item(j) 
	    
	    If pp.Name.StartsWith("tempRollerPlane:") Then
	        pp.Delete()
	    End If
	Next
	']
	
	'[ Constraint iterators
	Dim mateCountStart As Integer = 23
	Dim flushCountStart As Integer = 20
	']
	
	'[ Main loop
	For k As Integer = 1 To numRoller
	    '[ Create new offset planes
	    Dim offsetDistance As Double = (k * rollerOffset) * 2.54 ' Assuming conversion to cm
		
		If offsetDistance = availableSpace * 2.54 Then
			Exit For
		End If
		
	    Dim offsetPlane As WorkPlane = partPlanes.AddByPlaneAndOffset(partCompDef.WorkPlanes.Item("endConnector"), offsetDistance)
	    offsetPlane.Visible = True
	    
	    ' Name the offset plane
	    offsetPlane.Name = "rollerOffsetPlane:" & k
		']
	    
		'[ Add Rollers
	    ' Create an identity matrix
	    Dim identityMatrix As Matrix = ThisApplication.TransientGeometry.CreateMatrix()
	    
	    ' Place the "RACK_ROLLER" part into the current assembly
	    Dim rollerOccurrence As ComponentOccurrence = Nothing
	    rollerOccurrence = asmCompDefShelf.Occurrences.Add("C:\Users\240126022\Documents\Inventor\Flow Rack\RACK_ROLLER.ipt", identityMatrix)
		rollerOccurrence.Name = "rackRoller:" & k
		']
	    
		'[ Add Constraints & Iterators
	    Dim rollerOffsetName As String = "rollerOffsetPlane:" & k
	    Dim rollerName As String = "rackRoller:" & k
	    Dim mateCountBot As String = "Mate:" & mateCountStart
		Dim mateCountTop As String = "Mate:" & (mateCountStart + 1)
	    Dim flushCountXZ As String = "Flush:" & flushCountStart
	    Dim flushCountYZ As String = "Flush:" & (flushCountStart + 1)
	    
		Try
			' Add constraints using the provided syntax
		    Constraints.AddMate(mateCountBot, rollerName, "botRollerPlane", "RACK_SHELF_TUBE_X:3", "botRollerPlane")
			Constraints.AddFlush(flushCountXZ, rollerName, "angledPlane", "RACK_SHELF_TUBE_X:3", "rollerPlane")
			Constraints.AddFlush(flushCountYZ, rollerName, "YZ Plane", "RACK_SHELF_TUBE_X:3", rollerOffsetName)
		Catch ex As Exception
		End Try
		
		Try
			Constraints.AddMate(mateCountTop, rollerName, "topRollerPlane", "RACK_SHELF_TUBE_X:2", "topRollerPlane")	
		Catch ex As Exception
		End Try
	
		mateCountStart += 2
		flushCountStart += 2
		']
	
	    ' Rename the offset plane
	    offsetPlane.Name = "tempRollerPlane:" & k
		
		rollerOccurrence.Name = "tempRoller:" & k
	Next
	']
	
	RuleParametersOutput()
	iLogicVb.UpdateWhenDone = True
Catch ex As Exception
End Try