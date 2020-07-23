Public Sub ExtrudeMultipleBodies()

Dim StillSelecting As Boolean
StillSelecting = True

Dim partDoc As PartDocument
Set partDoc = ThisApplication.ActiveEditDocument

Dim compDef As PartComponentDefinition
Set compDef = partDoc.ComponentDefinition


Dim regionSelections As ObjectCollection
Set regionSelections = ThisApplication.TransientObjects.CreateObjectCollection()



While StillSelecting = True

Call regionSelections.Add(ThisApplication.CommandManager.Pick(kSketchProfileFilter, "Select a profile to extrude"))

 Dim keepSelecting As String
 keepSelecting = MsgBox("Do you want to select more profiles?", vbYesNo)
 
 '''selecting Yes outputs an integer of 6, No an integer of 7
 If keepSelecting = 7 Then
 StillSelecting = False
 
 End If
 
 Wend
 
Dim extrudeToPoint As SketchPoint
Dim extrudeToFace As Face

Dim extrudeTo As Object
Dim selection As SelectionFilterEnum

Set extrudeTo = ThisApplication.CommandManager.Pick(kAllEntitiesFilter, "Select the point to extrude to or face to extrude from")

Dim extrudeDef As ExtrudeDefinition

Dim extFeature As ExtrudeFeature

If TypeOf extrudeTo Is SketchPoint Then
 
 Set extrudeToPoint = extrudeTo
 
 For Each skProfile In regionSelections
 
 Set extrudeDef = compDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(skProfile, kNewBodyOperation)
 
 Call extrudeDef.SetToExtent(extrudeToPoint)
 
 Set extFeature = compDef.Features.ExtrudeFeatures.Add(extrudeDef)
 
 Next
 
ElseIf TypeOf extrudeTo Is Face Then

 Dim fromFace As Face
 Dim toFace As Face
 
 Set fromFace = extrudeTo
 Set toFace = ThisApplication.CommandManager.Pick(kPartFaceFilter, "Select the face to extrude to")
 
 For Each skProfile In regionSelections
 
 Set extrudeDef = compDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(skProfile, kNewBodyOperation)
 
 Call extrudeDef.SetFromToExtent(fromFace, True, toFace, True)
 
 Set extFeature = compDef.Features.ExtrudeFeatures.Add(extrudeDef)
 
 Next
 
 
 End If

End Sub