
'From https://forums.autodesk.com/t5/inventor-customization/turn-off-adaptively-used/m-p/6691069/highlight/true#M67989

'##################################
'###   Turn off all Adaptivity  ###
'##################################

' Note this iLogic rule will remove any adaptive links between:

' 1) All parts in this assembly
' 2) All parts in this assembly and other assemblies they may also be assembled into

'check this file is an assembly
Dim doc As Document = ThisApplication.ActiveDocument
If doc.DocumentType <> kAssemblyDocumentObject Then
MessageBox.Show("This rule can only be run in an assembly file!",
"Excitech iLogic")
Return
End If

Dim iCounter As Integer = 0

' Loop through all referenced docs in assembly
For Each oDoc As Document In doc.AllReferencedDocuments
Try
If oDoc.ModelingSettings.AdaptivelyUsedInAssembly = True Then
oDoc.ModelingSettings.AdaptivelyUsedInAssembly = False
iCounter += 1
End If
Catch
End Try
Next

MessageBox.Show("Adaptivity was switched off for " & iCounter & "
file(s)!", "Excitech iLogic")
