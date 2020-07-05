'You can insert Content Center parts programmatically. The following sample is looking through the whole Content Center content in 'order to find a specific part based on its display name.
'https://adndevblog.typepad.com/manufacturing/2013/06/insert-content-center-part.html

'Example use:
'CcTest

Sub Main()
    PlaceContentCenterPart()

End Sub

Public Sub PlaceContentCenterPart(partName As String)
    Dim asm As AssemblyDocument = ThisApplication.ActiveDocument

    Dim cf As ContentFamily = GetFamily(partName, Nothing)

    Dim ee As MemberManagerErrorsEnum
    'insert the first row in the family
    Dim member As String = cf.CreateMember(1, ee, "Problem")

    Dim tg As TransientGeometry = ThisApplication.TransientGeometry

    Call asm.ComponentDefinition.Occurrences.Add(member, tg.CreateMatrix())
End Sub


Public Function GetFamily(name As String, node As ContentTreeViewNode) As ContentFamily
  Dim cc As ContentCenter = ThisApplication.ContentCenter
    
  If node Is Nothing Then node = cc.TreeViewTopNode
  
  Dim cf As ContentFamily
  For Each cf In node.Families
        If cf.DisplayName = name Then
            GetFamily = cf
            Exit Function
        End If
    Next
  
  Dim child As ContentTreeViewNode
    For Each child In node.ChildNodes
        cf = GetFamily(name, child)
        If Not cf Is Nothing Then
            GetFamily = cf
            Exit Function
        End If
    Next

    Throw New System.ArgumentException("The content center part could not be found.  Exiting.")
End Function


'The API Help File (C:\Users\Public\Documents\Autodesk\Inventor 2017\Local Help\admapi_21_0.chm) also contains a couple of samples concerning this topic:
'- Place Content Center Parts API Sample
'- Replace content center part API Sample