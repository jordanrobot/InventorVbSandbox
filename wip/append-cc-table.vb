'This is a WIP, not fully working

Sub Main()

    Dim oContentCenter As ContentCenter
    oContentCenter = ThisApplication.ContentCenter
    
    Dim hexHeadNode As ContentTreeViewNode
    hexHeadNode = ThisApplication.ContentCenter.TreeViewTopNode.ChildNodes.Item("Other Parts").ChildNodes.Item("Knobs and Handles")'.ChildNodes.Item("Phenolic Ball Knobs")
    
    
    Dim family As ContentFamily
    Dim checkFamily As ContentFamily
    For Each checkFamily In hexHeadNode.Families
        If checkFamily.DisplayName = "Phenolic Ball Knobs" Then
            family = checkFamily
            Exit For
        End If
    Next

    
     Dim oContentTableColumns As ContentTableColumns
    oContentTableColumns = family.TableColumns
    
    'dump all columns
    Dim oContentTableColumn As ContentTableColumn
    For Each oContentTableColumn In oContentTableColumns
    
     Debug.Print ("internal name:" & oContentTableColumn.InternalName & "<<<<>>>>display heading:" & oContentTableColumn.DisplayHeading)
    
    Next
    
    'add one column
    Dim oNewCol As ContentTableColumn
    oNewCol = oContentTableColumns.Add("MaterialDWG", "MaterialDWG", kStringType)
    
    'modify cell
    Dim oRow As ContentTableRow
    For Each oRow In family.TableRows
        
        'cell value of new column
        oRow.Item("MaterialDWG").Value = "=<MFGR><VENDOR> <MFGR NO><VEND NO>"
        'cell value of a built-in column named 'GD1', display name: "Pitch Diameter"
        'e.g. double it
        'use internal name of the column
        'oRow.Item("KLG").Value = oRow.Item("KLG").Value * 2
    Next
    
    'save change
    family.Save

End Sub
