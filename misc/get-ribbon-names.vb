Public Sub PrintRibbon()
    Open "C:\RibbonNames.txt" For Output As #1

    Print #1, "File Controls (Application Menu)"
    Call PrintControls(ThisApplication.UserInterfaceManager.FileBrowserControls, "", 1)
    Print #1, "------------------------------------------------------------------"

    Print #1, "Help Controls"
    Call PrintControls(ThisApplication.UserInterfaceManager.HelpControls, "", 1)
    Print #1, "------------------------------------------------------------------"

    Dim oRibbon As Ribbon
    For Each oRibbon In ThisApplication.UserInterfaceManager.Ribbons
        Print #1, "Ribbon: " & oRibbon.InternalName

        Print #1, "    QAT controls"
        Call PrintControls(oRibbon.QuickAccessControls, "            ", 0)

        Dim oTab As RibbonTab
        For Each oTab In oRibbon.RibbonTabs
            Print #1, "    Tab: " & oTab.DisplayName & ", " & oTab.InternalName & ", Visible: " & oTab.Visible

            Dim oPanel As RibbonPanel
            For Each oPanel In oTab.RibbonPanels
                Print #1, "        Panel: " & oPanel.DisplayName & ", " & oPanel.InternalName & ", Visible: " & oPanel.Visible

                Call PrintControls(oPanel.CommandControls, "            ", 0)

                If oPanel.SlideoutControls.Count > 0 Then
                    Print #1, "            --- Slideout Controls ---"
                    Call PrintControls(oPanel.SlideoutControls, "            ", 0)
                End If
            Next
        Next

    Print #1, "------------------------------------------------------------------"
    Next
    On Error GoTo 0

    Close #1

    MsgBox "Result written to: C:\RibbonNames.txt"
End Sub

Private Sub PrintControls(Controls As CommandControls, LeadingSpace As String, Level As Integer)
    Dim oControl As CommandControl
    For Each oControl In Controls
        If oControl.ControlType = kSeparatorControl Then
            Print #1, LeadingSpace & Space(Level * 4) & "Control: Seperator"
        Else
            Print #1, LeadingSpace & Space(Level * 4) & "Control: " & oControl.DisplayName & ", " & oControl.InternalName & ", Visible: " & oControl.Visible

            If Not oControl.ChildControls Is Nothing Then
                Call PrintControls(oControl.ChildControls, LeadingSpace, Level + 1)
            End If
        End If
    Next
End Sub