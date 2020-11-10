Attribute VB_Name = "ATK_CompareRTD_6RobotVisual"
Sub showChangesInRobot(project As RobotProject, sortedChanges As Dictionary)

Dim oGroupServ As RobotGroupServer
Set oGroupServ = project.Structure.Groups

For Each Group In sortedChanges.Keys
    If Group = "NODES" Or Group = "BARS" Or Group = "PANELS" Then
        For Each Category In sortedChanges(Group).Keys
            changeType = sortedChanges(Group)(Category)
            Select Case Group
                Case "NODES"
                    oGroupServ.Create I_OT_NODE, changeType(0), changeType(2), getCategoryColor(changeType(1))
                Case "BARS"
                    oGroupServ.Create I_OT_BAR, changeType(0), changeType(2), getCategoryColor(changeType(1))
                Case "PANELS"
                    oGroupServ.Create I_OT_PANEL, changeType(0), changeType(2), getCategoryColor(changeType(1))
            End Select
        Next Category
    End If
Next Group

End Sub

Function getCategoryColor(Category As Variant) As Long

Select Case Category
    Case "GEOM"
        getCategoryColor = 3
    Case "SUPP"
        getCategoryColor = 2
    Case "SECTION"
        getCategoryColor = 4
    Case "MAT"
        getCategoryColor = 6
    Case "RELEASE"
        getCategoryColor = 8
    Case "NEW"
        getCategoryColor = 13
    Case "MULTIPLE"
        getCategoryColor = 0
End Select

End Function

Sub updateRobotViewSettings(project As RobotProject)

Dim Rv As RobotView
Dim RVP As RobotViewDisplayParams

Set Rv = project.ViewMngr.GetView(1)
Set RVP = Rv.ParamsDisplay

RVP.Set I_VDA_STRUCTURE_GROUP_COLORS, True
RVP.Set I_VDA_STRUCTURE_BAR_NAMES, True
RVP.Set I_VDA_STRUCTURE_ONLY_FOR_SELECTED_OBJECTS, False
RVP.Set I_VDA_STRUCTURE_NODE_NUMBERS, False
RVP.Set I_VDA_STRUCTURE_BAR_NUMBERS, True
RVP.Set I_VDA_FE_PANEL_NUMBERS, False
Rv.ParamsDisplay.SymbolSize = 1
Rv.Redraw (1)

End Sub
