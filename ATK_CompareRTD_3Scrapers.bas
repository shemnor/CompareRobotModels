Attribute VB_Name = "ATK_CompareRTD_3Scrapers"
Option Explicit
Function get_Project_information(proj As RobotProject) As Dictionary
    
    Dim objNodeServer As RobotNodeServer
    Dim objBarServer As RobotBarServer
    Dim objPanelServer As RobotObjObjectServer
    Dim objLabelServer As RobotLabelServer
    Dim objCaseServer As RobotCaseServer
    
    Dim nodeData As Dictionary
    Dim barData As Dictionary
    Dim panelData As Dictionary
    Dim supportData As Dictionary
    Dim materialData As Dictionary
    Dim thicknessData As Dictionary
    Dim barReleaseData As Dictionary
    Dim loadsAndCasesData() As Variant
    
    Dim Selection As RobotSelection
    Dim panel_col As RobotObjObjectCollection
    
    Set objNodeServer = proj.Structure.Nodes
    Set objBarServer = proj.Structure.Bars
    Set objPanelServer = proj.Structure.Objects
    Set objLabelServer = proj.Structure.labels
    Set objCaseServer = proj.Structure.cases
    
    Set get_Project_information = New Dictionary
    
    'delete meshes
    update_progress_loop "Deleting meshes", 0, 1
    proj.Structure.Objects.Mesh.Remove
    update_progress_loop "Deleting meshes", 1, 1
                    
    'get node information on all objects in server
    Set nodeData = get_node_information(objNodeServer.GetAll())
    
    'get bar information on all objects in server
    Set barData = get_bar_information(objBarServer.GetAll())
    
    'get panel information on all objects in server
    'panels are grouped with other objects in panelServer so panels first must be selected using the selection server.
    Set Selection = proj.Structure.Selections.Create(I_OT_PANEL)
    Selection.FromText "all"
    Set panelData = get_panel_information(objPanelServer.GetMany(Selection))
    
    'get suport information
    Set supportData = get_support_information(objLabelServer.GetMany(I_LT_SUPPORT))
    
    'get material data
    Set materialData = get_material_information(objLabelServer.GetMany(I_LT_MATERIAL))
    
    'get thickness data
    'labelserver.getmany(tincknesses) returns 0 so _
    'use labelServer.getavailablenames first and create a collection manually.
    'use getLabelColl function in this module to create a collection and pass to the thickness info scraper
    Set thicknessData = get_thickness_information(getLabelColl(I_LT_PANEL_THICKNESS, objLabelServer))
    
    'get release data
    Set barReleaseData = get_barRelease_information(objLabelServer.GetMany(I_LT_BAR_RELEASE))
    
    'get cases and load data
    Set Selection = proj.Structure.Selections.CreatePredefined(I_PS_CASE_SIMPLE_CASES)
    loadsAndCasesData = get_loadsAndCases_information(objCaseServer.GetMany(Selection))
    
    
    'construct the project data dictionary
    get_Project_information.Add "NODES", nodeData
    get_Project_information.Add "BARS", barData
    get_Project_information.Add "PANELS", panelData
    get_Project_information.Add "SUPPORTS", supportData
    get_Project_information.Add "MATERIALS", materialData
    get_Project_information.Add "THICKNESSES", thicknessData
    get_Project_information.Add "BAR_RELEASES", barReleaseData
    get_Project_information.Add "S_CASES", loadsAndCasesData(0)
    get_Project_information.Add "LOADS", loadsAndCasesData(1)

End Function
Function getLabelColl(labelType As IRobotLabelType, labelServer As RobotLabelServer) As Collection
    
    'creates collection of labels based on available labels
    'used when labelserver.getmany(labeltype) returns 0
    
    'for some reason, this happens for thicknesses so fist one must loop through available thickness labels _
    'and request them individually??!!
    'I swear im loosing patience with ROBOT....
    
    Set getLabelColl = New Collection
    Dim availThicknesses As RobotNamesArray
    Dim thickness As Object
    Dim i As Integer
    
    Set availThicknesses = labelServer.GetAvailableNames(labelType)
    
    For i = 1 To availThicknesses.Count
        Set thickness = labelServer.Get(labelType, availThicknesses.Get(i))
        If Not thickness Is Nothing Then
            getLabelColl.Add thickness
        End If
    Next i

End Function

Function get_node_information(node_objects As IRobotCollection) As Dictionary

    'dAllNodeLabelProperties were removed - _
    'Node DOF are restrained by applying labels; _
    'It is therefore the support label not in the node object which has the restraint info. _
    'There is no point for checking restraints for each node, _
    'and its more robust to loop throgh the support server to check the label definition _
    'Original script used On error resume next to filter if a support has been already mapped _
    'Again, this is not robust code and would fail in exceptions. This step is unneccesary if label definition_
    'is checked for differences instead of each node.
    
    Set get_node_information = New Dictionary
    
    Dim vNode_properties(4) As Variant
    Dim node As RobotNode
    
    Dim i As Long
    
    Dim StartTime2 As Double
    StartTime2 = Timer
    Debug.Print "node iteration starts at" & Str(StartTime2)
    
    
    'loop over all nodes in the collection
    For i = 1 To node_objects.Count
        
        update_progress_loop "Scraping node", i, node_objects.Count
        
        'get the node object
        Set node = node_objects.Get(i)
        
        'get properties node number
        vNode_properties(0) = node.Number
        
        'get coordinates (x, y, z)
        vNode_properties(1) = Round(node.x, 3)
        vNode_properties(2) = Round(node.Y, 3)
        vNode_properties(3) = Round(node.Z, 3)
        
        'If node has a label corresponding to a support then:
        'get the label object to get the support data object and save the name
        If node.HasLabel(I_LT_SUPPORT) Then
            vNode_properties(4) = node.GetLabel(I_LT_SUPPORT).name
        Else
            'is node is not supported collect empty string
            vNode_properties(4) = ""
        End If
        
        'now that node information is collected, _
        'add node information to respective dictionaries
        get_node_information.Add vNode_properties(0), vNode_properties
        
    'and move onto the next node
    Next i
    
    'return colected information back to sender
    Debug.Print "all nodes completed in" & Str(Round(Timer - StartTime2, 2))
    
    update_progress_loop " ", 1, 1, True

End Function

Function get_bar_information(bar_objects As IRobotCollection) As Dictionary
    'tested using WITH statements to substitute bar.property or StartNode.property but showed LOWERED performance
    
    
    Set get_bar_information = New Dictionary
    Dim vbar_properties(7) As Variant
    Dim bar As RobotBar
    
    Dim startTime As Double
    Dim StartTime1 As Double
    Dim StartTime2 As Double
    Dim SecondsElapsed As Double
    
    Dim i As Long
    
    StartTime2 = Timer
    Debug.Print "bar iteration starts at" & Str(StartTime2)
    
    'loop over all nodes in the collection
    For i = 1 To bar_objects.Count
    
        update_progress_loop "Scraping bar", i, bar_objects.Count
    
        'get the bar object
        Set bar = bar_objects.Get(i)
    
        'get bar number
        vbar_properties(0) = bar.Number
    
        'get section
        'no real performance improvement with removing haslabel method
        'seems majority of itme is spent on the getlabel data line
        
    '    On Error Resume Next
    '    vbar_properties(1) = bar.GetLabel(I_LT_BAR_SECTION).Data.name
    '    On Error GoTo 0
        
        If bar.HasLabel(I_LT_BAR_SECTION) = True Then
            vbar_properties(1) = bar.GetLabel(I_LT_BAR_SECTION).Data.name
        Else
            vbar_properties(1) = ""
        End If
    
        'get release label name
        If bar.HasLabel(I_LT_BAR_RELEASE) = True Then
        Else
            vbar_properties(2) = "N/A (Default fixed-fixed)"
        End If
    
        'get start nodes
        vbar_properties(3) = bar.StartNode
     
        'get end node
        vbar_properties(4) = bar.EndNode
    
        'get Length
        vbar_properties(5) = Round(bar.Length, 3)
    
        'get Material
        If bar.HasLabel(I_LT_MATERIAL) = True Then
            vbar_properties(6) = bar.GetLabel(I_LT_MATERIAL).Data.name
        Else
            vbar_properties(6) = ""
        End If
    
        'get Gamma angle
        vbar_properties(7) = bar.Gamma
    
        'add array to dictionary
        get_bar_information.Add vbar_properties(0), vbar_properties
    
    'and move onto the next bar
    Next i
    Debug.Print "all bars completed in" & Str(Round(Timer - StartTime2, 2))
    
    'reset progress bar
    update_progress_loop " ", 1, 1, True

End Function

Function get_panel_information(panel_objects As RobotObjObjectCollection) As Dictionary

    Set get_panel_information = New Dictionary
    Dim vPanel_properties(6) As Variant
    Dim panel As RobotObjObject
    Dim panelPart As IRobotObjPart
    Dim panelNodes() As Variant
    Dim panelNode As Object
    
    Dim i As Long
    Dim j As Long
    
    'loop over all nodes in the collection
    For i = 1 To panel_objects.Count
    
        update_progress_loop "Scraping panel", i, panel_objects.Count
    
        'get the panel object
        Set panel = panel_objects.Get(i)
        
        
        'get panel number
        vPanel_properties(0) = panel.Number
        
        
        'get panel thickness
        If panel.HasLabel(I_LT_PANEL_THICKNESS) = True Then
            vPanel_properties(1) = panel.GetLabel(I_LT_PANEL_THICKNESS).name
        Else
            vPanel_properties(1) = ""
        End If
        
    
        'get panel point count
        Set panelPart = panel.GetPart(1)
        vPanel_properties(2) = panelPart.ModelPoints.Count
        
        
        'create array of point coordinates
        ReDim panelNodes(vPanel_properties(2))
        For j = 1 To UBound(panelNodes)
            Set panelNode = panel.Main.DefPoints.Get(j)
            panelNodes(j) = Array(panelNode.x, panelNode.Y, panelNode.Z)
        Next j
        
        'add points to array
        vPanel_properties(3) = getPanelPerimeter(panelNodes)
        
        'get support label name
        If panel.HasLabel(I_LT_EDGE_SUPPORT) = True Then
            vPanel_properties(4) = panel.GetLabel(I_LT_EDGE_SUPPORT).name
        Else
            vPanel_properties(4) = ""
        End If
    
        'add array to dictionary
        get_panel_information.Add vPanel_properties(0), vPanel_properties
    
    'and move onto the next panel
    Next i
    
    'reset progress bar
    update_progress_loop " ", 1, 1, True

End Function

Function get_support_information(supports As IRobotCollection) As Dictionary

    Set get_support_information = New Dictionary
    Dim vSupport_properties(21) As Variant
    Dim support As Object
    Dim supportData As Object
    
    Dim i As Long
    
    For i = 1 To supports.Count
    
        update_progress_loop "Scraping support", i, supports.Count
    
        Set support = supports.Get(i)
        Set supportData = support.Data
    
        'name
        vSupport_properties(0) = support.name
        
        'angles
        vSupport_properties(1) = supportData.Alpha
        vSupport_properties(2) = supportData.Beta
        vSupport_properties(3) = supportData.Gamma
        
        'releases
        vSupport_properties(4) = supportData.RX
        vSupport_properties(5) = supportData.RY
        vSupport_properties(6) = supportData.RZ
        vSupport_properties(7) = supportData.UX
        vSupport_properties(8) = supportData.UY
        vSupport_properties(9) = supportData.UZ
        
        'spring stiffnesses
        vSupport_properties(10) = supportData.AX
        vSupport_properties(11) = supportData.AY
        vSupport_properties(12) = supportData.AZ
        vSupport_properties(13) = supportData.BX
        vSupport_properties(14) = supportData.BY
        vSupport_properties(15) = supportData.BZ
        vSupport_properties(16) = supportData.HX
        vSupport_properties(17) = supportData.HY
        vSupport_properties(18) = supportData.HZ
        vSupport_properties(19) = supportData.KX
        vSupport_properties(20) = supportData.KY
        vSupport_properties(21) = supportData.KZ
        
        'add to return dictionary
        get_support_information.Add Replace(vSupport_properties(0), " ", "_"), vSupport_properties
        
    Next i
    
    'reset progress bar
    update_progress_loop " ", 1, 1, True

End Function

Function get_material_information(materials As IRobotCollection) As Dictionary

    Set get_material_information = New Dictionary
    Dim vMaterial_properties(20) As Variant
    Dim material As Object
    Dim materialData As Object
    
    Dim i As Long
    
    For i = 1 To materials.Count
    
        update_progress_loop "Scraping material", i, materials.Count
    
        Set material = materials.Get(i)
        Set materialData = material.Data
        
        vMaterial_properties(0) = material.name
        vMaterial_properties(1) = materialData.E
        vMaterial_properties(2) = materialData.E_5
        vMaterial_properties(3) = materialData.E_Trans
        vMaterial_properties(4) = materialData.EC_Deformation
        vMaterial_properties(5) = materialData.GMean
        vMaterial_properties(6) = materialData.Kirchoff
        vMaterial_properties(7) = materialData.LX
        vMaterial_properties(8) = materialData.NU
        vMaterial_properties(9) = materialData.RE
        vMaterial_properties(10) = materialData.RE_AxCompr
        vMaterial_properties(11) = materialData.RE_AxTens
        vMaterial_properties(12) = materialData.RE_Bending
        vMaterial_properties(13) = materialData.RE_Shear
        vMaterial_properties(14) = materialData.RE_TrCompr
        vMaterial_properties(15) = materialData.RE_TrTens
        vMaterial_properties(16) = materialData.RO
        vMaterial_properties(17) = materialData.RT
        vMaterial_properties(18) = materialData.Steel_Thermal
        vMaterial_properties(19) = materialData.Timber_Type
        vMaterial_properties(20) = materialData.Type
        
        'add to material dictionary; remove any spaces from name and replace them with underscore
        get_material_information.Add Replace(vMaterial_properties(0), " ", "_"), vMaterial_properties
        
    Next i
    
    update_progress_loop " ", 1, 1, True

End Function

Function get_thickness_information(thicknesses As Collection) As Dictionary

    Set get_thickness_information = New Dictionary
    Dim vThickness_properties(9) As Variant
    Dim thickness As Object
    Dim thicknessData As RobotThicknessData
    Dim thicknessDataData As RobotThicknessHomoData
    Dim i As Long
    
    For i = 1 To thicknesses.Count
    
        update_progress_loop "Scraping thickness", i, thicknesses.Count
        
        
        Set thickness = thicknesses(i)
        Set thicknessData = thickness.Data
        
        vThickness_properties(0) = thickness.name
        vThickness_properties(1) = thicknessData.MaterialName
        vThickness_properties(2) = thicknessData.ThicknessType
        vThickness_properties(3) = thicknessData.Uplift
        If thicknessData.ThicknessType = I_TT_HOMOGENEOUS Then
            Set thicknessDataData = thicknessData.Data
            vThickness_properties(4) = thicknessDataData.Thick1
            vThickness_properties(5) = thicknessDataData.Thick2
            vThickness_properties(6) = thicknessDataData.Thick3
            vThickness_properties(7) = thicknessDataData.ThickConst
            vThickness_properties(8) = thicknessDataData.GetReduction(1#)
            vThickness_properties(9) = thicknessDataData.Type
    '    Might not need it as 0 is default?
    '    Else
    '        vThickness_properties(4) = 0
    '        vThickness_properties(5) = 0
    '        vThickness_properties(6) = 0
    '        vThickness_properties(7) = 0
    '        vThickness_properties(8) = 0
    '        vThickness_properties(9) = 0
        End If
                                                
        
        
        'add to thickness dictionary; remove any spaces from name and replace them with underscore
        get_thickness_information.Add Replace(vThickness_properties(0), " ", "_"), vThickness_properties
        
    Next i
    
    update_progress_loop " ", 1, 1, True

End Function

Function get_barRelease_information(barReleases As IRobotCollection) As Dictionary
'tested using WITH statements to substitute bar.property or StartNode.property but showed LOWERED performance

    Set get_barRelease_information = New Dictionary
    Dim vbarRelease_properties(2) As Variant
    Dim barRelease As RobotBarRelease
    Dim barReleaseData As RobotBarReleaseData
    Dim endRelease As RobotBarEndReleaseData
    
    Dim i As Long
    
    'loop over all nodes in the collection
    For i = 1 To barReleases.Count
    
        update_progress_loop "Scraping bar release", i, barReleases.Count
    
        'get the bar object
        Set barRelease = barReleases.Get(i)
    
        'get bar number
        vbarRelease_properties(0) = barRelease.name
        
        Set barReleaseData = barRelease.Data
        Set endRelease = barReleaseData.StartNode
    
        vbarRelease_properties(1) = Array(endRelease.AX, _
                                    endRelease.AY, _
                                    endRelease.AZ, _
                                    endRelease.BX, _
                                    endRelease.BY, _
                                    endRelease.BZ, _
                                    endRelease.HX, _
                                    endRelease.HY, _
                                    endRelease.HZ, _
                                    endRelease.KX, _
                                    endRelease.KY, _
                                    endRelease.KZ, _
                                    endRelease.RX, _
                                    endRelease.RY, _
                                    endRelease.RZ, _
                                    endRelease.UX, _
                                    endRelease.UY, _
                                    endRelease.UZ _
                                    )
    
        Set endRelease = barReleaseData.EndNode
    
        vbarRelease_properties(2) = Array(endRelease.AX, _
                                    endRelease.AY, _
                                    endRelease.AZ, _
                                    endRelease.BX, _
                                    endRelease.BY, _
                                    endRelease.BZ, _
                                    endRelease.HX, _
                                    endRelease.HY, _
                                    endRelease.HZ, _
                                    endRelease.KX, _
                                    endRelease.KY, _
                                    endRelease.KZ, _
                                    endRelease.RX, _
                                    endRelease.RY, _
                                    endRelease.RZ, _
                                    endRelease.UX, _
                                    endRelease.UY, _
                                    endRelease.UZ _
                                    )
                                    
        'add array to dictionary
        get_barRelease_information.Add Replace(vbarRelease_properties(0), " ", "_"), vbarRelease_properties
    
    'and move onto the next bar
    Next i
    
    update_progress_loop " ", 1, 1, True

End Function

Function get_loadsAndCases_information(cases As IRobotCaseCollection) As Variant()

    'This retursn an Array(simpleCases, loads)(of Dictionary)
    'Loads on objects are defined inside each loadcase, so Its quicker to create the loads and loadcase _
    'dictionaries at the same time.
    
    Dim simpleCases As Dictionary
    Dim simpleCase As RobotSimpleCase
    Dim listOfRecords As String ' used to combine load info to save in loadcase data
    Dim vCase_properties(6) As Variant
    
    Dim loads As Dictionary
    Dim records As RobotLoadRecordMngr
    Dim commonRecord As IRobotLoadRecordCommon
    Dim record As Object
    Dim vRecord_properties(7) As Variant
    Dim recordObjects As Object
    Dim objList As String
    Dim valueList As String
    Dim pointCount As Integer
    
    Set simpleCases = New Dictionary
    Set loads = New Dictionary
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim m As Long
    Dim n As Long
    
    Dim x As Double
    Dim px As Double
    Dim py As Double
    Dim pz As Double
    
    For i = 1 To cases.Count
    
        update_progress_loop "Scraping Simple Cases", i, cases.Count
        
        Set simpleCase = cases.Get(i)
        
        'case number
        vCase_properties(0) = simpleCase.Number
        'case label
        vCase_properties(1) = simpleCase.label
        'case nature
        vCase_properties(2) = simpleCase.Nature
        'name
        vCase_properties(3) = simpleCase.name
        'case type
        vCase_properties(4) = simpleCase.NatureName
        'case analysis type
        vCase_properties(5) = simpleCase.AnalizeType
        
        'look at loads which are attached to this loadcase
        Set records = simpleCase.records
        
        'get amount of loads in loadcase
        vCase_properties(6) = records.Count
        
        'loop through each load in simple case
        For j = 1 To records.Count
            
            'get common interface of RobotLoadRecord to access isautogenerated property
            Set commonRecord = records.Get(j)
            'If autogenerated then skip
            If commonRecord.IsAutoGenerated = False Then
                
                'get load record
                Set record = records.Get(j)
                
                'unique load number (not visible to user in robot)
                vRecord_properties(0) = record.UniqueId
                'loadcase no
                vRecord_properties(1) = simpleCase.Number
                'load index in case
                vRecord_properties(2) = j
                'type of load like BAR_UNIFORM
                vRecord_properties(3) = record.Type
                'what type of object this load can be applies like NODE
                vRecord_properties(4) = record.ObjectType
                
                'loop through object collection and combine in string (sperated by space)
                Set recordObjects = record.Objects
                For k = 1 To recordObjects.Count
                    objList = objList & " " & recordObjects.Get(k)
                Next k
                
                'store the string of objects this load is applied to
                vRecord_properties(5) = objList
                
                'check if load is empty (not applied to anything)
                If objList = "" Then
                    vRecord_properties(6) = True
                Else
                    vRecord_properties(6) = False
                End If
                
                'get load values
                
                'k represents a load value identifier. See p.38 II.1.1.1 of robot API docu
                'load value identifiers are ENUM values which correspond to _
                'specific load parameters. e.g a bar uniform load record (typical bar udl) _
                'can be defined by an enum corresponding to a linear pressure in x, y or z dir. _
                'these enums are 0,1,2 respectively. So using getValue(0 or 1 or 2) _
                'would return the UDL load value. Its not practical to write code _
                'for each type of load (there are 23 loads types) so all ENUMs are checked
                
                If record.Type <> I_LRT_BAR_TRAPEZOIDALE Then
                
                    For k = 0 To 15
                        valueList = valueList & record.GetValue(k) & " "
                    Next k
                    
                    'store the vlaues in array
                    vRecord_properties(7) = Split(valueList, " ")
                    
                'if load is trapezoidal, need to check load values at each point created during definition
                Else
                    'count how many points were defined
                    pointCount = record.pointCount
                    'add 2p/3p/4p to name
                    vRecord_properties(4) = vRecord_properties(4) & " (" & pointCount & "p)"
                    'loop through maximum no of points
                    For k = 1 To 4
                        'if iteration is within boudary of the getpoint method then get information
                        If k <= pointCount Then
                            record.GetPoint k, px, py, pz, x
                        'if out of bounds then populate with 0 to keep a uniform format
                        Else
                            px = 0
                            py = 0
                            pz = 0
                            x = 0
                        End If
                        
                        'combine into a string
                        valueList = valueList & px & " " & py & " " & pz & " " & x & " "
                    Next k
                    
                    'add remaining information to string
                    For k = 8 To 13
                        valueList = valueList & record.GetValue(k) & " "
                    Next k
                    
                    'store the values in array
                    vRecord_properties(7) = Split(valueList, " ")
                    
                End If
                
                'reset strings
                valueList = ""
                objList = ""
    
            End If
        
        'add loadRecord to LOADS
        
        loads.Add vRecord_properties(0) & " (" & vRecord_properties(3) & " Load in loadcase " & vRecord_properties(1) & ")", vRecord_properties
        
        Next j
        
        'add simpleCase to S_CASES
        simpleCases.Add vCase_properties(0), vCase_properties
        
    Next i
    
    get_loadsAndCases_information = Array(simpleCases, loads)
    
    update_progress_loop " ", 1, 1, True

End Function

Function getPanelPerimeter(points() As Variant) As Double
'calculates perimeter based on point coordinates

    Dim Point1 As Variant
    Dim Point2 As Variant
    Dim Length As Double
    
    Dim i As Integer
    
    For i = 1 To UBound(points) - 1
        Point1 = points(i)
        Point2 = points(i + 1)
        Length = Length + ((Point1(0) - Point2(0)) ^ 2 + (Point1(1) - Point2(1)) ^ 2 + (Point1(2) - Point2(2)) ^ 2) ^ (1 / 2)
        
    Next i
    
    getPanelPerimeter = Length

End Function
Function getReleaseName(index As Integer) As String

Dim labels() As String

labels = Array(AX, AY, AZ, BX, BY, BZ, HX, HY, HZ, KX, KY, KZ, RX, RY, RZ, UX, UY, UZ)

getReleaseName = labels(i)
'Select CAse
'Case 1
'    getReleaseName = "AX"
'Case 2
'    getReleaseName = "AY"
'Case 3
'    getReleaseName = "AZ"
'Case 4
'    getReleaseName = "BX"
'Case 5
'    getReleaseName = "BY"
'Case 6
'    getReleaseName = "BZ"
'Case 7
'    getReleaseName = "HX"
'Case 8
'    getReleaseName = "HY"
'Case 9
'    getReleaseName = "HZ"
'Case 10
'    getReleaseName = "KX"
'Case 11
'    getReleaseName = "KY"
'Case 12
'    getReleaseName = "KZ"
'Case 13
'    getReleaseName = "RX"
'Case 14
'    getReleaseName = "RY"
'Case 15
'    getReleaseName = "RZ"
'Case 16
'    getReleaseName = "UX"
'Case 17
'    getReleaseName = "UY"
'Case 18
'    getReleaseName = "UZ"





End Function


Function getLoadTypeName(ByVal Rtype As Integer) As String

    Dim LType As String
        
        Select Case Rtype
            
            Case I_LRT_BAR_UNIFORM
                LType = "Uniform Load"
                
            Case I_LRT_NODE_FORCE
                LType = "Nodal Force"
             
            Case I_LRT_BAR_TRAPEZOIDALE
                LType = "Trapezoidal load"
            
            Case I_LRT_BAR_THERMAL
                LType = "Thermal load"
                
            Case I_LRT_BAR_FORCE_CONCENTRATED
                LType = "Bar force"
                
            Case I_LRT_NODE_DISPLACEMENT
                LType = "Imp displacement"
                
            Case I_LRT_NODE_VELOCITY
                LType = "Imp velocity"
    
            Case I_LRT_NODE_ACCELERATION
                LType = "Imp acceleration"
                
            Case I_LRT_NODE_FORCE_IN_POINT
                LType = "(FE) Force at a point"
                
            Case I_LRT_BAR_DILATATION
                LType = "Dilatation"
                
            Case I_LRT_PRESSURE
                LType = "(FE) Hydrostatic pressure"
                
            Case I_LRT_UNIFORM
                LType = "(FE) Uniform"
    
            Case I_LRT_LINEAR_ON_EDGES
                LType = "(FE) Linear on edges"
                
            Case I_LRT_DEAD
                LType = "Self-weight"
    
            Case I_LRT_BAR_MOMENT_DISTRIBUTED
                LType = "Uniform moment"
                
            Case I_LRT_IN_CONTOUR
                LType = "(FE) Planar on contour"
    
            Case I_LRT_THERMAL_IN_3_POINTS
                LType = "(FE) Thermal load 3p"
                
            Case I_LRT_LINEAR_3D
                LType = "(FE) Linear 2p (3D)"
                
            Case Is = I_LRT_IN_3_POINTS
                LType = "(FE) Planar"
    
        End Select
        
        getLoadTypeName = LType
    
End Function

Function getLoadValueInfo(ByVal loadType As Long, ByVal index As Integer) As Variant

    Dim valueNames() As Variant
    Dim valueBools() As Variant
        
        Select Case loadType
            
            Case I_LRT_BAR_UNIFORM
                valueNames = Array("PX", "PY", "PZ", "", "", _
                                    "", "", "", "Alpha", "Beta", _
                                    "Gamma", "Load coordinate system", "Projection option", "Load position")
                ReDim valueBools(UBound(valueNames))
                valueBools(11) = Array("Global", "Local")
                valueBools(12) = Array("Disabled", "Enabled")
                valueBools(13) = Array("Absolute", "Relative")
                
            Case I_LRT_NODE_FORCE
                valueNames = Array("FX", "FY", "FZ", "MX", "MY", _
                                    "MZ", "", "", "Alpha", "Beta", _
                                    "Gamma")
                ReDim valueBools(UBound(valueNames))
             
            Case I_LRT_BAR_TRAPEZOIDALE Or I_LRT_BAR_TRAPEZOIDALE_MASS
                valueNames = Array("PX1", "PY1", "PZ1", "X1", "PX2", "PY2", "PZ2", "X2", _
                                    "PX3", "PY3", "PZ3", "X3", "PX4", "PY4", "PZ4", "X4", _
                                    "Alpha", "Beta", "Gamma", _
                                    "Load coordinate system", "Projection option", "Load position")
                ReDim valueBools(UBound(valueNames))
                valueBools(20) = Array("Global", "Local")
                valueBools(21) = Array("Disabled", "Enabled")
                valueBools(22) = Array("Absolute", "Relative")
            
            Case I_LRT_BAR_THERMAL
                valueNames = Array("TX", "TY", "TZ", "", "", _
                                    "", "", "", "", "", _
                                    "", "", "", "", "", _
                                    "")
                ReDim valueBools(UBound(valueNames))
                
            Case I_LRT_BAR_FORCE_CONCENTRATED Or I_LRT_BAR_FORCE_CONCENTRATED_MASS
                valueNames = Array("FX", "FY", "FZ", "CX", "CY", _
                                    "CZ", "X", "", "Alpha", "Beta", _
                                    "Gamma", "Load Coordinate system", "Applied to Node", "Load Position", "", _
                                    "")
                ReDim valueBools(UBound(valueNames))
                valueBools(11) = Array("Global", "Local")
                valueBools(13) = Array("Not on node", "On a Node")
                valueBools(13) = Array("Absolute", "Relative")
                
            Case I_LRT_NODE_DISPLACEMENT
                valueNames = Array("FX", "FY", "FZ", "MX", "MY", _
                                    "MZ")
                ReDim valueBools(UBound(valueNames))
                
            Case I_LRT_NODE_VELOCITY
                valueNames = Array("UX", "UY", "UZ", "", "", _
                                    "", "", "", "Alpha", "Beta", _
                                    "Gamma", "", "", "", "", _
                                    "")
                ReDim valueBools(UBound(valueNames))
    
            Case I_LRT_NODE_ACCELERATION
                valueNames = Array("UX", "UY", "UZ", "", "", _
                                    "", "", "", "Alpha", "Beta", _
                                    "Gamma", "", "", "", "", _
                                    "")
                ReDim valueBools(UBound(valueNames))
                
            Case I_LRT_NODE_FORCE_IN_POINT
                valueNames = Array("FX", "FY", "FZ", "MX", "MY", _
                                    "MZ", "", "", "Alpha", "Beta", _
                                    "Gamma", "X", "Y", "Z", "", _
                                    "")
                ReDim valueBools(UBound(valueNames))
                
            Case I_LRT_BAR_DILATATION
                valueNames = Array("UL", "", "", "", "", _
                                    "", "", "", "", "", _
                                    "", "", "", "Load Position", "", _
                                    "")
                ReDim valueBools(UBound(valueNames))
                valueBools(13) = Array("Absolute", "Relative")
                
            Case I_LRT_PRESSURE
                valueNames = Array("P", "RO", "H", "", "", _
                                    "", "Direction")
                ReDim valueBools(UBound(valueNames))
                valueBools(6) = Array("-X", "-Y", "-Z", "X", "Y", "Z")
                
            Case I_LRT_UNIFORM
                valueNames = Array("PX", "PY", "PZ", "", "", _
                                    "", "", "", "", "", _
                                    "", "Load Coordinate system", "Projection option")
                ReDim valueBools(UBound(valueNames))
                valueBools(11) = Array("Global", "Local")
                valueBools(12) = Array("Disabled", "Enabled")
    
            Case I_LRT_LINEAR_ON_EDGES
                valueNames = Array("(FE) Linear on edges", "")
                
            Case I_LRT_DEAD
                valueNames = Array("X Direction", "Y Direction", "Z Direction", "Factor", "", _
                                    "", "", "", "", "", _
                                    "", "", "", "", "", "Extent")
                ReDim valueBools(UBound(valueNames))
                valueBools(11) = Array("Global", "Local")
                valueBools(12) = Array("Disabled", "Enabled")
                valueBools(13) = Array("Absolute", "Relative")
                valueBools(15) = Array("Part of Structure", "Whole Structure")
    
            Case I_LRT_BAR_MOMENT_DISTRIBUTED
                valueNames = Array("MX", "MY", "MZ", "", "", _
                                    "", "", "", "", "", _
                                    "", "Load Coordinate System", "", "", "", _
                                    "")
                ReDim valueBools(UBound(valueNames))
                valueBools(11) = Array("Global", "Local")
                
            Case I_LRT_IN_CONTOUR
                valueNames = Array("PX1", "PY1", "PZ1", "PX2", "PY2", _
                                    "PZ2", "PX3", "PY3", "PZ3", "", _
                                    "", "", "Projection option", "NPoints", "", _
                                    "Load Coordinate system")
                ReDim valueBools(UBound(valueNames))
                valueBools(12) = Array("Disabled", "Enabled")
                valueBools(15) = Array("Global", "Local")
    
            Case I_LRT_THERMAL_IN_3_POINTS
                valueNames = Array("TX1", "", "TZ1", "TX2", "", "TZ2", "TX3", _
                                    "", "TZ3", "", "", "", "", _
                                    "", "NPoints")
                ReDim valueBools(UBound(valueNames))
                
            Case I_LRT_LINEAR_3D
                valueNames = Array("PX1", "PY1", "PZ1", "MX1", "MY1", _
                                    "MZ1", "PX2", "PY2", "PZ2", "MX2", _
                                    "MY2", "MZ2", "", "Load Coordinate system", "", _
                                    "Gamma")
                ReDim valueBools(UBound(valueNames))
                valueBools(13) = Array("Global", "Local")
                
            Case Is = I_LRT_IN_3_POINTS
                valueNames = Array("PX1", "PY1", "PZ1", "PX2", "PY2", _
                                    "PZ2", "PX3", "PY3", "PZ3", "N1", _
                                    "N2", "N3", "Projection option", "Load Coordinate system", "", _
                                    "")
                ReDim valueBools(UBound(valueNames))
                valueBools(11) = Array("Disabled", "Enabled")
                valueBools(13) = Array("Global", "Local")

    
        End Select
        
        getLoadValueInfo = Array(valueNames(index), valueBools(index))
    
End Function
