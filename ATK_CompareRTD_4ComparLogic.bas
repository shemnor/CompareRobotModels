Attribute VB_Name = "ATK_CompareRTD_4ComparLogic"
Option Explicit
Function comparisionHandler(projA_data As Dictionary, projB_data As Dictionary) As Dictionary

'quick logic summary:
'
'for each group
'   for each object
'        for each property
'            if property is complex
'                compare individual components
'            if group is "loads"
'                do a "load comparision"
'            else - property is a simple value
'                do a direct comparision
'            End If
'        Next property

'    Store comparirison results
'    Delete object from projB

'    If object was not found in projB
'        store As deleted
'    End If

'    Next Object

'If anything is left in projB
'    store as new
'End If

'Next Group
        
Set comparisionHandler = New Dictionary

Dim dAllChanges_Groups As Dictionary
Dim dAllChanges_Objects As Dictionary
Dim groupA As Dictionary
Dim groupB As Dictionary

Dim Group As Variant

Set dAllChanges_Groups = New Dictionary


'loop through groups
For Each Group In projA_data.Keys

    Set groupA = projA_data(Group)
    Set groupB = projB_data(Group)
    
    Set dAllChanges_Objects = New Dictionary
    
    'check if other than loads
    If Group <> "LOADS" Then
        Set dAllChanges_Objects = compareProperties(groupA, groupB)
    End If
    
    'check If loads
    If Group = "LOADS" Then
        Set dAllChanges_Objects = compareLoads(groupA, groupB)
    End If
     
    'save all object changes to main group dictionary
    dAllChanges_Groups.Add Group, dAllChanges_Objects
    
Next Group

Set comparisionHandler = dAllChanges_Groups

End Function
    
   
Function compareProperties(groupA As Dictionary, groupB As Dictionary) As Dictionary

Dim dAllChanges As Dictionary
Dim sObjChanges() As Variant

Dim objID As Variant

Dim propA As Variant
Dim propB As Variant
Dim pack() As Variant
Dim comp() As Variant

Dim sNewObj(0) As Variant

Dim i As Integer
Dim j As Integer


Set dAllChanges = New Dictionary

    'loop through ROBOT objects
    For Each objID In groupA.Keys
        
        'if objID in project B exists, that means both projects have an object with the same ID, so we can compare their properties
        If groupB.Exists(objID) Then
        
            'reset array tracking changes in properties
            Erase sObjChanges
            'resize accordingly to amount of properties in array
            ReDim Preserve sObjChanges(UBound(groupA(objID)))
            
            'loop through ROBOT object properties
            'Informaiton is in array format, so we loop though each index.
            For i = 0 To UBound(sObjChanges)
                    
                'get property vales
                propA = groupA(objID)(i)
                propB = groupB(objID)(i)
                
                'for property values which are arrays (such as release definitions that need definition for each DOF)
                If IsArray(propA) = True Or IsArray(propB) = True Then
                    
                    '----------------------------------------------------------
                    'check each value for this property
                    For j = 0 To UBound(propA)
                        'if any dont match then exit check
                        If propA(j) <> propB(j) Then
                            'pack info; propA and B are not required because messages are general
                            If Group = "LOADS" Then
                                pack = Array(propA(j), propB(j), i, j, groupA(objID)(3))
                            Else
                                pack = Array(propA(j), propB(j), i, j)
                            End If
                            'get message back
                            sObjChanges(i) = getPresetMessage_ByGroup(Group, pack)
                        'end if property matches
                        End If
                    'next array index in property
                    Next j
                    '----------------------------------------------------------
               
               'for properties which are single values
                Else
                    
                    '----------------------------------------------------------
                    'if property in each project does not match then compare it
                    If propA <> propB Then
                        'pack object property information
                        pack = Array(propA, propB, i)
                        'save message
                        sObjChanges(i) = getPresetMessage_ByGroup(Group, pack)
                    'end if property matches
                    End If
                    '----------------------------------------------------------
                
                'end if property is array
                End If
            'next property in object
            Next i
            
            'since obj was found, remove from projectB data to improve loop speed
            groupB.Remove (objID)
    
            dAllChanges.Add objID, sObjChanges
        
        'obj was not found in project B
        Else
            sNewObj(0) = Array("MISSING", "MISSING", "Has been replaced or deleted")
            dAllChanges_Objects.Add objID, sNewObj
        'end if exists
        End If
    'next object in group
    Next objID



End Function


    
    
    


                
                
                            If Group = "LOADS" Then
                                pack = Array(propA(j), propB(j), i, j, groupA(objID)(3))
                            Else
                                pack = Array(propA(j), propB(j), i, j)
                            End If
                                


                

                
                    
NextProp:

            



        

    
    'all objects in projectA for this group have now been checked,
    'if there are any objs left in projectB data_
    'these must be new obj IDs which are not present in projectA data.
    'loop through all remaining values and add message
    If groupB.Count > 0 Then
        For Each objID In groupB.Keys
            pack = Array(propA, propB, 100)
            sNewObj(0) = compareProperties_ByGroup(Group, pack)
            dAllChanges_Objects.Add objID, sNewObj
        Next objID
    End If
    
    'save all object changes to main group dictionary
    dAllChanges_Groups.Add Group, dAllChanges_Objects
    
Next Group

Set comparisionHandler = dAllChanges_Groups

End Function

Function getPresetMessage_ByGroup(groupName As Variant, pack() As Variant) As Variant()

'select approperiate function depending on group

Select Case groupName

    Case "NODES"
        getPresetMessage_ByGroup = getNodeMessage(pack)
    Case "BARS"
        getPresetMessage_ByGroup = compareBarProperties(pack)
    Case "PANELS"
        getPresetMessage_ByGroup = comparePanelProperties(pack)
    Case "SUPPORTS"
        getPresetMessage_ByGroup = compareSupportProperties(pack)
    Case "MATERIALS"
        getPresetMessage_ByGroup = compareMaterialProperties(pack)
    Case "THICKNESSES"
        getPresetMessage_ByGroup = compareThicknessProperties(pack)
    Case "BAR_RELEASES"
        getPresetMessage_ByGroup = compareBarReleaseProperties(pack)
    Case "S_CASES"
        getPresetMessage_ByGroup = compareSimpleCases(pack)
    Case "LOADS"
        getPresetMessage_ByGroup = compareLoads(pack)
    
End Select

End Function

Function getNodeMessage(pack() As Variant) As Variant()

Dim message As String
Dim groupName As String
Dim colorGroup As String
Dim direction() As String

'select approperiate message depending on what property doesnt match
Select Case pack(2)
    
    'for properties X, Y or Z
    Case 1, 2, 3
    direction = Split(" ,X,Y,Z", ",")
        message = "Moved in " & direction(pack(2)) & " by " & Round(pack(1) - pack(0), 3)
        colorGroup = "GEOM"
        groupName = "NODE LOCATION"
    
    'for support property
    Case 4
        message = "Support changed from " & pack(0) & " to " & pack(1)
        colorGroup = "SUPP"
        groupName = "NODE SUPPORT"
    Case 100
        message = "Has been added or created in place of previous object"
        colorGroup = "NEW"
        groupName = "NEW"

End Select

getNodeMessage = Array(groupName, colorGroup, message)

End Function

Function getBarMessage(pack() As Variant) As Variant()

Dim message As String
Dim groupName As String
Dim colorGroup As String


'select approperiate message depending on what property doesnt match
Select Case pack(2)

    Case 1
        message = "Section changed from " & pack(0) & " to " & pack(1)
        colorGroup = "SECTION"
        groupName = "BAR SECTION"
    Case 2
        message = "Release at end node changed from " & pack(0) & " to " & pack(1)
        colorGroup = "RELEASE"
        groupName = "BAR END RELEASE"
    Case 3
        message = "Start node changed from " & pack(0) & " to " & pack(1)
        colorGroup = "GEOM"
        groupName = "BAR START NODE"
    Case 4
        message = "End node changed from " & pack(0) & " to " & pack(1)
        colorGroup = "GEOM"
        groupName = "BAR END NODE"
    Case 5
        message = "Length changed from  " & pack(0) & " to " & pack(1)
        colorGroup = "GEOM"
        groupName = "BAR LENGTH"
    Case 6
        message = "Material changed from  " & pack(0) & " to " & pack(1)
        colorGroup = "MAT"
        groupName = "BAR MATERIAL"
    Case 7
        message = "Gamma angle changed from  " & pack(0) & " to " & pack(1)
        colorGroup = "SECTION"
        groupName = "BAR GAMMA"
    Case 100
        message = "Has been added or created in place of previous object"
        colorGroup = "NEW"
        groupName = "NEW"
End Select

getBarMessage = Array(groupName, colorGroup, message)

End Function

Function getPanelMessage(pack() As Variant) As Variant()

Dim propA As Variant
Dim propB As Variant
Dim j As Integer
Dim message As String
Dim groupName As String
Dim colorGroup As String

propA = pack(0)
propB = pack(1)

'select approperiate message depending on what property doesnt match
Select Case pack(2)
    
    Case 1
        message = "Thickness changed from " & pack(0) & " to " & pack(1)
        colorGroup = "SECTION"
        groupName = "PANEL THICKNESS"
    Case 2
        message = "Point count changed from " & pack(0) & " to " & pack(1)
        colorGroup = "GEOM"
        groupName = "PANEL POINT COUNT"
    Case 3
        message = "Panel perimeter length changed from " & Round(pack(0), 3) & " to " & Round(pack(1), 3)
        colorGroup = "GEOM"
        groupName = "PANEL SIZE"
    Case 4
        message = "Edge support changed " & pack(0) & " to " & pack(1)
        colorGroup = "SUPP"
        groupName = "PANEL SUPPORT"
    Case 100
        message = "Has been added or created in place of previous object"
        colorGroup = "NEW"
        groupName = "NEW"
End Select

getPanelMessage = Array(groupName, colorGroup, message)

End Function

Function getSupportMessage(pack() As Variant) As Variant()

Dim propA As Variant
Dim propB As Variant
Dim j As Integer
Dim message As String
Dim groupName As String

propA = pack(0)
propB = pack(1)

groupName = "SUPPORT PROPERTY"

'select approperiate message depending on what property doesnt match
Select Case pack(2)
    
    Case 1
        message = "Alpha changed from " & pack(0) & " to " & pack(1)
    Case 2
        message = "Beta changed from " & pack(0) & " to " & pack(1)

    Case 3
        message = "Gamma changed from " & pack(0) & " to " & pack(1)

    Case 4
        message = "Translation restraint in X changed from " & pack(0) & " to " & pack(1)

    Case 5
        message = "Translation restraint in Y changed from " & pack(0) & " to " & pack(1)

    Case 6
        message = "Translation restraint in Z changed from " & pack(0) & " to " & pack(1)

    Case 7
        message = "Rotation restraint in X changed from " & pack(0) & " to " & pack(1)

    Case 8
        message = "Rotation restraint in Y changed from " & pack(0) & " to " & pack(1)

    Case 9
        message = "Rotation restraint in Z changed from " & pack(0) & " to " & pack(1)

    Case 10
        message = "AX changed from " & pack(0) & " to " & pack(1)

    Case 11
        message = "AY changed from " & pack(0) & " to " & pack(1)

    Case 12
        message = "AZ changed from " & pack(0) & " to " & pack(1)

    Case 13
        message = "BX changed from " & pack(0) & " to " & pack(1)

    Case 14
        message = "BY changed from " & pack(0) & " to " & pack(1)

    Case 15
        message = "BZ changed from " & pack(0) & " to " & pack(1)

    Case 16
        message = "HX changed from " & pack(0) & " to " & pack(1)

    Case 17
        message = "HY hanged from " & pack(0) & " to " & pack(1)

    Case 18
        message = "HZ hanged from " & pack(0) & " to " & pack(1)

    Case 19
        message = "KX hanged from " & pack(0) & " to " & pack(1)

    Case 20
        message = "KX hanged from " & pack(0) & " to " & pack(1)

    Case 21
        message = "KX hanged from " & pack(0) & " to " & pack(1)
    Case 100
        message = "Has been applied to an object"
        groupName = "NEW"
End Select

getSupportMessage = Array(groupName, groupName, message)

End Function

Function getMaterialMessage(pack() As Variant) As Variant()

Dim propA As Variant
Dim propB As Variant
Dim message As String
Dim groupName As String

propA = pack(0)
propB = pack(1)

groupName = "MATERIAL PROPERTY"

'select approperiate message depending on what property doesnt match
Select Case pack(2)
    
    Case 1
        message = "Young's modulus/axial modulus changed from " & pack(0) & " to " & pack(1)
    
    Case 2
        message = "Young's 5% modulus (timber) changed from " & pack(0) & " to " & pack(1)

    Case 3
        message = "Young's transverse modulus (timber) changed from " & pack(0) & " to " & pack(1)

    Case 4
        message = "Shear deformation changed from " & pack(0) & " to " & pack(1)

    Case 5
        message = "Shear Modulus changed from " & pack(0) & " to " & pack(1)

    Case 6
        message = "Transversal shear modulus changed from " & pack(0) & " to " & pack(1)

    Case 7
        message = "Thermal expansion coeff. changed from " & pack(0) & " to " & pack(1)

    Case 8
        message = "Poisson ratio changed from " & pack(0) & " to " & pack(1)
    
    Case 9
        message = "Yield point (alu, steel), compression resistance (concr) changed from " & pack(0) & " to " & pack(1)
    
    Case 10
        message = "Axial compression res. changed from " & pack(0) & " to " & pack(1)

    Case 11
        message = "Axial tension res. changed from " & pack(0) & " to " & pack(1)

    Case 12
        message = "Bending res. (timber) changed from " & pack(0) & " to " & pack(1)

    Case 13
        message = "shear res. (timber) changed from " & pack(0) & " to " & pack(1)

    Case 14
        message = "Transverse compression res. changed from " & pack(0) & " to " & pack(1)

    Case 15
        message = "Transverse tension res. changed from " & pack(0) & " to " & pack(1)

    Case 16
        message = "Material density changed from " & pack(0) & " to " & pack(1)

    Case 17
        message = "Tension res. (steel) changed from " & pack(0) & " to " & pack(1)

    Case 18
        message = "Steel Themrmal option hanged from " & pack(0) & " to " & pack(1)

    Case 19
        message = "Timber type hanged from " & pack(0) & " to " & pack(1)

    Case 20
        message = "Material type hanged from " & pack(0) & " to " & pack(1)
    Case 100
        message = "Has been applied to an object"
        groupName = "NEW"
End Select

getMaterialMessage = Array(groupName, groupName, message)

End Function
Function getThicknessMessage(pack() As Variant) As Variant()

Dim propA As Variant
Dim propB As Variant
Dim message As String
Dim groupName As String

propA = pack(0)
propB = pack(1)

groupName = "THICKNESS PROPERTY"

'select approperiate message depending on what property doesnt match
Select Case pack(2)
    
    Case 1
        message = "Material changed from " & pack(0) & " to " & pack(1)
    Case 2
        message = "Thickness type changed from " & pack(0) & " to " & pack(1)

    Case 3
        message = "Thickness Uplift changed from " & pack(0) & " to " & pack(1)

    Case 4
        message = "Thickness 1 changed from " & pack(0) & " to " & pack(1)

    Case 5
        message = "Thickness 2 changed from " & pack(0) & " to " & pack(1)

    Case 6
        message = "Thickness 3 changed from " & pack(0) & " to " & pack(1)

    Case 7
        message = "Constant Thickness changed from " & pack(0) & " to " & pack(1)

    Case 8
        message = "Stiffness reduction changed from " & pack(0) & " to " & pack(1)

    Case 9
        message = "Homogenous variation type changed from " & pack(0) & " to " & pack(1)
    Case 100
        message = "Has been applied to an object"
        groupName = "NEW"
End Select

getThicknessMessage = Array(groupName, groupName, message)

End Function

Function getBarReleaseMessage(pack() As Variant) As Variant()

Dim propA As Variant
Dim propB As Variant
Dim message As String
Dim groupName As String
Dim labels() As String

propA = pack(0)
propB = pack(1)

labels = Array(AX, AY, AZ, BX, BY, BZ, HX, HY, HZ, KX, KY, KZ, RX, RY, RZ, UX, UY, UZ)

groupName = "BAR RELEASES"

'select approperiate message depending on what property doesnt match
Select Case pack(2)
    
    Case 1
        message = "Release property " & Array(pack(3)) & " at the begining of bar changed from " & pack(0) & " to " & pack(1)
    Case 2
        message = "Release property " & Array(pack(3)) & " at the end of bar changed from " & pack(0) & " to " & pack(1)
    Case 100
        message = "Has been applied to an object"
        groupName = "NEW"
End Select

getBarReleaseMessage = Array(groupName, groupName, message)

End Function

Function getSimpleCaseMessage(pack() As Variant) As Variant()

Dim propA As Variant
Dim propB As Variant
Dim message As String
Dim groupName As String

propA = pack(0)
propB = pack(1)

groupName = "SIMPLE_CASES"

'select approperiate message depending on what property doesnt match
Select Case pack(2)
    
    Case 1
        message = "Label changed from " & pack(0) & " to " & pack(1)
    Case 2
        message = "Nature changed from " & pack(0) & " to " & pack(1)
    Case 3
        message = "Name changed from " & pack(0) & " to " & pack(1)
    Case 4
        message = "Nature name changed from " & pack(0) & " to " & pack(1)
    Case 5
        message = "Analysis type changed from " & pack(0) & " to " & pack(1)
    Case 6
        message = "Number of load definitions changed from " & pack(0) & " to " & pack(1)
    Case 100
        message = "Has been applied to an object"
        groupName = "NEW"
End Select

getSimpleCaseMessage = Array(groupName, groupName, message)

End Function

Function getLoadMessage(pack() As Variant) As Variant()

Dim propA As Variant
Dim propB As Variant
Dim message As String
Dim groupName As String
Dim labels() As String
Dim LoadValuePreset As Variant

propA = pack(0)
propB = pack(1)


groupName = "LOADS"

'select approperiate message depending on what property doesnt match
Select Case pack(2)
    
    Case 1
        message = "Simple case Number changed from " & pack(0) & " to " & pack(1)
    Case 2
        message = "Load index (when sorted by simple case) changed from " & pack(0) & " to " & pack(1)
    Case 3
        message = "Load type changed from " & getLoadTypeName(pack(0)) & " to " & getLoadTypeName(pack(1))
    Case 4
        message = "Allowable object type changed from " & pack(0) & " to " & pack(1)
    Case 5
        message = "List of objects changed from " & pack(0) & " to " & pack(1)
    Case 6
        If pack(1) = True Then
            message = "Load was removed from all objects in model"
        Else
            message = "Load was re-applied to model"
        End If
    Case 7
        LoadValuePreset = getLoadValueInfo(pack(4), pack(3))
        If Not (LoadValuePreset(1) = Empty) Then
            message = "Load definition value for " & LoadValuePreset(1) & " changed from " & LoadValuePreset(1)(pack(0)) & " to " & LoadValuePreset(1)(pack(1))
        Else
            message = "Load definition value for " & LoadValuePreset(0) & " changed from " & pack(0) & " to " & pack(1)
        End If
    Case 100
        message = "Has been applied to an object"
        groupName = "NEW"
End Select

getLoadMessage = Array(groupName, groupName, message)

End Function

Function compareLoads(pack() As Variant) As Variant()

Dim propA As Variant
Dim propB As Variant
Dim message As String
Dim groupName As String
Dim labels() As String
Dim LoadValuePreset As Variant

propA = pack(0)
propB = pack(1)


groupName = "LOADS"

'select approperiate message depending on what property doesnt match
Select Case pack(2)
    
    Case 1
        message = "New load added to case " & pack(1)
    Case 2
        message = "Load index (when sorted by simple case) changed from " & pack(0) & " to " & pack(1)
    Case 3
        message = "Defined as a " & getLoadTypeName(pack(1))
    Case 4
        message = "Allowable object type changed from " & pack(0) & " to " & pack(1)
    Case 5
        message = "Applied to objects " & pack(1)
    Case 6
        If pack(1) = True Then
            message = "Load was removed from all objects in model"
        Else
            message = "Load was re-applied to model"
        End If
    Case 7
        LoadValuePreset = getLoadValueInfo(pack(4), pack(3))
        If Not (LoadValuePreset(1) = Empty) Then
            message = "Load definition value for " & LoadValuePreset(1) & " changed from " & LoadValuePreset(1)(pack(0)) & " to " & LoadValuePreset(1)(pack(1))
        Else
            message = "Load definition value for " & LoadValuePreset(0) & " changed from " & pack(0) & " to " & pack(1)
        End If
    Case 100
        message = "Has been applied to an object"
        groupName = "NEW"
End Select

compareLoads = Array(groupName, groupName, message)

End Function


Function sortComparisions(allChanges As Dictionary) As Dictionary
'fucntion groups nodes, bars and panels by type of change
'this is used for making groups in ROBOT
'end result is to have lists of all object that have undergone a certain change
'eg
'NODES
'   NODE SUPP (name of change type shown in robot)
'       NODE SUPP (same)
'       SUPP (change group used to set the group color)
'       1 2 3 4 5 (list of items)
'   NODE LOCATION
'       NODE LOCATION
'       GEOM
'       5 6 7 8


Set sortComparisions = New Dictionary

Dim dGroups As Dictionary
Dim Group As Variant
Dim objectID As Variant
Dim objProperties As Variant
Dim propertyChanges As Variant

Dim changeType As String
Dim changeArray As Variant
Dim changesCount As Integer

Dim changeTypes As Dictionary

Set dGroups = New Dictionary
Set changeTypes = New Dictionary

Dim i As Integer

'each group
For Each Group In allChanges.Keys
    
    'If Group = "NODES" Or Group = "BARS" Or Group = "PANELS" Then
        'reset dict holding changes for this object type
        Set changeTypes = New Dictionary
        
        'check each object for this object type
        For Each objectID In allChanges(Group).Keys
        
            'get properties array
            objProperties = allChanges(Group)(objectID)
            'reset changecounter
            changesCount = 0
            
            'for each property
            For i = 0 To UBound(objProperties)
            
                'get object representing all changes to this property
                propertyChanges = objProperties(i)
                
                'if the object is an array (ie changes have been recorded) then group them by category
                
                If IsArray(propertyChanges) = True Then
                    'eg BAR START NODE - represents name of change
                    changeType = propertyChanges(0)
                    
                    'if this type of change has already been found; add this object id to the list
                    If changeTypes.Exists(changeType) Then
                        changeArray = changeTypes(changeType)
                        changeArray(2) = changeArray(2) & " " & objectID
                        changeTypes(changeType) = changeArray
                        changesCount = changesCount + 1
                    'if this change type has not yet been found; make one
                    Else
                        'key:value pair - changeType:{changeType, colorGroup, objectList}
                        changeTypes.Add changeType, Array(changeType, propertyChanges(1), " " & objectID)
                        changesCount = changesCount + 1
                    End If
                End If
            Next i
            'if more than one change has been found tfor this object, then add it to the multiple group
            If changesCount > 1 Then
                    'if multiple group already exists then add it to the list
                    If changeTypes.Exists(Group & " " & "MULTIPLE") Then
                        changeArray = changeTypes(Group & " " & "MULTIPLE")
                        changeArray(2) = changeArray(2) & " " & objectID
                        changeTypes(Group & " " & "MULTIPLE") = changeArray
                    'if multuiple group has not yet been made; then make one
                    Else
                        changeTypes.Add Group & " " & "MULTIPLE", Array(Group & " " & "MULTIPLE", "MULTIPLE", " " & objectID)
                    End If
            End If
        
        'do next object
        Next objectID
        
        'all objects have been checked for changes; add the dictionary to main container holding each object group
        dGroups.Add Group, changeTypes
    
    'End If
    
'check next group
Next Group

'return output dictionary
Set sortComparisions = dGroups

End Function

