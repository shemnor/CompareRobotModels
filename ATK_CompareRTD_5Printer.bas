Attribute VB_Name = "ATK_CompareRTD_5Printer"
Option Explicit

Sub print_summary(ws As Worksheet, sortedChanges As Dictionary)

    Dim FreeRow As Integer
    Dim topGroupRow As Integer
    Dim topObjRow As Integer
    Dim endRow As Integer
    Dim typeColumn As Integer
    Dim countColumn As Integer
    Dim descriptionColumn As Integer
    Dim Group As Variant
    Dim changeType As Variant
    
    FreeRow = ws.Range("sum_table_header").Row + 1
    typeColumn = ws.Range("sum_Type").Column
    countColumn = ws.Range("sum_Count").Column
    descriptionColumn = ws.Range("sum_Desc").Column
    
    Dim obj_changes As Variant
    
    'each group
    For Each Group In sortedChanges.Keys
        
        'save top of group for border formatting
        topGroupRow = FreeRow
        
        'display group type
        ws.Cells(FreeRow, typeColumn).Value = Group
        
        'format group header row
        ws.Range(ws.Cells(FreeRow, typeColumn), ws.Cells(FreeRow, descriptionColumn)).Interior.Color = RGB(217, 217, 217)
        With ws.Range(ws.Cells(FreeRow, typeColumn), ws.Cells(FreeRow, descriptionColumn)).Borders
            .Color = vbBlack
            .Weight = xlThin
        End With
        
        'new line for changeType
        FreeRow = FreeRow + 1
        
        'each changeType (ie NODE SUPPORT or BAR SECTION)
        For Each changeType In sortedChanges(Group).Keys
            If InStr(changeType, "MULTIPLE") = 0 Then
                
                'get changes
                obj_changes = sortedChanges(Group)(changeType)
                
                'save top row of changeType for formatting
                topObjRow = FreeRow
                
                'display how many objects are in this category_
                'works by looking at the list of objects, which is delimited by space
                'Length of list - length of list with spaces removed = spaces
                'where spaces = count of objects
                'as each object has a space in front of its ID
                ws.Cells(FreeRow, countColumn).Value = Len(obj_changes(2)) - Len(Replace(obj_changes(2), " ", ""))
                
                'display message description
                If changeType <> "NEW" And changeType <> "MISSING" Then
                    ws.Cells(FreeRow, descriptionColumn).Value = "change(s) to " & LCase(obj_changes(0))
                ElseIf changeType = "NEW" Then
                    ws.Cells(FreeRow, descriptionColumn).Value = "new " & LCase(Group)
                ElseIf changeType = "MISSING" Then
                    ws.Cells(FreeRow, descriptionColumn).Value = LCase(Group) & " could not be found"
                End If
                
                'new line for next change
                FreeRow = FreeRow + 1
                
            End If
        Next changeType
        
        'new line for next group
        FreeRow = FreeRow + 1
        
        'formats the object type column
        endRow = FreeRow
        ws.Range(ws.Cells(topGroupRow, typeColumn), ws.Cells(endRow - 1, typeColumn)).BorderAround Weight:=xlThin, Color:=vbBlack
        ws.Range(ws.Cells(topGroupRow, countColumn), ws.Cells(endRow - 1, countColumn)).BorderAround Weight:=xlThin, Color:=vbBlack
        ws.Range(ws.Cells(topGroupRow, descriptionColumn), ws.Cells(endRow - 1, descriptionColumn)).BorderAround Weight:=xlThin, Color:=vbBlack
        
    Next Group
    
    ' delete gap to next table
    endRow = ws.Range("full_table_header").Row - 4
    ws.Range(ws.Cells(FreeRow, typeColumn), ws.Cells(endRow, typeColumn)).EntireRow.Delete

End Sub

Sub print_changes(ws As Worksheet, allChanges As Dictionary)

    Dim FreeRow As Integer
    Dim topGroupRow As Integer
    Dim topObjRow As Integer
    Dim endRow As Integer
    Dim typeColumn As Integer
    Dim countColumn As Integer
    Dim descriptionColumn As Integer
    Dim idColumn As Integer
    
    Dim Group As Variant
    Dim object As Variant
    Dim objChanged As Boolean
    Dim i As Integer
    
    FreeRow = ws.Range("full_table_header").Row + 1
    typeColumn = ws.Range("full_type").Column
    idColumn = ws.Range("full_id").Column
    descriptionColumn = ws.Range("full_desc").Column
    
    
    Dim obj_changes As Variant
    
    'each group
    For Each Group In allChanges.Keys
        
        'save top of group for border formatting
        topGroupRow = FreeRow
        
        'display group type
        ws.Cells(FreeRow, typeColumn).Value = Group
        
        'format group header row
        ws.Range(ws.Cells(FreeRow, typeColumn), ws.Cells(FreeRow, descriptionColumn)).Interior.Color = RGB(217, 217, 217)
        With ws.Range(ws.Cells(FreeRow, typeColumn), ws.Cells(FreeRow, descriptionColumn)).Borders
            .Color = vbBlack
            .Weight = xlThin
        End With
        
        'new line for changeType
        FreeRow = FreeRow + 1
        
        'each ID
        For Each object In allChanges(Group).Keys
            
            'get changes
            obj_changes = allChanges(Group)(object)
            
            'check if anything changed
            objChanged = False
            For i = 0 To UBound(obj_changes)
                If IsArray(obj_changes(i)) = True Then
                    objChanged = True
                    Exit For
                End If
            Next i
            
            'if changed then display changes
            If objChanged = True Then
                
                'save top row for formating
                topObjRow = FreeRow
                
                'display object ID
                ws.Cells(FreeRow, idColumn).Value = Left(Group, Len(Group) - 1) & " " & object
                
                'list all changes for this object
                For i = 0 To UBound(obj_changes)
                    If IsArray(obj_changes(i)) = True Then
                        ws.Cells(FreeRow, descriptionColumn).Value = obj_changes(i)(UBound(obj_changes(i)))
                        FreeRow = FreeRow + 1
                    End If
                Next i
                
                'add a line space between next object
                FreeRow = FreeRow + 1
                
                'format borders around object
                endRow = FreeRow
                If topObjRow <> endRow Then
                    ws.Range(ws.Cells(topObjRow, idColumn), ws.Cells(endRow - 1, descriptionColumn)).BorderAround Weight:=xlThin, Color:=vbBlack
                    ws.Range(ws.Cells(topObjRow, idColumn), ws.Cells(endRow - 1, idColumn)).BorderAround Weight:=xlThin, Color:=vbBlack
                End If
            End If
            
        Next object
        
        'save end row for this group
        endRow = FreeRow
        
        'format borders around this group
        ws.Range(ws.Cells(topGroupRow, typeColumn), ws.Cells(endRow - 1, typeColumn)).BorderAround Weight:=xlThin, Color:=vbBlack
    Next Group

End Sub

Sub print_file_metadata(ws As Worksheet)

    Dim oShell As Object
    Dim oDir As Object
    
    Dim sAddress As String
    Dim sDirectory As Variant 'must be variant for shell.namespace method
    Dim sFilename As Variant
    
    Dim i As Integer
    
    Set oShell = CreateObject("Shell.Application")
    
    'PROJ A
    
    'get address
    sAddress = ATK_CompareRTD_Form.tbModelA_address.Text
    'get directory
    sDirectory = Left(sAddress, InStrRev(sAddress, "\"))
    'get filename
    sFilename = Right(sAddress, Len(sAddress) - InStrRev(sAddress, "\"))
    
    Set oDir = oShell.Namespace(sDirectory)
    ws.Range("PA_filename").Value = sFilename
    ws.Range("PA_directory").Value = sDirectory
    ws.Range("PA_created").Value = oDir.GetDetailsOf(oDir.Items.Item(sFilename), 3)
    ws.Range("PA_modified").Value = oDir.GetDetailsOf(oDir.Items.Item(sFilename), 4)
    
    'PROJ B
    
    'get address
    sAddress = ATK_CompareRTD_Form.tbModelA_address.Text
    'get directory
    sDirectory = Left(sAddress, InStrRev(sAddress, "\"))
    'get filename
    sFilename = Right(sAddress, Len(sAddress) - InStrRev(sAddress, "\"))
    
    Set oDir = oShell.Namespace(sDirectory)
    ws.Range("PB_filename").Value = sFilename
    ws.Range("PB_directory").Value = sDirectory
    ws.Range("PB_created").Value = oDir.GetDetailsOf(oDir.Items.Item(sFilename), 3)
    ws.Range("PB_modified").Value = oDir.GetDetailsOf(oDir.Items.Item(sFilename), 4)
    
    
    'REPORT Data
    
    'date
    ws.Range("Ddate").Value = Date
    'time
    ws.Range("Dtime").Value = TimeValue(Now)
    'author
    ws.Range("Dauthor").Value = Application.UserName

End Sub

Function createReportSheet()

    Dim ws As Worksheet
    Dim defaultName As String
    Dim defaultFree As Boolean
    Dim nameTaken As Boolean
    Dim i As Integer
    
    defaultName = Sheet1.Range("defaultReportName").Value
    
        Sheet2.Visible = xlSheetVisible
        Sheet2.Copy After:=Sheet1
        Set createReportSheet = ActiveSheet
        Sheet2.Visible = xlSheetHidden
        
        defaultFree = True
        For Each ws In ThisWorkbook.Sheets
            If ws.name = defaultName Then defaultFree = False
        Next ws
        
        If defaultFree = True Then
            createReportSheet.name = defaultName
        Else
            For i = 1 To 100
                nameTaken = False
                For Each ws In ThisWorkbook.Sheets
                    If ws.name = defaultName & i Then nameTaken = True
                Next ws
                If nameTaken = False Then
                    createReportSheet.name = defaultName & i
                    Exit For
                End If
            Next i
        End If
        
    End Function
