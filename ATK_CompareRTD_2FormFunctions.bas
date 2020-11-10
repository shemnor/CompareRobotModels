Attribute VB_Name = "ATK_CompareRTD_2FormFunctions"
Function request_robotmodel_filepath(Optional defaultpath As String) As String

'Display a Dialog Box that allows to select a single file.
'The path for the file picked will be stored in fullpath variable

Dim Fullpath As String
Dim wrongfilemsg As Integer

Retry:

  With Application.FileDialog(msoFileDialogFilePicker)
        'Makes sure the user can select only one file
        .AllowMultiSelect = False
        'Filter to just the following types of files to narrow down selection options
        .Filters.Add "ROBOT Model", "*.rtd", 1
        'set default path
        If defaultpath <> "" Then .InitialFileName = defaultpath
        'Show the dialog box
        .Show
        
        'Store in fullpath variable
        If .SelectedItems.Count < 1 Then Exit Function
        Fullpath = .SelectedItems.Item(1)
    End With
    
    'It's a good idea to still check if the file type selected is accurate.
    'Quit the procedure if the user didn't select the type of file we need.
    If InStr(Fullpath, ".rtd") = 0 Then
        wrongfilemsg = MsgBox("Selected file is not a Robot Model file. Only .rtd files can be selected", vbRetryCancel, "Wrong file type")
        If wrongfilemsg = vbRetry Then
            GoTo Retry
        Else
            Exit Function
        End If
    End If
 
    'set return value to filepath
    request_robotmodel_filepath = Fullpath

End Function
Function get_User_Geometry_Request() As Dictionary

Set get_User_Request = New Dictionary

'set IRobotObjectType to true if requested by user

If ATK_CompareRTD_Form.cbNodes_check.Value = True Then get_User_Request.Add IRobotObjectType.I_OT_NODE, True
If ATK_CompareRTD_Form.cbBars_check.Value = True Then get_User_Request.Add IRobotObjectType.I_OT_BAR, True
If ATK_CompareRTD_Form.cbPanels_check.Value = True Then get_User_Request.Add IRobotObjectType.I_OT_PANEL, True

End Function

Public Function get_UserRequest(InputWorksheet As Worksheet) As Dictionary

Set get_UserRequest = New Dictionary
Dim requests(4) As String

requests = ["Request_nodes","Request_panels","Request_panels","Request_meshes"]

For i = 0 To requests.Count
    If InputWorksheet.Range(requests(i)).Value = True Then
        get_UserRequest.Add requests(i), True
    Else
        get_UserRequest.Add requests(i), False
    End If
Next i

End Function
