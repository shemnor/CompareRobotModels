Attribute VB_Name = "ATK_CompareRTD_1Main"
Option Explicit
Public startTime As Long

Sub resetROBO()

Dim robapp As New RobotApplication

robapp.Interactive = True

End Sub


Sub Main()

'Main control sub.
'
'Purpose of this is to collate all high level information into a clear sequence_
'for a quick overview of the full process and logic. Once user is familiar with the reasoning,_
'its easy to enquire about deeper level process in each step, understanding what objects are being used etc.
'
'(for quick reference to each function/sub - right click onto name and select "definition".)
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim dUserGeometryRequest As Dictionary 'of IRobotObjectType:Boolean key:value pair
    Dim dUserPropertiesRequest As Dictionary
    Dim oRobotApp As New RobotApplication
    Dim oProject As IRobotProject
    Dim dProjectAdata As Dictionary
    Dim dProjectBdata As Dictionary
    Dim dChanges As Dictionary
    Dim reportSheet As Worksheet
    
    'start
    startTime = Timer
    
    'Setup
    Set wb = ThisWorkbook
    Set ws = Sheet1
    oRobotApp.Visible = True
    oRobotApp.Interactive = False
    
    'show progress bar
    ATK_CompareRTD_Form.Hide
    ufProgress.LabelLoop.Width = 0
    ufProgress.LabelProgress.Width = 0
    ufProgress.Show
    
    'Get Project A
    update_progress 0.1, "Opening Project A... Please wait"
    oRobotApp.project.Open (ATK_CompareRTD_Form.tbModelA_address.Text)
    Set oProject = oRobotApp.project
    estimate_script_time oProject
    
    'Get Project A Information
    update_progress 0.2, "Scraping Project A Info"
    Set dProjectAdata = get_Project_information(oProject)
    oProject.Close
    
    'Get Project B
    update_progress 0.3, "Opening Project B... Please wait"
    oRobotApp.project.Open (ATK_CompareRTD_Form.tbModelB_address.Text)
    Set oProject = oRobotApp.project
    
    'Get Project B information
    update_progress 0.4, "Scraping Project B Info"
    Set dProjectBdata = get_Project_information(oProject)
    
    'get output copy of project
    'ATK_CompareRTD_Form.tbOutput_address.Text &
    update_progress 0.5, "Creating a comparision file... Please wait"
    oRobotApp.project.SaveAs ("C:\Users\SZCZ1360\OneDrive Corp\OneDrive - Atkins Ltd\Documents\Robot Tests\Output\compmodel.rtd")
    Set oProject = oRobotApp.project
    
    'Do Comparision
    update_progress 0.6, "Comparing project A and B"
    Set dChanges = comparisionHandler(dProjectAdata, dProjectBdata)
    
    'show in robot
    update_progress 0.7, "Visualising results"
    Call showChangesInRobot(oProject, sortComparisions(dChanges))
    Call updateRobotViewSettings(oProject)
    
    'Create Report
    update_progress 0.8, "Printing results to excel"
    Set reportSheet = createReportSheet()
    Call print_summary(reportSheet, sortComparisions(dChanges))
    Call print_changes(reportSheet, dChanges)
    Call print_file_metadata(reportSheet)
    
    'Finish
    update_progress 0.9, "Saving comparision model... Please wait"
    oProject.Save
    
    'Error Handling
    update_progress 1, "Finished", True
    Debug.Print ("finished")

End Sub

Sub estimate_script_time(proj As RobotProject)

Dim time1 As Double
Dim tmpObj As Object
Dim nodeTime As Double
Dim barTime As Double
Dim timeToOpenModel As Long
Dim remainingTime As Long
Dim objNodeServer As RobotNodeServer
Dim objBarServer As RobotBarServer
Dim objPanelServer As RobotObjObjectServer

Set objNodeServer = proj.Structure.Nodes
Set objBarServer = proj.Structure.Bars
Set objPanelServer = proj.Structure.Objects

time1 = Timer
Set tmpObj = objNodeServer.GetAll().Get(1)
nodeTime = Timer - time1

time1 = Timer
Set tmpObj = objBarServer.GetAll().Get(1)
barTime = Timer - time1

'time1 = Timer
'Set tmpObj = objPanelServer.GetAll().Get(1)
'panelTime = Timer - time1

timeToOpenModel = Timer - startTime

remainingTime = 2 * timeToOpenModel + timeToOpenModel + 2 * nodeTime * objNodeServer.GetAll().Count + 4 * barTime * objBarServer.GetAll().Count

ufProgress.lblTimeRemain = Fix(remainingTime / 60) & "min, " & (remainingTime Mod 60) & "s"

End Sub
Sub update_progress(perc As Double, message As String, Optional finished As Boolean)

Dim timeElapsed As Long

timeElapsed = Timer - startTime
ufProgress.lblTimeElaps = Fix(timeElapsed / 60) & "min, " & (timeElapsed Mod 60) & "s"

ufProgress.LabelProgress.Width = 250 * perc
ufProgress.LabelCaption.Caption = message
If finished = True Then
    ufProgress.Caption = "Script Ended"
    ufProgress.CancelButton.Caption = "Close"
End If

DoEvents

End Sub

Sub update_progress_loop(preMessage As String, loopno As Long, totalobjects As Long, Optional reset As Boolean)

Dim timeElapsed As Long
Dim modlimit As Long
modlimit = 20


timeElapsed = Timer - startTime
ufProgress.lblTimeElaps = Fix(timeElapsed / 60) & "min," & (timeElapsed Mod 60) & "s"

If reset = False Then
    ufProgress.CaptionLoop.Caption = preMessage & " " & Str(loopno) & " / " & Str(totalobjects)
    ufProgress.LabelLoop.Width = loopno * (ufProgress.FrameLoop.Width / totalobjects)
Else
    ufProgress.CaptionLoop.Caption = ""
    ufProgress.LabelLoop.Width = 0
End If

If loopno Mod modlimit = 0 Then DoEvents

End Sub
