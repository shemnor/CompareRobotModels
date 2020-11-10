VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ATK_CompareRTD_Form 
   Caption         =   "Atkins ROBOT model comparision tool"
   ClientHeight    =   5775
   ClientLeft      =   630
   ClientTop       =   2475
   ClientWidth     =   11325
   OleObjectBlob   =   "ATK_CompareRTD_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ATK_CompareRTD_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBrowseModelA_Click()
Dim address As String

address = request_robotmodel_filepath()
If address <> "" Then Me.tbModelA_address = address

End Sub

Private Sub btnBrowseModelB_Click()
Dim address As String

address = request_robotmodel_filepath()
If address <> "" Then Me.tbModelB_address = address

End Sub

Private Sub btnBrowseOutput_Click()

Dim address As String

address = request_robotmodel_filepath()
If address <> "" Then Me.tbOutput_address = address

End Sub

Private Sub btnRun_Click()
Call Main
End Sub
