VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufProgress 
   Caption         =   "Running..."
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   OleObjectBlob   =   "ufProgress.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CancelButton_Click()

If Me.Caption <> "Script Ended" Then
    End
Else
    Unload ufProgress
End If

End Sub
