VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_DeleteSave 
   Caption         =   "DeleteSave"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "UserForm_DeleteSave.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_DeleteSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton_Cancel_Click()
    Unload Me
End Sub

Private Sub CommandButton_OK_Click()
    CreateObject("Scripting.FileSystemObject").DeleteFolder gl_save_path_selected, True
    gl_save_name_selected_range.MergeArea.ClearContents
    Unload Me
End Sub

Private Sub UserForm_Click()

End Sub
