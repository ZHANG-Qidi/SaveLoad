VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_SelectGame 
   Caption         =   "SelectGame"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "UserForm_SelectGame.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_SelectGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton_Cancel_Click()
    Unload Me
End Sub

Private Sub CommandButton_OK_Click()
    gl_game_name_range.Value = Me.ListBox_Game.Value
    gl_profile_name_range.MergeArea.ClearContents
    gl_save_name_all_range.ClearContents
    Unload Me
End Sub

Private Sub UserForm_Click()

End Sub
