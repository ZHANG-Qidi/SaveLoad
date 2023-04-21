VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_SelectProfile 
   Caption         =   "SelectProfile"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "UserForm_SelectProfile.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_SelectProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton_Cancel_Click()
    Unload Me
End Sub

Private Sub CommandButton_OK_Click()
    gl_profile_name = Me.ListBox_Profile.Value
    gl_profile_name_range.Value = gl_profile_name
    gl_save_name_all_range.ClearContents
    If IsNull(gl_profile_name) Then
        Unload Me
        Exit Sub
    End If
    gl_profile_path = gl_game_path & "\" & gl_profile_name
    gl_profile_save_write_func
    Unload Me
End Sub

Private Sub UserForm_Click()

End Sub
