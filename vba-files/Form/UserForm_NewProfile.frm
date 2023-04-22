VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_NewProfile 
   Caption         =   "NewProfile"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "UserForm_NewProfile.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_NewProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton_Cancel_Click()
    Unload Me
End Sub

Private Sub CommandButton_OK_Click()
    gl_game_name = Me.TextBox_Game.Text
    gl_game_path = ThisWorkbook.Path & "\SaveLoad\" & gl_game_name
    gl_profile_name = Me.TextBox_Profile.Text
    If gl_profile_name = "" Then
        MsgBox "Please Input Profile Name"
        Exit Sub
    End If
    gl_profile_path = gl_game_path & "\" & gl_profile_name
    With CreateObject("Scripting.FileSystemObject")
        If .FolderExists(gl_profile_path) = False Then
            .CreateFolder gl_profile_path
        End If
    End With
    gl_profile_name_range.Value = gl_profile_name
    gl_save_name_all_range.ClearContents
    gl_profile_save_write_func
    Unload Me
End Sub

Private Sub UserForm_Click()

End Sub
