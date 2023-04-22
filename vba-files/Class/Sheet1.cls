VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton_Game_New_Click()
    gl_variable_read_func
    UserForm_NewGame.Show
End Sub

Private Sub CommandButton_Game_Select_Click()
    gl_variable_read_func
    With UserForm_SelectGame
        With .ListBox_Game
            .Clear
            Dim Folder As Variant
            For Each Folder In CreateObject("Scripting.FileSystemObject").GetFolder(ThisWorkbook.Path & "\SaveLoad").SubFolders
                .AddItem Right(Folder, Len(Folder) - InStrRev(Folder, "\"))
            Next
        End With
        .Show
    End With
End Sub

Private Sub CommandButton_Game_Delete_Click()
    gl_variable_read_func
    If gl_game_name = "" Then
        Exit Sub
    End If
    With UserForm_DeleteGame
        .Label_Game.Caption = "Are you sure to Delete the Game: " & gl_game_name & " ?"
        .Show
    End With
End Sub

Private Sub CommandButton_Profile_New_Click()
    gl_variable_read_func
    If gl_game_name = "" Then
        Exit Sub
    End If
    With UserForm_NewProfile
        .TextBox_Game.Text = gl_game_name
        .Show
    End With
End Sub

Private Sub CommandButton_Profile_Select_Click()
    gl_variable_read_func
    If gl_game_name = "" Then
        Exit Sub
    End If
    With UserForm_SelectProfile
        With .ListBox_Profile
            .Clear
            Dim Folder As Variant
            For Each Folder In CreateObject("Scripting.FileSystemObject").GetFolder(gl_game_path).SubFolders
                .AddItem Right(Folder, Len(Folder) - InStrRev(Folder, "\"))
            Next
        End With
        .Show
    End With
End Sub

Private Sub CommandButton_Profile_Delete_Click()
    gl_variable_read_func
    If gl_game_name = "" Or gl_profile_name = "" Then
        Exit Sub
    End If
    With UserForm_DeleteProfile
        .Label_Game.Caption = gl_game_name
        .Label_Profile.Caption = "Are you sure to Delete the Profile: " & gl_profile_name & " ?"
        .Show
    End With
End Sub

Private Sub CommandButton_Save_Save_Click()
    gl_variable_read_func
    If gl_game_name = "" Or gl_profile_name = "" Then
        Exit Sub
    End If
    For Each cell In gl_save_name_all_range.Cells
        If cell.Value = gl_save_name_new Then
            Exit Sub
        End If
    Next
    With CreateObject("Scripting.FileSystemObject")
        If gl_save_name_last <> "" And .FolderExists(gl_save_name_last_path) Then
            .DeleteFolder gl_save_name_last_path, True
        End If
        .CopyFolder gl_save_path, gl_save_path_new, True
    End With
    gl_profile_save_write_func
End Sub

Private Sub CommandButton_Save_Load_Click()
    gl_variable_read_func
    If gl_game_name = "" Or gl_profile_name = "" Then
        Exit Sub
    End If
    If gl_save_name_selected <> "" Then
        With CreateObject("Scripting.FileSystemObject")
            If .FolderExists(gl_save_path_selected) Then
                .CopyFolder gl_save_path_selected, gl_save_path, True
            End If
        End With
    End If
End Sub

Private Sub CommandButton_Save_Delete_Click()
    gl_variable_read_func
    If gl_game_name = "" Or gl_profile_name = "" Or gl_save_name_selected = "" Then
        Exit Sub
    End If
    With UserForm_DeleteSave
        .Label_Game.Caption = gl_game_name
        .Label_Save.Caption = "Are you sure to Delete the Save: " & gl_save_name_selected & " ?"
        .Show
    End With
    gl_profile_save_write_func
End Sub
