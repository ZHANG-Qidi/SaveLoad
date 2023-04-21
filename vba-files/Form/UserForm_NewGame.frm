VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_NewGame 
   Caption         =   "NewGame"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "UserForm_NewGame.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_NewGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton_Cancel_Click()
    Unload Me
End Sub

Private Sub CommandButton_OK_Click()
    gl_game_name = TextBox_Name.Value
    gl_save_path = TextBox_Path.Value
    gl_game_path = ThisWorkbook.Path & "\SaveLoad\" & gl_game_name
    If gl_game_name = "" Then
        MsgBox "Please Input Game Name"
        Exit Sub
    ElseIf gl_save_path = "" Then
        MsgBox "Please Input Game Path"
        Exit Sub
    End If
    With CreateObject("Scripting.FileSystemObject")
        If .FolderExists(gl_game_path) = False Then
            .CreateFolder gl_game_path
        End If
        With .CreateTextFile(gl_game_path & "\Path.txt", True, True)
            .WriteLine (gl_save_path)
            .Close
        End With
    End With
    gl_game_name_range.Value = gl_game_name
    gl_profile_name_range.MergeArea.ClearContents
    gl_save_name_all_range.ClearContents
    Unload Me
End Sub

Private Sub CommandButton_Path_Click()
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = "C:\Users\" & Environ("USERNAME") & "\AppData\"
        .AllowMultiSelect = False
        .Title = "Save Folder Select"
        If .Show = True Then
            Me.TextBox_Path.Text = .SelectedItems(1)
            If Me.TextBox_Name.Text = "" Then
                With Me.TextBox_Path
                    Me.TextBox_Name = Right(.Text, Len(.Text) - InStrRev(.Text, "\"))
                End With
            End If
        End If
    End With
End Sub

Private Sub UserForm_Click()

End Sub
