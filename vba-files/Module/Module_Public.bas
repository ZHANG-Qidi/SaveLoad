Attribute VB_Name = "Module_Public"
Option Explicit
Public gl_game_name_range, gl_profile_name_range, gl_save_name_all_range As Object
Public gl_game_name, gl_game_path, gl_profile_name, gl_profile_path, gl_save_path As String
Public gl_save_name_selected, gl_save_path_selected, gl_save_name_new, gl_save_path_new As String
Public gl_save_name_last_range As Object
Public gl_save_name_last, gl_save_name_last_path As String
Public gl_save_name_selected_range As Object

Private Sub OptionButtonSlot_range_read_func()
    Dim i As Integer
    For i = 1 To 10 Step 1
        With Worksheets("SaveLoad").OLEObjects("OptionButtonSlot" & i).Object
            If .Value = True Then
                Set gl_save_name_selected_range = Worksheets("SaveLoad").Range("C" & (i + 11))
                Exit For
            End If
        End With
    Next i
End Sub

Public Sub gl_variable_read_func()
    gl_game_name = gl_game_name_range.Value
    gl_game_path = ThisWorkbook.Path & "\SaveLoad" & "\" & gl_game_name
    gl_profile_name = gl_profile_name_range.Value
    gl_profile_path = gl_game_path & "\" & gl_profile_name
    With CreateObject("Scripting.FileSystemObject")
        If .FileExists(gl_game_path & "\Path.txt") Then
            With .GetFile(gl_game_path & "\Path.txt").OpenAsTextStream(1, -1)
                gl_save_path = .ReadLine
                .Close
            End With
        End If
    End With
    OptionButtonSlot_range_read_func
    gl_save_name_selected = gl_save_name_selected_range.Value
    gl_save_path_selected = gl_profile_path & "\" & gl_save_name_selected
    gl_save_name_new = gl_game_name & "." & Format(Now, "yyyy-mm-dd-hh-mm-ss") & ".bak"
    gl_save_path_new = gl_profile_path & "\" & gl_save_name_new
    gl_save_name_last = gl_save_name_last_range.Value
    gl_save_name_last_path = gl_profile_path & "\" & gl_save_name_last
End Sub

Public Sub gl_profile_save_write_func()
    gl_save_name_all_range.ClearContents
    With CreateObject("Scripting.FileSystemObject").GetFolder(gl_profile_path)
        Dim Folder As Object
        Dim i As Integer
        i = 12
        For Each Folder In .SubFolders
            With Worksheets("SaveLoad")
                .Range("C" & i).Value = Right(Folder, Len(Folder) - InStrRev(Folder, "\"))
            End With
            i = i + 1
        Next
    End With
    gl_save_sort_func
End Sub

Public Sub gl_save_sort_func()
  With ActiveSheet
    .Sort.SortFields.Clear
    .Sort.SortFields.Add Key:=.Range("C12"), Order:=xlDescending
    .Sort.SetRange .Range("C12:C21")
    .Sort.Header = xlNo
    .Sort.SortMethod = xlPinYin
    .Sort.Apply
  End With
End Sub
