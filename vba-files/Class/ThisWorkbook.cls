VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    With Worksheets("SaveLoad")
        Set gl_game_name_range = .Range("C8")
        Set gl_profile_name_range = .Range("C9")
        Set gl_save_name_all_range = .Range("C12", "I21")
        Set gl_save_name_last_range = .Range("C21")
    End With
End Sub
