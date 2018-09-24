Public Sub CloseTables()

    Dim tbls As AllTables
    Dim tbl As Variant
    
    Set tbls = Access.Application.CurrentData.AllTables
    For Each tbl In tbls
       If tbl.IsLoaded Then
          If vbYes = MsgBox("Close " & tbl.Name & "?") Then
             DoCmd.Close acTable, tbl.Name, acSavePrompt
          End If
       End If
    Next

End Sub
