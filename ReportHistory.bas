Function ReportHistory(sRpt As String) As String
'Get all report history (properties)

  Dim acObj As AccessObject
  Dim sRptIn As String
  Dim sDatePrinted As String
  Dim sDateModified As String
  Dim sDateCreated As String
  Dim sPort As String
  Dim sDevice As String
  Dim sBuild As String
  
  sRptIn = sRpt
  sPort = "Port name: " & Application.Printers(0).Port
  sDevice = "Device name: " & Application.Printers(0).DeviceName
  sBuild = ""
  
  For Each acObj In CurrentProject.AllReports
    With acObj

      If acObj.Name = sRptIn Then
        sDatePrinted = "Date printed: " & Now()
        sDateModified = "Date modified: " & .DateModified
        sDateCreated = "Date created: " & .DateCreated
        Exit For
      End If

    End With
  Next acObj

  sBuild = sDatePrinted & ", " & sPort & ", " & sDevice & ", " & _
           sDateCreated & ", " & sDateModified & "."

  ReportHistory = sBuild

End Function
