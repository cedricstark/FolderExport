Sub FolderExport()
  Dim Rng As Range
  Dim maxRows, maxCols, r As Integer
  Dim workOrder As String
  Dim orderDir As String
  Set Rng = Selection
  maxRows = Rng.Rows.Count
  maxCols = Rng.Columns.Count
    r = 1
    Do While r <= maxRows
      If Len(Dir(ActiveWorkbook.Path & "\" & Rng(r, 1), vbDirectory)) = 0 Then
        orderDir = "C:\Users\name\Documents\Open Jobs" & "\" & "[" & Rng(r, 1) & "]" & " " & Rng(r, 2) & " " & Rng(r, 3) & " " & Rng(r, 4) & " " & Rng(r, 5) 'Line declares the location and name of the new folder.
        MkDir (orderDir)'Don't be a fucking dumb ass this needs changed for the server
        On Error Resume Next
      End If
      workOrder = Rng(r,1)
      Call populateOrder(workOrder,orderDir)
      Call checkList(orderDir, workOrder)
      r = r + 1
    Loop
End Sub

Sub populateOrder(wOrder,oDir)
  Dim fso As Object
  Dim sourceFile As String
  Dim targetFile As String
  Dim answer As Integer
  Dim pathFrom as String
  pathFrom = "C:\Users\name\Documents\Work Orders" 'Don't be a fucking dumb ass this needs changed for the server

  Set fso = CreateObject("Scripting.FileSystemObject")
  sourceFile = pathFrom & "\" & wOrder & ".pdf"
  targetFile = oDir & "\" & wOrder & ".pdf"

  If fso.FileExists(targetFile) Then
      answer = MsgBox("File already exists in this location. " _
        & "Are you sure you want to continue? If you continue " _
        & "the file at destination will be deleted!", _
        vbInformation + vbYesNo)
      If answer = vbNo Then
        Exit Sub
      End If
      Kill targetFile
  End If
  fso.MoveFile sourceFile, targetFile
  Set fso = Nothing
End Sub

Sub checkList(oDir, wOrder)
  Dim sourceFile As String
  Dim destinationFile As String
  destinationFile = oDir & "\" & wOrder & "-checklist.pdf"
  sourceFile = "C:\Users\name\Documents\hidden\checklist.pdf"'Don't be a fucking dumb ass this needs changed for the server

  FileCopy sourceFile, destinationFile
End Sub
