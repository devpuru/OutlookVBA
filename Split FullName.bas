' Split FullName to First, Middle, Last Name


Public Sub SplitNames()
    Dim currentExplorer As Explorer
    Dim Selection As Selection
    Dim obj As Object
    Dim cnt As Integer
    Dim Full As String
    
  
    Set currentExplorer = Application.ActiveExplorer
    Set Selection = currentExplorer.Selection

    On Error Resume Next

    For Each obj In Selection
        'Test for contact and not distribution list
        If obj.Class = olContact Then
            Set objContact = obj

     With objContact

        Full = obj.FullName
        Names = Split(Full)
        cnt = Len(Full) - Len(Replace(Full, " ", ""))

        'MsgBox obj.FullName & ": " & cnt
        
            If cnt = 2 Then
         '   MsgBox "F" & Names(0) & ",M" & Names(1) & ",L" & Names(2)
            obj.FirstName = Names(0)
            obj.MiddleName = Names(1)
            obj.LastName = Names(2)
            obj.Save
    
            ElseIf cnt = 1 Then
          '  MsgBox "F" & Names(0) & ",L" & Names(1)
            obj.FirstName = Names(0)
            obj.LastName = Names(1)
            obj.Save
        
            End If

'          If .FirstName <> "" Then
'          Let .User3 = .FirstName
'
'          If .LastName <> "" Then
'          Let .User4 = .LastName
'
'
'        .FirstName = .User4
'        .LastName = .User3
'        .Save
         
      'If you don't want to keep the values in the user fields for tracking purposes,
      ' uncomment these two lines. I recommend keeping the names in the user fields
      
       ' .User3 = ""
       ' .User4 = ""
       ' .Save
'        End If
'        End If
     End With
        End If

     Err.Clear
    Next

    Set obj = Nothing
    Set objContact = Nothing
End Sub

