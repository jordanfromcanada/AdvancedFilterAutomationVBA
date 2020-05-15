Sub filter_input_salaries()

'Uses a dynamic criteria range for Data>Filter>Advanced
'To automate multiple Advanced Filter actions with VBA for display purposes
'e.g. generating PDF statements from Excel files...

Dim i, j, c As Integer
c = 0

'Loop over columns
For i = 1 To 50
    'Identify the 4-digit code associated with the project
    code_row = ActiveSheet.Range("A:A").Find(what:="ProjectID", after:=Range("A1")).Row + 1
    If Len(Cells(code_row, i)) = 4 Then
    
        'Loop over rows
        For j = 5 To 50
            'Find the Advanced Filter criteria range (begins with cell 'Name' for each project)
            If InStr(UCase(Cells(j, i)), "NAME") > 0 Then
            
                'Print to Immediate Window the location of each project's criteria cells
                'where the Advanced Filter will transpose data
                Debug.Print "Found Name at Row " & j & " Column " & i
                Debug.Print "Inserting salaries for Project " & Cells(code_row, i) & "..."
                
                'Color the criteria cells to indicate data is being inserted below them
                Range(Cells(j, i), Cells(j, i + 2)).Interior.ColorIndex = 44
                
                'Advanced Filter from the table in employee-data
                'and a dynamic criteria range (the first project's criteria range is A10:C10,
                'second project's criteria range is D10:F10...
                Sheets("employee-data").Range("employees[#All]").AdvancedFilter Action:= _
                    xlFilterCopy, CriteriaRange:=Range(Cells(2, i), Cells(3, i)), _
                    CopyToRange:=Range(Cells(j, i), Cells(j, i + 2)), Unique:=False
                c = c + 1
                
            'If filter is performed for the project, exit the row for loop and keep iterating over columns
            j = 5
            GoTo flag1
            End If
        Next j
    End If
'Continue iterating over columns
flag1:
Next i

Debug.Print "Completed " & c & " Advanced Filter actions"

End Sub
