'This subroutine is for updating variable names on a duplicated sheet.
'
'Context: I had a large sheet (~800 rows) with around 200 workbook-scope 
'variables. I wanted to duplicate this sheet and modify it slightly for a
'new model feature.
'When I tried to use the Excel "duplicate sheet" feature, it copied all the 
'cells, but it copied the named variables from the old sheet with the exact
'same name, but worksheet (instead of workbook) scope.
'Since VBA does not allow setting scope of an existing variable, I created
'this subroutine instead.
'The subroutine will also iterate through all copied formulas on the new
'sheet and update the named variables (since the formula was copied from the
'old sheet with different var names) to their corresponding versions on the 
'new sheet

Sub UpdateDuplicatedNames()
    
    'Iterator variables for going through all workbook names
    Dim nm As Name
    Dim slot As Name
    
    'For storing the old variable name (from original sheet) and new variable 
    'name (updated, in my case, with a replaced substring)
    Dim oldName As String
    Dim newName As String
    
    'For storing variable's cell reference on old sheet and the desired new 
    'reference on the target sheet
    Dim oldRef As String
    Dim newRef As String
    
    'Intermediate variables for adding newly created name to workbook names 
    'directory
    Dim cell As Range
    Dim cellName As String
    
    'Counter variable (not necessary) just to see how many variables were 
    'created/updated
    Dim ctr
    
    'Used for searching through sheet's cells to find cells containing formulas
    Dim xRg As Range
    Dim xCell As Range
    
    'Variable to hold the formula (and update the appropriate substring)
    Dim frmla As String
    
    'Stores the first formula found containing desired name, used to 
    'determine when Excel has seen all formulas on the sheet
    Dim firstAddress
    
    'For writing to output file
    Dim str1 As String
    Dim str2 As String
    
    'Set up writing to an output file (optional)
    Dim n As Integer
    n = FreeFile()
    Open "complete_file_path_here (ex. C:\Temp\output.txt" For Output As #n

    Dim originSheet As String
    '-----------'

    'SETUP

    'Change this string to the original sheet's name in this 
    'format: ='Refueling Station - Gaseous H2'
    Set originSheet = "='Sample Sheet - Type 1'"
    
    'Iterate through all names in the workbook
    For Each nm In ThisWorkbook.Names
        
        'Check if the name refers to something on the desired sheet
        If (StrComp(Split(nm.RefersTo, "!")(0), 
            originSheet, vbTextCompare) = 0) Then
            'Increment counter (just for curiosity to see how many variables there are)
            ctr = ctr + 1

            'Set oldName equal to the current name
            oldName = nm.Name
            'Temporarily set newName equal to the oldName
            newName = oldName
            
            'Rules for updating names for the new sheet
            'In my case, I wanted to replace a certain substring "gas" with MH
            If (InStr(1, oldName, "gas", vbTextCompare) <> 0) Then
                newName = Replace(oldName, "gas", "mh")
            ElseIf (InStr(1, oldName, "forecourt", vbTextCompare) <> 0) Then
                newName = Replace(oldName, "forecourt", "forecourt_MH")
            ElseIf (InStr(1, oldName, "GH2_", vbTextCompare) <> 0) Then
                newName = Replace(oldName, "GH2_", "MH_")
            ElseIf (InStr(1, oldName, "GH2", vbTextCompare) <> 0) Then
                newName = Replace(oldName, "GH2", "MH_")
            ElseIf (InStr(1, oldName, "ForeC", vbBinaryCompare) <> 0) Then
                newName = Replace(oldName, "ForeC", "ForeC_MH")
            ElseIf (InStr(1, oldName, "Forec", vbBinaryCompare) <> 0) Then
                newName = Replace(oldName, "Forec", "Forec_MH")
            Else
                newName = oldName & "_MH"
            End If
            
            'Set the old cell reference
            oldRef = nm.RefersTo

            'Set new reference by changing the sheet name only 
            '(update this for your needs)
            newRef = Replace(oldRef, "Gaseous H2", "MH")
            
            'Extracts the cell name (ex. $A$7) by separating out the part after
            'exclamation mark
            cellName = Split(newRef, "!")(1)
            
            'Uses intermediate variable to assign appropriate range
            'In this case, "Refueling Station - MH" is the new sheet's name
            Set cell = Worksheets("Refueling Station - MH").Range(cellName)
            
            'Ensure that we are working with the right sheet 
            '(new sheet's name)
            Sheets("Refueling Station - MH").Select
            
            'Define range to search as all cells on MH sheet that 
            'contain formulas
            Set xRg = Sheets("Refueling Station - MH").Cells.SpecialCells(xlCellTypeFormulas)
            
            'Search for the first cell containing a formula referencing current oldName
            'Set xCell = xRg.Find(What:=oldName, LookIn:=xlFormulas)
            
            'Print to console the gaseous name and new MH name
            '(for debugging, uncomment below)
            'str1 = nm.Name
            'Debug.Print str1
            'Write #n, str1
            'str2 = newName
            'Debug.Print str2
            'Write #n, str2
            
            'Limit scope to cells on the new sheet containing formulas
            With Sheets("Refueling Station - MH").Cells.SpecialCells(xlCellTypeFormulas)
                'Search for all formula(s) on the sheet that contain the oldName
                Set xCell = .Find(What:=oldName, LookIn:=xlFormulas)
                If Not xCell Is Nothing Then
                    'Mark the location of first encountered formula
                    firstAddress = xCell.Address
                    Do
                        'Replace any variables in the formula
                        frmla = Replace(xCell.Formula, oldName, newName)
                        
                        '(for debugging, uncomment below)
                        'Debug.Print frmla
                        'Write #n, frmla

                        'Set formula to the adjusted formula
                        xCell.Formula = frmla

                        'Look for the next cell containing a formula referring to this oldName
                        Set xCell = xRg.FindNext(xCell)
                    If xCell Is Nothing Then
                        'Exit this routine once we're done
                        GoTo DoneFinding
                    End If
                    'Keep going until we come across the first place we found the oldName
                    Loop While xCell.Address <> firstAddress
                End If
DoneFinding:
            End With
        
            'Add the new name to the workbook directory
            ThisWorkbook.Names.Add Name:=newName, RefersTo:=cell
        End If
    Next nm
        'Not necessary, uncomment to see how many names were updated
        'Debug.Print ctr
End Sub

