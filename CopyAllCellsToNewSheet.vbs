'This subroutine will copy all cells from an existing sheet to a new sheet. Cell formatting is preserved, too!
Sub CopyAllCellsToNewSheet()
    
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim sourceName As String
    Dim targetName As String

    'Int variable to tell Excel where to place new worksheet
    Dim targetLocation

    Set sourceName = "Change this to the name of the source worksheet"
    Set targetName = "Change this to the name of the sheet you want to create"

    Set targetLocation = 1

    Set wsSource = Worksheets(sourceName)
    Set wsTarget = Worksheets.Add(After:=Worksheets(targetLocation))
    wsTarget.Name = targetName
    
    wsSource.Cells.Copy Destination:=wsTarget.Cells
    
    
End Sub