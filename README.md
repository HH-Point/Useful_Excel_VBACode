### Just some useful VBA code I've used in Excel

```
Private Sub Worksheet_Activate()
'This will auto create and auto update a table of centents for the sheets in the workbook and hyperlink to them'
Dim sht As Worksheet
Dim TOCsht As Worksheet
Dim RowNo As Integer

Set TOCsht = Sheet1


With TOCsht.Cells(1, 1)
    .Value = "Sheet Link"
    .Font.Bold = True
    .Font.Size = 11
    .Font.Color = vbWhite
    .Interior.Color = RGB(68, 114, 196)
End With

RowNo = 1

For Each sht In ThisWorkbook.Worksheets
    If sht.CodeName <> "Sheet1" Then
        RowNo = RowNo + 1
        TOCsht.Cells(RowNo, 1).Hyperlinks.Add _
        Anchor:=Cells(RowNo, 1), _
        Address:="", SubAddress:="'" & sht.Name & "'!A1", _
        ScreenTip:="", _
        TextToDisplay:=sht.Name
    End If
Next sht

Columns.AutoFit
End Sub
```

```
Sub delete_rows()
    'This will filter a section of a table to a criteria, select the data, delete the data, and remove the filter'
    Dim lo As ListObject
    
        Set lo = Sheet1.ListObjects(1)
        
        lo.Range.AutoFilter Field:=2, Criteria1:="Item"
        
        Application.DisplayAlerts = False
            lo.DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
        Application.DisplayAlerts = True
        
        lo.AutoFilter.ShowAllData
            
End Sub
```

```
Function SheetCheck(sheet_name As String) As Boolean

            Dim ws As Worksheet

                SheetCheck = False
 
            For Each ws In ThisWorkbook.Worksheets
 
                If ws.Name = sheet_name Then
    
                    SheetCheck = True
        
                End If
 
            Next
 
End Function

Sub update_sheets()
'This will pull sheet names from the updated sharepoint query list and create a sheet for anything not existing'
    Dim lo As ListObject
    Dim sheet_name As String
    
        Set lo = Sheet1.ListObjects(1)
        
        If SheetCheck(sheet_name) = False And sheet_name <> "" Then
        
        Sheets.Add.Name = lo.Name
        
        End If

End Sub
```
```
Option Explicit 
 
Sub AutoFitAll() 
'Autofit all worksheets'     
    Application.ScreenUpdating = False 
    Dim wkSt As String 
    Dim wkBk As Worksheet 
    wkSt = ActiveSheet.Name 
    For Each wkBk In ActiveWorkbook.Worksheets 
        On Error Resume Next 
        wkBk.Activate 
        Cells.EntireColumn.AutoFit 
    Next wkBk 
    Sheets(wkSt).Select 
    Application.ScreenUpdating = True 
     
End Sub 
```
