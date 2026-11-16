Attribute VB_Name = "Module1"
Sub ExampleModule1()
    Dim FolderPath As String, FileName As String
    Dim WbSrc As Workbook, WbDst As Workbook
    Dim WsLoc As Worksheet, WsPallet As Worksheet, WsScanned As Worksheet, WsTarget As Worksheet
    Dim LastRowSrc As Long, LastRowDst As Long
    
    FolderPath = "C:\Put\The\Source\Path\"
    
    Set WbDst = Workbooks.Add
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Create the 3 target tabs
    Set WsLoc = WbDst.Sheets(1)
    WsLoc.Name = "Range_of_Locations_Detail"
    
    Set WsPallet = WbDst.Sheets.Add(After:=WsLoc)
    WsPallet.Name = "Pallet_Detail"
    
    Set WsScanned = WbDst.Sheets.Add(After:=WsPallet)
    WsScanned.Name = "Scanned_Detail"
    
    FileName = Dir(FolderPath & "*.xlsx")
    Do While FileName <> ""
        ' Route each file to its matching tab
        If InStr(1, FileName, "Cycle_Count_Range_of_Locations_Detail", vbTextCompare) > 0 Then
            Set WsTarget = WsLoc
        ElseIf InStr(1, FileName, "Cycle_Count_Pallet_Detail", vbTextCompare) > 0 Then
            Set WsTarget = WsPallet
        ElseIf InStr(1, FileName, "Cycle_Count_Scanned_Detail", vbTextCompare) > 0 Then
            Set WsTarget = WsScanned
        Else
            Set WsTarget = Nothing
        End If
        
        If Not WsTarget Is Nothing Then
            Set WbSrc = Workbooks.Open(FolderPath & FileName, ReadOnly:=True)
            With WbSrc.Sheets(1)
                LastRowSrc = .Cells(.Rows.Count, 1).End(xlUp).Row
                
                If WsTarget.Range("A1").Value = "" Then
                    ' First file: copy headers + all rows
                    .Rows("1:" & LastRowSrc).Copy Destination:=WsTarget.Range("A1")
                ElseIf LastRowSrc >= 2 Then
                    ' Additional files: skip header row and append data rows
                    LastRowDst = WsTarget.Cells(WsTarget.Rows.Count, 1).End(xlUp).Row
                    .Rows("2:" & LastRowSrc).Copy Destination:=WsTarget.Range("A" & (LastRowDst + 1))
                End If
            End With
            WbSrc.Close SaveChanges:=False
        End If
        
        FileName = Dir
    Loop
    
    WsLoc.Columns.AutoFit
    WsPallet.Columns.AutoFit
    WsScanned.Columns.AutoFit
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Done! All files combined into 3 tabs (Locations, Pallet, Scanned).", vbInformation
End Sub
