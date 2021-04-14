Attribute VB_Name = "idk"
Sub transform(tbl)

' range E:I is targeted range
' col send? is at the right of column I

Range("a1:d1").EntireColumn.Delete

Set col_to_loop = tbl.ListColumns("send?")

For i = tbl.DataBodyRange.Rows.Count To 1 Step -1
    
    If col_to_loop.DataBodyRange(i) <> 1 Then
        
        Rows(i + 1).EntireRow.Delete
    
    End If

Next i

Range("f1:xfd1").EntireColumn.Delete

End Sub

Sub save_table_as_csv(tbl, a_filename)

    tbl.Range.Copy
    
    ActiveWorkbook.SaveAs Filename:=a_filename, FileFormat:=xlCSV ', CreateBackup:=True
    
    'ActiveWorkbook.Close

End Sub

Sub main()

Set wb = ActiveWorkbook
Set ws = ActiveSheet

Set tbl = ws.ListObjects("Table_medifab_nz")

new_filename = "This is a meaningful filename.csv"

Call save_table_as_csv(tbl, new_filename)

Set wb = ActiveWorkbook
Set ws = ActiveSheet

Set tbl = ws.ListObjects("Table_medifab_nz")

Call transform(tbl)


End Sub
