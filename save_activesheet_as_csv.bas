Attribute VB_Name = "save_activesheet_as_csv"
Sub save_activesheet_as_csv(ws, a_filename)

    ws.Copy
    
    ActiveWorkbook.SaveAs Filename:=a_filename, FileFormat:=xlCSV ', CreateBackup:=True
    
    ActiveWorkbook.Close

End Sub
