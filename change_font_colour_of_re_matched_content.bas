Attribute VB_Name = "highlighting"
Sub hl_char(a_cell, a_first_index, a_len)
    
    With a_cell.Characters(a_first_index, a_len).Font
            
        .Color = -1003520
        
    End With

End Sub

Sub hl_cell(a_cell)

    a_pat = get_pattern()

    Set re = New RegExp
        
        With re
            .Pattern = a_pat
            .IgnoreCase = True
            .Global = True
            'MultiLine - If True, patterns are matched across line breaks in multi-line strings
        End With
    
    Set matches = re.Execute(a_cell.Value)
    
    For Each m In matches
'        MsgBox (m)
'        MsgBox (m.FirstIndex)
'        MsgBox (Len(m))
        Call hl_char(a_cell, m.FirstIndex + 1, Len(m))
    Next m
    
End Sub

Function get_pattern()
        

        get_pattern = "([ap]-)*\d+[""'] *x *\d+[""']"  ' plain size, e.g. 11x11
        
        get_pattern = get_pattern & "|"
        get_pattern = get_pattern & "\[?\d+\]?[""']*(?: *\w+){1,2} *x(?: *\w+){0,2} *\[?\d+\]?[""']*(?: [^o]\w+)"

        get_pattern = get_pattern & "|"
        get_pattern = get_pattern & "size +\d+"  ' size 1 size 2 size 3
'
'        a_pat = a_pat & "|"
'        a_pat = a_pat & "spare[ \w]+cover"  ' spare cover

        get_pattern = get_pattern & "|"
        get_pattern = get_pattern & "ink.?black|wipe.?down|cut.?out|\bt\w*.shaped?|apply \w+ asset label|attach\W+jay\W+label|hardware\W+detached"  ' modification
        
        get_pattern = get_pattern & "|"
        get_pattern = get_pattern & "1257[-\w]+[fics]"  ' modification
    
End Function

Sub cells_blue()
    
    Range("D4").Font.Color = -1003520
    Range("D5").Font.Color = -1003520
    
End Sub
