Attribute VB_Name = "colour_code_weekday"

Function colour_code_weekday(a_date)

' This function returns string to be used in defining footter colous

    my_weekday = Weekday(a_date, vbMonday)

    colour_code_weekday = Choose( _
        my_weekday, "&K000000", "&K00FF00", "&KFFC0CB", "&KFFA500", "&K800080", "&KFFFF00" _
    )

End Function
