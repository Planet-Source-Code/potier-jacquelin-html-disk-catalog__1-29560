Attribute VB_Name = "module_restrictions"
Public known_files_restrictions_array()
Public Sub fill_restrictions_array(ByVal min_kb_file_size As Long, ByVal max_kb_file_size As Long, ByVal all_and_words As String, ByVal all_or_words As String, ByVal all_not_words As String, ByVal all_extensions As String)
    On Error Resume Next
    Dim tmparray                As Variant
    Dim cpt                     As Integer
    Dim size                    As Integer
    Dim pos                     As Integer
    
    ReDim known_files_restrictions_array(2)
    known_files_restrictions_array(0) = min_kb_file_size 'in kb
    known_files_restrictions_array(1) = min_kb_file_size 'in kb

    pos = 2
    If all_and_words <> "" Then
        tmparray = Split(all_and_words, ";")
        size = UBound(tmparray)
        known_files_restrictions_array(pos) = size + 1
        pos = pos + 1
        ReDim Preserve known_files_restrictions_array(pos + 1)
        For cpt = 0 To size
            known_files_restrictions_array(pos) = tmparray(cpt)
            pos = pos + 1
            ReDim Preserve known_files_restrictions_array(pos)
        Next cpt
    Else
        known_files_restrictions_array(pos) = 0
        pos = pos + 1
        ReDim Preserve known_files_restrictions_array(pos)
    End If
    
    If all_or_words <> "" Then
        tmparray = Split(all_or_words, ";")
        size = UBound(tmparray)
        known_files_restrictions_array(pos) = size + 1
        pos = pos + 1
        ReDim Preserve known_files_restrictions_array(pos)
        For cpt = 0 To size
            known_files_restrictions_array(pos) = tmparray(cpt)
            pos = pos + 1
            ReDim Preserve known_files_restrictions_array(pos)
        Next cpt
    Else
        known_files_restrictions_array(pos) = 0
        pos = pos + 1
        ReDim Preserve known_files_restrictions_array(pos)
    End If
    
    If all_not_words <> "" Then
        tmparray = Split(all_not_words, ";")
        size = UBound(tmparray)
        known_files_restrictions_array(pos) = size + 1
        pos = pos + 1
        ReDim Preserve known_files_restrictions_array(pos)
        For cpt = 0 To size
            known_files_restrictions_array(pos) = tmparray(cpt)
            pos = pos + 1
            ReDim Preserve known_files_restrictions_array(pos)
        Next cpt
    Else
        known_files_restrictions_array(pos) = 0
        pos = pos + 1
        ReDim Preserve known_files_restrictions_array(pos)
    End If
    
    If all_extensions <> "" Then
        tmparray = Split(all_extensions, ";")
        size = UBound(tmparray)
        known_files_restrictions_array(pos) = size + 1
        pos = pos + 1
        ReDim Preserve known_files_restrictions_array(pos)
        For cpt = 0 To size
            known_files_restrictions_array(pos) = tmparray(cpt)
            pos = pos + 1
            ReDim Preserve known_files_restrictions_array(pos)
        Next cpt
    Else
        MsgBox "You must choose file extension(s)", vbExclamation, "HTML Disk Catalog"
        known_files_restrictions_array(pos) = 0
        pos = pos + 1
        ReDim Preserve known_files_restrictions_array(pos)
    End If
End Sub



Public Function check_knownfiles_restrictions(ByVal file_name As String, ByVal file_size As Single) As Boolean
    On Error Resume Next
    Dim cpt         As Integer
    Dim word_ok     As Boolean
    Dim pos         As Integer
    Dim begin       As Integer
    
    'check file size
        If known_files_restrictions_array(0) > file_size Then Exit Function 'min
        If known_files_restrictions_array(1) < file_size And known_files_restrictions_array(1) > 0 Then Exit Function 'max
        begin = 2
    'check and
        For cpt = 1 To known_files_restrictions_array(begin) 'known_files_restrictions_array(begin) contain the number of and words
            pos = InStr(1, file_name, known_files_restrictions_array(cpt + begin))
            If pos < 1 Then Exit Function
        Next cpt
        begin = begin + known_files_restrictions_array(begin) + 1
    'check or
        For cpt = 1 To known_files_restrictions_array(begin)
            pos = InStr(1, file_name, known_files_restrictions_array(cpt + begin))
            If pos > 0 Then word_ok = True
        Next cpt
        If known_files_restrictions_array(begin) > 0 And Not word_ok Then Exit Function
        begin = begin + known_files_restrictions_array(begin) + 1
    'check not
        For cpt = 1 To known_files_restrictions_array(begin)
            pos = InStr(1, file_name, known_files_restrictions_array(cpt + begin))
            If pos > 0 Then Exit Function
        Next cpt
        begin = begin + known_files_restrictions_array(begin) + 1
    'check extension
        For cpt = 1 To known_files_restrictions_array(begin)
            pos = InStr(1, get_extension(file_name), known_files_restrictions_array(cpt + begin))
            If pos > 0 Then word_ok = True
        Next cpt
        If known_files_restrictions_array(begin) > 0 And Not word_ok Then Exit Function
        
        check_knownfiles_restrictions = True
        
        
End Function
