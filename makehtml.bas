Attribute VB_Name = "makehtml"
Public Sub make_html(filename As String, preview As Boolean, Optional absolute_path As Boolean = False)
    On Error Resume Next
    Dim output_path As String
    output_path = LCase$(get_folder_name(filename))
    
    Dim buffer As String
    'make body css style
    If Right$(output_path, 1) <> "\" Then output_path = output_path & "\"
    buffer = "<html>" & vbCrLf & "<head>" & vbCrLf & "<title>" & style_title.title _
            & "</title>" & vbCrLf & "<style type=""text/css"">" & vbCrLf & _
            "<!--" & vbCrLf & "body{"
    
    If style_background.attachement <> "" Then
        buffer = buffer & "background-attachment:" & style_background.attachement & " ;"
    End If
    If style_background.use_color Then
        buffer = buffer & " background-color: #" & long_to_html(style_background.color) & " ; "
    End If
    If style_background.use_picture Then
        buffer = buffer & "background-image: "
        If preview Or absolute_path Then
            buffer = buffer & "url(""" & style_background.picture_name & """)" & "; "
        Else
            If is_file_existing(style_background.picture_name) Then
                FileCopy style_background.picture_name, output_path & get_file_name(style_background.picture_name)
            End If
            buffer = buffer & "url(""" & get_file_name(style_background.picture_name) & """)" & "; "
        End If
        
        If style_background.repeat <> "" Then
            buffer = buffer & "background-repeat: " & style_background.repeat & "; "
        End If
        If style_background.hpos <> "" Or style_background.vpos <> "" Then
            buffer = buffer & "background-position: "
            If style_background.hpos <> "left" And style_background.hpos <> "center" And style_background.hpos <> "right" Then
                buffer = buffer & val(style_background.hpos)
                If style_background.hpostype <> "%" Then
                    buffer = buffer & "px "
                Else
                    buffer = buffer & "% "
                End If
            Else
                buffer = buffer & style_background.hpos & " "
            End If
            
            If style_background.vpos <> "top" And style_background.vpos <> "center" And style_background.vpos <> "bottom" Then
                buffer = buffer & val(style_background.vpos)
                If style_background.vpostype <> "%" Then
                    buffer = buffer & "px;"
                Else
                    buffer = buffer & "%;"
                End If
            Else
                buffer = buffer & style_background.vpos & ";"
            End If
        End If
    End If
    buffer = buffer & "}" & vbCrLf
    'end of body css
    
    'css for title
    buffer = buffer & ".title_style{ "
    If style_title.style.size <> 0 Then
        buffer = buffer & "font-size: " & style_title.style.size & "px" & ";"
    End If
    buffer = buffer & "color: #" & long_to_html(style_title.style.color) & " ;"
    If style_title.style.bold Then buffer = buffer & "font-weight: bold; "
    If style_title.style.italic Then buffer = buffer & "font-style: italic; "
    If style_title.style.underline Or style_title.style.line_through Then
        buffer = buffer & "text-decoration: "
        If style_title.style.underline Then buffer = buffer & "underline "
        If style_title.style.line_through Then buffer = buffer & "line-through"
        buffer = buffer & "; "
    End If
    buffer = buffer & "}" & vbCrLf
    
    'css for folder_names
    buffer = buffer & ".foldername_style{ "
    If style_folder.size <> 0 Then
        buffer = buffer & "font-size: " & style_folder.size & "px" & ";"
    End If
    buffer = buffer & "color: #" & long_to_html(style_folder.color) & " ;"
    If style_folder.bold Then buffer = buffer & "font-weight: bold; "
    If style_folder.italic Then buffer = buffer & "font-style: italic; "
    If style_folder.underline Or style_folder.line_through Then
        buffer = buffer & "text-decoration: "
        If style_folder.underline Then buffer = buffer & "underline "
        If style_folder.line_through Then buffer = buffer & "line-through"
        buffer = buffer & "; "
    End If
    buffer = buffer & "}" & vbCrLf
    
    'css for visited style_visited
    buffer = buffer & "a:visited { "
    If style_visited.size <> 0 Then
        buffer = buffer & "font-size: " & style_visited.size & "px" & ";"
    End If
    buffer = buffer & "color: #" & long_to_html(style_visited.color) & " ;"
    If style_visited.bold Then buffer = buffer & "font-weight: bold; "
    If style_visited.italic Then buffer = buffer & "font-style: italic; "
    If style_visited.none Then
        buffer = buffer & "text-decoration: none; "
    Else
        If style_visited.underline Or style_visited.line_through Then
            buffer = buffer & "text-decoration: "
            If style_visited.underline Then buffer = buffer & "underline "
            If style_visited.line_through Then buffer = buffer & "line-through"
            buffer = buffer & "; "
        End If
    End If
    buffer = buffer & "}" & vbCrLf
    
    'css for link
    buffer = buffer & "a:link { "
    If style_link.size <> 0 Then
        buffer = buffer & "font-size: " & style_link.size & "px" & ";"
    End If
    buffer = buffer & "color: #" & long_to_html(style_link.color) & " ;"
    If style_link.bold Then buffer = buffer & "font-weight: bold; "
    If style_link.italic Then buffer = buffer & "font-style: italic; "
    If style_link.none Then
        buffer = buffer & "text-decoration: none; "
    Else
        If style_link.underline Or style_link.line_through Then
            buffer = buffer & "text-decoration: "
            If style_link.underline Then buffer = buffer & "underline "
            If style_link.line_through Then buffer = buffer & "line-through"
            buffer = buffer & "; "
        End If
    End If
    buffer = buffer & "}" & vbCrLf
    
    'css for hover
     buffer = buffer & "a:hover { "
    If style_hover.size <> 0 Then
        buffer = buffer & "font-size: " & style_hover.size & "px" & ";"
    End If
    buffer = buffer & "color: #" & long_to_html(style_hover.color) & " ;"
    If style_hover.bold Then buffer = buffer & "font-weight: bold; "
    If style_hover.italic Then buffer = buffer & "font-style: italic; "
    If style_hover.none Then
        buffer = buffer & "text-decoration: none; "
    Else
        If style_hover.underline Or style_hover.line_through Then
            buffer = buffer & "text-decoration: "
            If style_hover.underline Then buffer = buffer & "underline "
            If style_hover.line_through Then buffer = buffer & "line-through"
            buffer = buffer & "; "
        End If
    End If
    buffer = buffer & "}" & vbCrLf
    
    'closing css and http header
    buffer = buffer & "-->" & vbCrLf & "</style>" & vbCrLf & "</head>" & "<body>" & vbCrLf
    'insert title
    buffer = buffer & "<p align=""center"" class=title_style>" & style_title.title & "</p>" & vbCrLf
    'generating array
    buffer = buffer & "<table "
    If style_table.use_bgcolor Then buffer = buffer & "bgcolor=""#" & long_to_html(style_table.bgcolor) & """ "
    If style_table.use_bordercolor Then buffer = buffer & "bordercolor=""#" & long_to_html(style_table.bordercolor) & """ "
    If style_table.bordersize <> 0 Then buffer = buffer & "border=""" & style_table.bordersize & """ "
    If style_table.bgpicture <> "" Then
        If preview Or absolute_path Then
            buffer = buffer & "background=""" & style_table.bgpicture & """ "
        Else
            If is_file_existing(style_table.bgpicture) Then
                FileCopy style_table.bgpicture, output_path & get_file_name(style_table.bgpicture)
            End If
            buffer = buffer & "background=""" & get_file_name(style_table.bgpicture) & """ "
        End If
    End If



    If style_table.height <> 0 Then
        buffer = buffer & "height=""" & style_table.height
        If style_table.heighttype = "%" Then
            buffer = buffer & "%"" "
        Else
            buffer = buffer & """ "
        End If
    End If
    If style_table.width <> 0 Then
        buffer = buffer & "width=""" & style_table.width
        If style_table.widthtype = "%" Then
            buffer = buffer & "%"" "
        Else
            buffer = buffer & """ "
        End If
    End If
    If style_table.align <> "Default" And style_table.align <> "" Then buffer = buffer & "align=""" & style_table.align & """ "
    If style_table.cellpad <> 0 Then buffer = buffer & "cellpadding=""" & style_table.cellpad & """ "
    If style_table.cellspace <> 0 Then buffer = buffer & "cellspacing=""" & style_table.cellspace & """ "
    
    'adding folders and file name
    Dim cpt As Long
    Dim tmplast_folder  As String
    Dim first_folder As Boolean
    first_folder = True
    tmplast_folder = ""
    For cpt = 0 To UBound(files_found) - 1 'we can do it because array is sorted
        'check if it's a new folder
        If tmplast_folder <> files_found(cpt).folder Then
            If Not first_folder Then
                'close table of files
                buffer = buffer & "</table>" & vbCrLf & "</td>" & vbCrLf & "</tr>"
            End If
            first_folder = False
            tmplast_folder = files_found(cpt).folder
            'write folder name
            If absolute_path Then
                buffer = buffer & "<tr><td valign=""middle"" align=""center"" class=foldername_style>" _
                                & files_found(cpt).folder & "</td>"
            Else
                buffer = buffer & "<tr><td valign=""middle"" align=""center"" class=foldername_style>" _
                                & remove_drive_letter(files_found(cpt).folder) & "</td>"
            End If
            'write file name
            buffer = buffer & "<td> <table border=""0"">" & vbCrLf & " <tr>" & vbCrLf & "<td><a href=""" _
                            & get_final_path(files_found(cpt).full_path, absolute_path, output_path) _
                            & """>" & files_found(cpt).file_name & "</a>" & vbCrLf & "</td>" & "</tr>" & vbCrLf
        Else
            buffer = buffer & "<tr>" & vbCrLf & "<td><a href=""" _
                            & get_final_path(files_found(cpt).full_path, absolute_path, output_path) _
                            & """>" & files_found(cpt).file_name & "</a>" & vbCrLf & "</td>" & "</tr>" & vbCrLf
        End If
    Next cpt
    'close last table of files
    If Not first_folder Then
        buffer = buffer & "</table>" & vbCrLf & "</td>" & vbCrLf & "</tr>"
    End If
    
    'closing table
    buffer = buffer & "</table>"
    'closing the http doc
    buffer = buffer & "</body>" & vbCrLf & "</html>"
    
    If is_file_existing(filename) Then Kill filename
    Dim numfile As Integer
    num_file = FreeFile
    Open filename For Append Access Write As num_file
        Print #num_file, buffer
    Close num_file
End Sub


Private Function get_final_path(full_path As String, absolute_path As Boolean, relative_root_path As String)
    On Error Resume Next
    Dim tmp         As Integer
    Dim tmparray1   As Variant
    Dim tmparray2   As Variant
    Dim size1       As Integer
    Dim size2       As Integer
    Dim cpt         As Integer
    Dim pos         As Integer
    
    If absolute_path Then
        get_final_path = full_path
    Else
        tmp = InStr(1, full_path, relative_root_path)
        If tmp > 0 Then
            'full_path is a subfolder of relative_root_path
            get_final_path = Mid$(full_path, tmp + Len(relative_root_path))
        Else
            tmparray1 = Split(full_path, "\")
            tmparray2 = Split(relative_root_path, "\")
            size1 = UBound(tmparray1) - 1
            Min = size1
            size2 = UBound(tmparray2) - 1
            If size2 < Min Then Min = size2
            pos = 0
            For cpt = 0 To Min
                If tmparray1(cpt) = tmparray2(cpt) Then
                    pos = cpt
                Else
                    Exit For
                End If
            Next cpt
            get_final_path = ""
            For cpt = pos + 1 To size2 'relative_root_path
                get_final_path = get_final_path & "../"
            Next cpt
        
            For cpt = pos + 1 To size1 'full_path
                get_final_path = get_final_path & tmparray1(cpt) & "/" 'add all folder names
            Next cpt
            'add file name
            get_final_path = get_final_path & tmparray1(size1 + 1)
        End If
    End If
    get_final_path = replace_char(get_final_path, "\", "/")
End Function

Private Function replace_char(ByVal str As String, replaced As String, replacing As String) As String
    On Error Resume Next
    Dim tmp As Long
    Dim strlen  As Long
    strlen = Len(str)
    Dim strtmp  As String
    tmp = InStr(tmp + 1, str, replaced)
    Do While tmp > 0
        Select Case tmp
            Case 1
                str = replacing & Mid$(str, 2)
            Case strlen
                str = Mid$(str, 1, strlen - 1) & replacing
            Case Else
                str = Mid$(str, 1, tmp - 1) & replacing & Mid$(str, tmp + 1)
        End Select
        tmp = InStr(tmp + 1, str, replaced)
    Loop
    replace_char = str
End Function

Private Function remove_drive_letter(full_path As String) As String
    On Error Resume Next
    Dim tmp As Integer
    tmp = InStr(1, full_path, ":")
    If tmp > 0 Then
        remove_drive_letter = Mid$(full_path, tmp + 1)
    Else
        remove_drive_letter = full_path
    End If
End Function
