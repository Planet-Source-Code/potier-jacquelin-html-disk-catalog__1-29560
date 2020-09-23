Attribute VB_Name = "module_options"
Public Type mystyle
    color           As Long
    size            As Integer
    bold            As Boolean
    italic          As Boolean
    underline       As Boolean
    line_through    As Boolean
End Type

Public Type mylink
    color           As Long
    size            As Integer
    bold            As Boolean
    italic          As Boolean
    underline       As Boolean
    line_through    As Boolean
    none            As Boolean
End Type

Public Type mytitle
    title   As String
    style   As mystyle
End Type

Public Type mybackground
    use_color       As Boolean
    color           As Long
    use_picture     As Boolean
    picture_name    As String
    repeat          As String
    attachement     As String
    hpos            As String
    hpostype        As String
    vpos            As String
    vpostype        As String
End Type

Public Type mytable
    use_bgcolor     As Boolean
    bgcolor         As Long
    use_bordercolor As Boolean
    bordercolor     As Long
    bordersize      As Integer
    height          As Integer
    heighttype      As String
    width           As Integer
    widthtype       As String
    align           As String
    cellpad         As Integer
    cellspace       As Integer
    bgpicture       As String
End Type

Public style_folder As mystyle
Public style_title  As mytitle
Public style_table  As mytable
Public style_link   As mylink
Public style_hover  As mylink
Public style_visited As mylink
Public style_background As mybackground

Public opt_all_and_words As String
Public opt_all_or_words As String
Public opt_all_not_words As String
Public opt_extensions As String
Public min_file_size As Long
Public max_file_size As Long
Public last_folder  As String

Public ini_file_name As String
'API Declaration for ini files

'Public Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
'

Public Sub update_ini()
    On Error Resume Next
        
    WritePrivateProfileString "default", "last_folder", last_folder, ini_file_name
    WritePrivateProfileString "default", "absolute", get_english_value(Form_main.mnugenerateabsolute.Checked), ini_file_name
        
    WritePrivateProfileString "style_folder", "bold", get_english_value(style_folder.bold), ini_file_name
    WritePrivateProfileString "style_folder", "color", CStr(style_folder.color), ini_file_name
    WritePrivateProfileString "style_folder", "italic", get_english_value(style_folder.italic), ini_file_name
    WritePrivateProfileString "style_folder", "line_through", get_english_value(style_folder.line_through), ini_file_name
    WritePrivateProfileString "style_folder", "size", CStr(style_folder.size), ini_file_name
    WritePrivateProfileString "style_folder", "underline", get_english_value(style_folder.underline), ini_file_name
    
    WritePrivateProfileString "style_title", "title", CStr(style_title.title), ini_file_name
    WritePrivateProfileString "style_title", "bold", get_english_value(style_title.style.bold), ini_file_name
    WritePrivateProfileString "style_title", "color", CStr(style_title.style.color), ini_file_name
    WritePrivateProfileString "style_title", "italic", get_english_value(style_title.style.italic), ini_file_name
    WritePrivateProfileString "style_title", "line_through", get_english_value(style_title.style.line_through), ini_file_name
    WritePrivateProfileString "style_title", "size", CStr(style_title.style.size), ini_file_name
    WritePrivateProfileString "style_title", "underline", get_english_value(style_title.style.underline), ini_file_name
    
    WritePrivateProfileString "style_link", "bold", get_english_value(style_link.bold), ini_file_name
    WritePrivateProfileString "style_link", "color", CStr(style_link.color), ini_file_name
    WritePrivateProfileString "style_link", "italic", get_english_value(style_link.italic), ini_file_name
    WritePrivateProfileString "style_link", "line_through", get_english_value(style_link.line_through), ini_file_name
    WritePrivateProfileString "style_link", "size", CStr(style_link.size), ini_file_name
    WritePrivateProfileString "style_link", "underline", get_english_value(style_link.underline), ini_file_name
    WritePrivateProfileString "style_link", "none", get_english_value(style_link.none), ini_file_name

    WritePrivateProfileString "style_hover", "bold", get_english_value(style_hover.bold), ini_file_name
    WritePrivateProfileString "style_hover", "color", CStr(style_hover.color), ini_file_name
    WritePrivateProfileString "style_hover", "italic", get_english_value(style_hover.italic), ini_file_name
    WritePrivateProfileString "style_hover", "line_through", get_english_value(style_hover.line_through), ini_file_name
    WritePrivateProfileString "style_hover", "size", CStr(style_hover.size), ini_file_name
    WritePrivateProfileString "style_hover", "underline", get_english_value(style_hover.underline), ini_file_name
    WritePrivateProfileString "style_hover", "none", get_english_value(style_hover.none), ini_file_name

    WritePrivateProfileString "style_visited", "bold", get_english_value(style_visited.bold), ini_file_name
    WritePrivateProfileString "style_visited", "color", CStr(style_visited.color), ini_file_name
    WritePrivateProfileString "style_visited", "italic", get_english_value(style_visited.italic), ini_file_name
    WritePrivateProfileString "style_visited", "line_through", get_english_value(style_visited.line_through), ini_file_name
    WritePrivateProfileString "style_visited", "size", CStr(style_visited.size), ini_file_name
    WritePrivateProfileString "style_visited", "underline", get_english_value(style_visited.underline), ini_file_name
    WritePrivateProfileString "style_visited", "none", get_english_value(style_visited.none), ini_file_name

    WritePrivateProfileString "backgroung", "use_color", get_english_value(style_background.use_color), ini_file_name
    WritePrivateProfileString "backgroung", "use_picture", get_english_value(style_background.use_picture), ini_file_name
    WritePrivateProfileString "backgroung", "color", CStr(style_background.color), ini_file_name
    WritePrivateProfileString "backgroung", "picture_name", CStr(style_background.picture_name), ini_file_name
    WritePrivateProfileString "backgroung", "repeat", CStr(style_background.repeat), ini_file_name
    WritePrivateProfileString "backgroung", "attachement", CStr(style_background.attachement), ini_file_name
    WritePrivateProfileString "backgroung", "hpos", CStr(style_background.hpos), ini_file_name
    WritePrivateProfileString "backgroung", "hpostype", CStr(style_background.hpostype), ini_file_name
    WritePrivateProfileString "backgroung", "vpos", CStr(style_background.vpos), ini_file_name
    WritePrivateProfileString "backgroung", "vpostype", CStr(style_background.vpostype), ini_file_name

    WritePrivateProfileString "table", "use_bgcolor", get_english_value(style_table.use_bgcolor), ini_file_name
    WritePrivateProfileString "table", "use_bordercolor", get_english_value(style_table.use_bordercolor), ini_file_name
    WritePrivateProfileString "table", "bgcolor", CStr(style_table.bgcolor), ini_file_name
    WritePrivateProfileString "table", "bordercolor", CStr(style_table.bordercolor), ini_file_name
    WritePrivateProfileString "table", "bordersize", CStr(style_table.bordersize), ini_file_name
    WritePrivateProfileString "table", "height", CStr(style_table.height), ini_file_name
    WritePrivateProfileString "table", "heighttype", CStr(style_table.heighttype), ini_file_name
    WritePrivateProfileString "table", "width", CStr(style_table.height), ini_file_name
    WritePrivateProfileString "table", "widthtype", CStr(style_table.heighttype), ini_file_name
    WritePrivateProfileString "table", "align", CStr(style_table.align), ini_file_name
    WritePrivateProfileString "table", "cellpad", CStr(style_table.cellpad), ini_file_name
    WritePrivateProfileString "table", "cellspace", CStr(style_table.cellspace), ini_file_name
    WritePrivateProfileString "table", "bgpicture", style_table.bgpicture, ini_file_name
    
    WritePrivateProfileString "Restrictions", "min_file_size", CStr(known_files_restrictions_array(0)), ini_file_name
    WritePrivateProfileString "Restrictions", "max_file_size", CStr(known_files_restrictions_array(1)), ini_file_name
    WritePrivateProfileString "Restrictions", "extensions", opt_extensions, ini_file_name
    WritePrivateProfileString "Restrictions", "all_and_words", opt_all_and_words, ini_file_name
    WritePrivateProfileString "Restrictions", "all_or_words", opt_all_or_words, ini_file_name
    WritePrivateProfileString "Restrictions", "all_not_words", opt_all_not_words, ini_file_name
    WritePrivateProfileString "Restrictions", "extensions", opt_extensions, ini_file_name

End Sub
Private Function get_english_value(var As Boolean) As String
    If var Then ' to avoid language troubles
        get_english_value = "True"
    Else
        get_english_value = "False"
    End If
End Function

Private Function get_bool_value(var As String) As Boolean
    If var = "True" Then
        get_bool_value = True
    Else
        get_bool_value = False
    End If
End Function
Public Sub get_ini_value()
    On Error Resume Next
    Dim retour              As String * 300
    Dim nb_char_retour      As Long
    Dim strtmp              As String
    Dim tmp                 As Boolean

   nb_char_retour = GetPrivateProfileString("default", "last_folder", "", retour, 300, ini_file_name)
        last_folder = Mid$(retour, 1, nb_char_retour)
   nb_char_retour = GetPrivateProfileString("default", "absolute", "False", retour, 300, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        tmp = get_bool_value(strtmp)
        Form_main.mnugenerateabsolute.Checked = tmp
        Form_main.mnugeneraterelative.Checked = Not tmp
        
   nb_char_retour = GetPrivateProfileString("style_folder", "bold", "False", retour, 300, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        style_folder.bold = get_bool_value(strtmp)
   nb_char_retour = GetPrivateProfileString("style_folder", "color", "0", retour, 300, ini_file_name)
       style_folder.color = Mid$(retour, 1, nb_char_retour)
   nb_char_retour = GetPrivateProfileString("style_folder", "italic", "False", retour, 300, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        style_folder.italic = get_bool_value(strtmp)
   nb_char_retour = GetPrivateProfileString("style_folder", "line_through", "False", retour, 300, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        style_folder.line_through = get_bool_value(strtmp)
   nb_char_retour = GetPrivateProfileString("style_folder", "size", "", retour, 300, ini_file_name)
        style_folder.size = Mid$(retour, 1, nb_char_retour)
   nb_char_retour = GetPrivateProfileString("style_folder", "underline", "False", retour, 300, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        style_folder.underline = get_bool_value(strtmp)
    
   nb_char_retour = GetPrivateProfileString("style_title", "title", "My CD 1", retour, 300, ini_file_name)
        style_title.title = Mid$(retour, 1, nb_char_retour)
   nb_char_retour = GetPrivateProfileString("style_title", "bold", "False", retour, 300, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        style_title.style.bold = get_bool_value(strtmp)
   nb_char_retour = GetPrivateProfileString("style_title", "color", "0", retour, 300, ini_file_name)
        style_title.style.color = Mid$(retour, 1, nb_char_retour)
   nb_char_retour = GetPrivateProfileString("style_title", "italic", "False", retour, 300, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        style_title.style.italic = get_bool_value(strtmp)
   nb_char_retour = GetPrivateProfileString("style_title", "line_through", "False", retour, 300, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        style_title.style.line_through = get_bool_value(strtmp)
   nb_char_retour = GetPrivateProfileString("style_title", "size", "30", retour, 300, ini_file_name)
        style_title.style.size = Mid$(retour, 1, nb_char_retour)
   nb_char_retour = GetPrivateProfileString("style_title", "underline", "False", retour, 300, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        style_title.style.underline = get_bool_value(strtmp)
        
   nb_char_retour = GetPrivateProfileString("style_link", "bold", "True", retour, 300, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        style_link.bold = get_bool_value(strtmp)
   nb_char_retour = GetPrivateProfileString("style_link", "color", "16777215", retour, 300, ini_file_name)
        style_link.color = Mid$(retour, 1, nb_char_retour)
   nb_char_retour = GetPrivateProfileString("style_link", "italic", "False", retour, 300, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        style_link.italic = get_bool_value(strtmp)
   nb_char_retour = GetPrivateProfileString("style_link", "line_through", "False", retour, 300, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        style_link.line_through = get_bool_value(strtmp)
   nb_char_retour = GetPrivateProfileString("style_link", "size", "", retour, 300, ini_file_name)
        style_link.size = Mid$(retour, 1, nb_char_retour)
   nb_char_retour = GetPrivateProfileString("style_link", "underline", "False", retour, 300, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        style_link.underline = get_bool_value(strtmp)
   nb_char_retour = GetPrivateProfileString("style_link", "none", "True", retour, 300, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        style_link.none = get_bool_value(strtmp)

   nb_char_retour = GetPrivateProfileString("style_hover", "bold", "True", retour, 300, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        style_hover.bold = get_bool_value(strtmp)
   nb_char_retour = GetPrivateProfileString("style_hover", "color", "255", retour, 300, ini_file_name)
        style_hover.color = Mid$(retour, 1, nb_char_retour)
   nb_char_retour = GetPrivateProfileString("style_hover", "italic", "False", retour, 300, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        style_hover.italic = get_bool_value(strtmp)
   nb_char_retour = GetPrivateProfileString("style_hover", "line_through", "False", retour, 300, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        style_hover.line_through = get_bool_value(strtmp)
   nb_char_retour = GetPrivateProfileString("style_hover", "size", "", retour, 300, ini_file_name)
        style_hover.size = Mid$(retour, 1, nb_char_retour)
   nb_char_retour = GetPrivateProfileString("style_hover", "underline", "False", retour, 300, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        style_hover.underline = get_bool_value(strtmp)
   nb_char_retour = GetPrivateProfileString("style_hover", "none", "True", retour, 300, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        style_hover.none = get_bool_value(strtmp)

   nb_char_retour = GetPrivateProfileString("style_visited", "bold", "True", retour, 300, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        style_visited.bold = get_bool_value(strtmp)
   nb_char_retour = GetPrivateProfileString("style_visited", "color", "12632256", retour, 300, ini_file_name)
        style_visited.color = Mid$(retour, 1, nb_char_retour)
   nb_char_retour = GetPrivateProfileString("style_visited", "italic", "False", retour, 300, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        style_visited.italic = get_bool_value(strtmp)
   nb_char_retour = GetPrivateProfileString("style_visited", "line_through", "False", retour, 300, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        style_visited.line_through = get_bool_value(strtmp)
   nb_char_retour = GetPrivateProfileString("style_visited", "size", "", retour, 300, ini_file_name)
        style_visited.size = Mid$(retour, 1, nb_char_retour)
   nb_char_retour = GetPrivateProfileString("style_visited", "underline", "False", retour, 300, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        style_visited.underline = get_bool_value(strtmp)
   nb_char_retour = GetPrivateProfileString("style_visited", "none", "True", retour, 300, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        style_visited.none = get_bool_value(strtmp)

   nb_char_retour = GetPrivateProfileString("backgroung", "use_color", "True", retour, 300, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        style_background.use_color = get_bool_value(strtmp)
   nb_char_retour = GetPrivateProfileString("backgroung", "use_picture", "False", retour, 300, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        style_background.use_picture = get_bool_value(strtmp)
   nb_char_retour = GetPrivateProfileString("backgroung", "color", "16744576", retour, 300, ini_file_name)
        style_background.color = Mid$(retour, 1, nb_char_retour)
   nb_char_retour = GetPrivateProfileString("backgroung", "picture_name", "", retour, 300, ini_file_name)
        style_background.picture_name = Mid$(retour, 1, nb_char_retour)
   nb_char_retour = GetPrivateProfileString("backgroung", "repeat", "no-repeat", retour, 300, ini_file_name)
        style_background.repeat = Mid$(retour, 1, nb_char_retour)
   nb_char_retour = GetPrivateProfileString("backgroung", "attachement", "fixed", retour, 300, ini_file_name)
        style_background.attachement = Mid$(retour, 1, nb_char_retour)
   nb_char_retour = GetPrivateProfileString("backgroung", "hpos", "", retour, 300, ini_file_name)
        style_background.hpos = Mid$(retour, 1, nb_char_retour)
   nb_char_retour = GetPrivateProfileString("backgroung", "hpostype", "pixels", retour, 300, ini_file_name)
        style_background.hpostype = Mid$(retour, 1, nb_char_retour)
   nb_char_retour = GetPrivateProfileString("backgroung", "vpos", "", retour, 300, ini_file_name)
        style_background.vpos = Mid$(retour, 1, nb_char_retour)
   nb_char_retour = GetPrivateProfileString("backgroung", "vpostype", "pixels", retour, 300, ini_file_name)
        style_background.vpostype = Mid$(retour, 1, nb_char_retour)

   nb_char_retour = GetPrivateProfileString("table", "use_bgcolor", "True", retour, 300, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        style_table.use_bgcolor = get_bool_value(strtmp)
   nb_char_retour = GetPrivateProfileString("table", "use_bordercolor", "True", retour, 300, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        style_table.use_bordercolor = get_bool_value(strtmp)
   nb_char_retour = GetPrivateProfileString("table", "bgcolor", "15406192", retour, 300, ini_file_name)
        style_table.bgcolor = Mid$(retour, 1, nb_char_retour)
   nb_char_retour = GetPrivateProfileString("table", "bordercolor", "8404992", retour, 300, ini_file_name)
        style_table.bordercolor = Mid$(retour, 1, nb_char_retour)
   nb_char_retour = GetPrivateProfileString("table", "bordersize", "3", retour, 300, ini_file_name)
        style_table.bordersize = Mid$(retour, 1, nb_char_retour)
   nb_char_retour = GetPrivateProfileString("table", "height", "", retour, 300, ini_file_name)
        style_table.height = Mid$(retour, 1, nb_char_retour)
   nb_char_retour = GetPrivateProfileString("table", "heighttype", "%", retour, 300, ini_file_name)
        style_table.heighttype = Mid$(retour, 1, nb_char_retour)
   nb_char_retour = GetPrivateProfileString("table", "width", "", retour, 300, ini_file_name)
        style_table.width = Mid$(retour, 1, nb_char_retour)
   nb_char_retour = GetPrivateProfileString("table", "widthtype", "%", retour, 300, ini_file_name)
        style_table.widthtype = Mid$(retour, 1, nb_char_retour)
   nb_char_retour = GetPrivateProfileString("table", "align", "Center", retour, 300, ini_file_name)
        style_table.align = Mid$(retour, 1, nb_char_retour)
   nb_char_retour = GetPrivateProfileString("table", "cellpad", "", retour, 300, ini_file_name)
        style_table.cellpad = Mid$(retour, 1, nb_char_retour)
   nb_char_retour = GetPrivateProfileString("table", "cellspace", "", retour, 300, ini_file_name)
        style_table.cellspace = Mid$(retour, 1, nb_char_retour)
   nb_char_retour = GetPrivateProfileString("table", "bgpicture", "", retour, 300, ini_file_name)
        style_table.bgpicture = Mid$(retour, 1, nb_char_retour)
   
    min_file_size = GetPrivateProfileInt("Restrictions", "min_file_size", 0, ini_file_name)
    max_file_size = GetPrivateProfileInt("Restrictions", "max_file_size", 0, ini_file_name)
    nb_char_retour = GetPrivateProfileString("Restrictions", "all_and_words", "", retour, 300, ini_file_name)
        opt_all_and_words = Mid$(retour, 1, nb_char_retour)
    nb_char_retour = GetPrivateProfileString("Restrictions", "all_or_words", "", retour, 300, ini_file_name)
        opt_all_or_words = Mid$(retour, 1, nb_char_retour)
    nb_char_retour = GetPrivateProfileString("Restrictions", "all_not_words", "", retour, 300, ini_file_name)
        opt_all_not_words = Mid$(retour, 1, nb_char_retour)
    nb_char_retour = GetPrivateProfileString("Restrictions", "extensions", "exe;zip;ace;rar", retour, 300, ini_file_name)
        opt_extensions = Mid$(retour, 1, nb_char_retour)
    'fill known_files_restrictions_array
    fill_restrictions_array min_file_size, max_file_size, opt_all_and_words, opt_all_or_words, opt_all_not_words, opt_extensions
    
    
End Sub

