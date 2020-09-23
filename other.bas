Attribute VB_Name = "other"
Public Declare Function ShellExecute Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As _
String, ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Type rgb_value
    red As Byte
    green As Byte
    blue As Byte
End Type

Public Function long_to_rgb(color_value As Long) As rgb_value
    On Error Resume Next
    Dim reste As Long
    Dim unite As Byte
    reste = color_value
    unite = CByte(reste Mod 256)
    long_to_rgb.red = unite
    reste = (reste - unite) / 256
    unite = CByte(reste Mod 256)
    long_to_rgb.green = unite
    reste = (reste - unite) / 256
    unite = CByte(reste Mod 256)
    long_to_rgb.blue = unite
End Function

Public Function long_to_html(color_value As Long) As String
    On Error Resume Next
    Dim tmp As rgb_value
    tmp = long_to_rgb(color_value)
    
    long_to_html = fixed_length_hex(tmp.red) & fixed_length_hex(tmp.green) & fixed_length_hex(tmp.blue)
End Function

Public Function fixed_length_hex(val As Byte) As String
    On Error Resume Next
    fixed_length_hex = Hex$(val)
    If Len(fixed_length_hex) < 2 Then
        fixed_length_hex = "0" & fixed_length_hex
    End If
End Function
