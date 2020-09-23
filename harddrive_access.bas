Attribute VB_Name = "harddrive_access"
Option Explicit
Public Type my_file
    folder      As String
    full_path   As String
    file_name   As String
    file_size   As Long
End Type

Public files_found()      As my_file

Private array_directories() As String


Public Sub find_files(ByVal directory As String, Optional search_in_subdirectories As Boolean = True)
    'search files from directories and add them to files_found
    On Error Resume Next
    
    If Right$(directory, 1) <> "\" Then directory = directory & "\"
    ReDim files_found(0)
    If search_in_subdirectories Then
        ReDim array_directories(1)
        array_directories(1) = directory
    
        Dim last_dir As String
        Do While UBound(array_directories) > 0
            last_dir = array_directories(UBound(array_directories))
            ReDim Preserve array_directories(UBound(array_directories) - 1)
            find_through_directory last_dir
        Loop
    Else
        find_through_directory directory
    End If
End Sub

Private Sub find_through_directory(ByVal directory As String)
    On Error Resume Next
    Dim myname          As String
    Dim tmpsize         As Long
    Dim table_size      As Variant
    Dim arr_dir_size    As Long

    myname = Dir(directory, vbArchive Or vbDirectory Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem)
    
    Do While myname <> ""   ' Start the loop.
        ' Ignore the current directory and the encompassing directory.
        If myname <> "." And myname <> ".." Then
            ' Use bitwise comparison to make sure myname is a directory.
            If (GetAttr(directory & myname) And vbDirectory) = vbDirectory Then
                arr_dir_size = UBound(array_directories) + 1
                ReDim Preserve array_directories(arr_dir_size)
                array_directories(arr_dir_size) = directory & myname & "\"
            Else
              'If (GetAttr(directory & myname) And vbArchive) = vbArchive _
                 Or (GetAttr(directory & myname) And vbHidden) = vbHidden _
                 Or (GetAttr(directory & myname) And vbNormal) = vbNormal _
                 Or (GetAttr(directory & myname) And vbReadOnly) = vbReadOnly _
                 Or (GetAttr(directory & myname) And vbSystem) = vbSystem _
               Then
              'add file to files_found
                tmpsize = FileLen(directory & myname) / 1000 ' in Kbytes
                If check_knownfiles_restrictions(myname, tmpsize) Then
                    table_size = UBound(files_found)
                    With files_found(table_size)
                        .file_name = LCase$(myname)
                        .file_size = tmpsize
                        .full_path = LCase$(directory & myname)
                        .folder = LCase$(directory)
                    End With
                    ReDim Preserve files_found(table_size + 1)
                End If
            End If
        End If
        myname = Dir ' Get next entry.
    Loop

End Sub


''''''''''''''''''''''' file system object functions
Public Function is_file_existing(full_path As String) As Boolean
    Dim fso As New FileSystemObject
    is_file_existing = fso.FileExists(full_path)
End Function

Public Function is_folder_existing(full_path As String) As Boolean
    Dim fso As New FileSystemObject
    is_folder_existing = fso.FolderExists(full_path)
End Function

Public Function get_file_name(full_path As String) As String
    Dim fso As New FileSystemObject
    get_file_name = fso.GetFileName(full_path)
End Function

Public Function get_folder_name(full_path As String) As String
    Dim fso As New FileSystemObject
    get_folder_name = fso.GetParentFolderName(full_path)
End Function

Public Function get_extension(full_path As String) As String
    Dim fso As New FileSystemObject
    get_extension = fso.GetExtensionName(full_path)
End Function
