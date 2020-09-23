VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Doc autorun"
   ClientHeight    =   615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2310
   Icon            =   "Form_main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   615
   ScaleWidth      =   2310
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As _
String, ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Load()
    Call autorun
    Unload Me
End Sub

Public Sub autorun()
    On Error GoTo lblerror
    Dim ShellExecutereturn  As Long
    Dim numfile             As Integer
    Dim filename            As String
    Dim buffer              As String
    Me.Hide
    
    filename = "doc_autorun.ini"
    
    
    filename = App.Path & "\" & filename
    num_file = FreeFile
    Open filename For Input As num_file
        Line Input #num_file, buffer
    Close num_file
    ShellExecutereturn = ShellExecute(0, "Open", buffer, "", App.Path, 1)
    Exit Sub
lblerror:
    MsgBox "An error as occured" & vbCrLf & "please verify if " & filename & vbCrLf _
            & "and " & buffer & " exist" _
            , vbExclamation, "Doc autorun"
End Sub
