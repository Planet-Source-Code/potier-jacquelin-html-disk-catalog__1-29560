VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form form_entries 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entries found"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   Icon            =   "form_entries.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   6825
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   2400
      Width           =   1335
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   9128
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Uncheck entries you don't want to see in the generated HTML page"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "form_entries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    On Error Resume Next
    Dim cpt As Long
    Dim tmparray() As my_file
    Dim nb_checked  As Long
    
    tmparray = files_found
    nb_checked = 0
    For cpt = 1 To UBound(files_found)
        If Me.ListView1.ListItems.Item(cpt).Checked = True Then
            files_found(nb_checked) = tmparray(cpt - 1)
            nb_checked = nb_checked + 1
        End If
    Next cpt
    ReDim Preserve files_found(nb_checked)
    
    '''''
    
    Dim full_path As String
    
    Form_main.CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    Form_main.CommonDialog1.Filter = "html files (*.htm;*.html)|*.htm;*.html| All (*.*) |*.*"
    Form_main.CommonDialog1.ShowSave
    full_path = Form_main.CommonDialog1.filename
    
    
    make_html full_path, False, Form_main.mnugenerateabsolute.Checked
    
    MsgBox "File generated", vbExclamation, "HTML Disk Catalog"
    Unload Me
ErrHandler:
   Exit Sub
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim cpt As Long
    Me.ListView1.ColumnHeaders.Item(1).width = Me.ListView1.width - 100
    For cpt = 0 To UBound(files_found) - 1
        Me.ListView1.ListItems.Add cpt + 1, , files_found(cpt).full_path
        Me.ListView1.ListItems.Item(cpt + 1).Checked = True
    Next cpt
End Sub
