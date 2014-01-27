VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ThisWorkbook.Saved = True
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim SaveDialog As FileDialog

    If Environ("username") <> "TReische" Then
        SaveThisBook
        Set SaveDialog = Application.FileDialog(msoFileDialogSaveAs)
        SaveDialog.InitialFileName = Environ("userprofile") & "\Desktop\"
        SaveDialog.Show
        If SaveDialog.SelectedItems.Count > 0 Then
            ActiveWorkbook.SaveAs SaveDialog.SelectedItems.Item(1), xlOpenXMLWorkbook
        End If
        Cancel = True
        ThisWorkbook.Close
    End If
End Sub

Private Sub Workbook_Open()
    CheckForUpdates RepositoryName, VersionNumber
End Sub
