Attribute VB_Name = "CodeExport"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc : ExportCode
' Date : 3/19/2013
' Desc : Exports all modules
'---------------------------------------------------------------------------------------
Sub ExportCode()
    Dim comp As Variant
    Dim codeFolder As String
    Dim FileName As String
    Dim File As String

    'References Microsoft Visual Basic for Applications Extensibility 5.3
    AddReference "{0002E157-0000-0000-C000-000000000046}", 5, 3
    codeFolder = GetWorkbookPath & "Code\" & Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5) & "\"

    On Error Resume Next
    RecMkDir codeFolder
    On Error GoTo 0

    'Remove all previously exported modules
    File = Dir(codeFolder)
    Do While File <> ""
        DeleteFile codeFolder & File
        File = Dir
    Loop

    'Export modules in current project
    For Each comp In ThisWorkbook.VBProject.VBComponents
        Select Case comp.Type
            Case 1
                FileName = codeFolder & comp.Name & ".bas"
                comp.Export FileName
            Case 2
                FileName = codeFolder & comp.Name & ".cls"
                comp.Export FileName
            Case 3
                FileName = codeFolder & comp.Name & ".frm"
                comp.Export FileName
        End Select
    Next
End Sub

'---------------------------------------------------------------------------------------
' Proc : AddReferences
' Date : 3/19/2013
' Desc : Adds a reference to VBProject
'---------------------------------------------------------------------------------------
Sub AddReference(GUID As String, Major As Integer, Minor As Integer)
    Dim ID As Variant
    Dim Ref As Variant
    Dim Result As Boolean


    For Each Ref In ThisWorkbook.VBProject.References
        If Ref.GUID = GUID And Ref.Major = Major And Ref.Minor = Minor Then
            Result = True
        End If
    Next

    'References Microsoft Visual Basic for Applications Extensibility 5.3
    If Result = False Then
        ThisWorkbook.VBProject.References.AddFromGuid GUID, Major, Minor
    End If
End Sub

'---------------------------------------------------------------------------------------
' Proc : GetWorkbookPath
' Date : 3/19/2013
' Desc : Gets the full path of ThisWorkbook
'---------------------------------------------------------------------------------------
Function GetWorkbookPath() As String
    Dim fullName As String
    Dim wrkbookName As String
    Dim pos As Long

    wrkbookName = ThisWorkbook.Name
    fullName = ThisWorkbook.fullName

    pos = InStr(1, fullName, wrkbookName, vbTextCompare)

    GetWorkbookPath = Left$(fullName, pos - 1)
End Function

'---------------------------------------------------------------------------------------
' Proc : DeleteFile
' Date : 3/19/2013
' Desc : Deletes a file
'---------------------------------------------------------------------------------------
Sub DeleteFile(FileName As String, Optional LogEntry As Boolean = False)
    Kill FileName
End Sub

