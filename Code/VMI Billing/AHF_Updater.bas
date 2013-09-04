Attribute VB_Name = "AHF_Updater"
Option Explicit

Private Enum Ver
    Major
    Minor
    Patch
End Enum

'---------------------------------------------------------------------------------------
' Proc : IncrementMajor
' Date : 9/4/2013
' Desc : Increments the macros major version number (major.minor.patch)
'---------------------------------------------------------------------------------------
Sub IncrementMajor()
    IncrementVer Major
End Sub

'---------------------------------------------------------------------------------------
' Proc : IncrementMinorVersion
' Date : 4/24/2013
' Desc : Increments the macros minor version number (major.minor.patch)
'---------------------------------------------------------------------------------------
Sub IncrementMinor()
    IncrementVer Minor
End Sub

'---------------------------------------------------------------------------------------
' Proc : IncrementPatch
' Date : 9/4/2013
' Desc : Increments the macros patch number (major.minor.patch)
'---------------------------------------------------------------------------------------
Sub IncrementPatch()
    IncrementVer Patch
End Sub

'---------------------------------------------------------------------------------------
' Proc : IncrementVer
' Date : 9/4/2013
' Desc :
'---------------------------------------------------------------------------------------
Private Sub IncrementVer(Version As Ver)
    Dim Path As String
    Dim Ver As Variant
    Dim FileNum As Integer
    Dim i As Integer

    Path = GetWorkbookPath & "Version.txt"
    FileNum = FreeFile

    If FileExists(Path) = True Then
        Open Path For Input As #FileNum
        Line Input #FileNum, Ver
        Close FileNum

        'Split version number
        Ver = Split(Ver, ".")

        'Increment version
        Select Case Version
            Case Major
                Ver(0) = CInt(Ver(0)) + 1
            Case Minor
                Ver(1) = CInt(Ver(1)) + 1
            Case Patch
                Ver(2) = CInt(Ver(2)) + 1
        End Select

        'Combine version
        Ver = Ver(0) & "." & Ver(1) & "." & Ver(2)

        Open Path For Output As #FileNum
        Print #FileNum, Ver
        Close #FileNum
    Else
        Open Path For Output As #FileNum
        Print #FileNum, "1.0.0"
        Close #FileNum
    End If
End Sub

'---------------------------------------------------------------------------------------
' Proc : CheckForUpdates
' Date : 4/24/2013
' Desc : Checks to see if the macro is up to date
'---------------------------------------------------------------------------------------
Sub CheckForUpdates(URL As String, Optional RepoName As String = "")
    Dim Ver As Variant
    Dim LocalVer As Variant
    Dim Path As String
    Dim LocalPath As String
    Dim FileNum As Integer
    Dim RegEx As Variant

    Set RegEx = CreateObject("VBScript.RegExp")
    Ver = DownloadTextFile(URL)
    Ver = Ver & vbCrLf
    Ver = Replace(Ver, vbLf, "")
    Ver = Replace(Ver, vbCr, "")
    RegEx.Pattern = "^[0-9]+\.[0-9]+\.[0-9]+$"
    Path = GetWorkbookPath & "Version.txt"
    FileNum = FreeFile

    Open Path For Input As #FileNum
    Line Input #FileNum, LocalVer
    Close FileNum

    If RegEx.Test(Ver) Then
        If Not Ver = LocalVer Then
            MsgBox Prompt:="An update is available. Please close the macro and get the latest version!", Title:="Update Available"
            If Not RepoName = "" Then
                Shell "C:\Program Files\Internet Explorer\iexplore.exe http://github.com/Wesco/" & RepoName & "/releases/", vbMaximizedFocus
            End If
        End If
    End If
End Sub

'---------------------------------------------------------------------------------------
' Proc : DownloadTextFile
' Date : 4/25/2013
' Desc : Returns the contents of a text file from a website
'---------------------------------------------------------------------------------------
Private Function DownloadTextFile(URL As String) As String
    Dim success As Boolean
    Dim responseText As String
    Dim oHTTP As Variant

    Set oHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")

    oHTTP.Open "GET", URL, False
    oHTTP.Send
    success = oHTTP.WaitForResponse()

    If Not success Then
        DownloadTextFile = ""
        Exit Function
    End If

    responseText = oHTTP.responseText
    Set oHTTP = Nothing

    DownloadTextFile = responseText
End Function
