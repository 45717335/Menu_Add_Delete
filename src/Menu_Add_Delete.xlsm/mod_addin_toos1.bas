Attribute VB_Name = "mod_addin_toos1"
Option Explicit

Function Update(sAddinServerPath As String, Optional b_r As Boolean = False) As String
    On Error Resume Next
    Dim NewAddin As Workbook
    Dim fs As Object
    Dim sAddinName As String
    sAddinName = Right(sAddinServerPath, Len(sAddinServerPath) - InStrRev(sAddinServerPath, "\"))
    Dim sAddinLocalPath As String
    sAddinLocalPath = Application.UserLibraryPath & sAddinName
    get_addin(sAddinName).Installed = False
    DoEvents
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Workbooks.Count = 0 Then Workbooks.Add
    If b_r = False Then
        fs.copyFile sAddinServerPath, Application.UserLibraryPath, True
        Application.AddIns(sAddinName).Installed = True
        Set NewAddin = Workbooks.Open(sAddinLocalPath)
        get_addin(sAddinName).Installed = True
    Else
        With Workbooks(sAddinName)
            .Saved = True
            .ChangeFileAccess xlReadOnly
            fs.copyFile sAddinLocalPath, ThisWorkbook.Path & "\"
            Kill .FullName
            .Close 0
        End With
    End If
    Update = sAddinLocalPath
End Function

Function get_addin(fln As String) As AddIn
    For Each get_addin In Application.AddIns
        If fln = get_addin.Name Then
            Exit For
        End If
    Next
End Function

Sub add_mm()

    Dim str1 As String, str2 As String, str3 As String
    Dim wb As Workbook
    Dim rg As Range
    str2 = ThisWorkbook.Name
    str1 = ActiveWorkbook.Name
    If str1 <> str2 Then
        Workbooks(str2).Activate
        update_status
        Exit Sub
    End If
    Set wb = ActiveWorkbook
    update_status
    For Each rg In Selection
        str3 = rg
        If str3 Like "*.xlam" Then
            Update str3
            update_status
        Else
            Update InputBox("input file full path of *.xlam")
            update_status
        End If
        Exit Sub
    Next
End Sub

Sub del_mm()

    Dim str1 As String, str2 As String, str3 As String
    Dim wb As Workbook
    Dim rg As Range
    str2 = ThisWorkbook.Name
    str1 = ActiveWorkbook.Name
    If str1 <> str2 Then
        Workbooks(str2).Activate
        update_status
        Exit Sub
    End If
    Set wb = ActiveWorkbook
    update_status
    For Each rg In Selection
        str3 = rg
        If str3 Like "*.xlam" Then
            Update str3, True
            update_status
        Else
            Update InputBox("input file full path of *.xlam"), True
            update_status
        End If
        Exit Sub
    Next

End Sub

Private Function get_files(ws As Worksheet, fd As String)
    On Error Resume Next
    If Right(fd, 1) <> "\" Then fd = fd & "\"
    Dim str1 As String
    Dim i As Integer
    Dim c As Range
    str1 = Dir(fd & "*.xlam")
    i = 3
    Do While Len(str1) > 0
        ws.Range("A" & i) = fd & str1
        If get_addin(str1) Is Nothing Then
            ws.Range("A" & i).Interior.Color = RGB(255, 255, 255)
        Else
            If get_addin(str1).Installed = True Then
                ws.Range("A" & i).Interior.Color = RGB(0, 255, 0)
            Else
                ws.Range("A" & i).Interior.Color = RGB(255, 255, 255)
            End If
        End If
        str1 = Dir()
        i = i + 1
    Loop
End Function

Private Sub update_status()
    get_files ThisWorkbook.ActiveSheet, ThisWorkbook.Path & "\"
End Sub
            
Sub copy_addin()
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    Dim str1 As String, str2 As String
    str1 = Application.UserLibraryPath
    If right(str1, 1) <> "\" Then str1 = str1 & "\"
    str2 = dir(str1 & "*.xlam")
    Do While Len(str2) > 0
        If fs.FileExists(ThisWorkbook.Path & "\" & str2) = False Then
            fs.copyFile str1 & str2, ThisWorkbook.Path & "\"
        End If
        str2 = dir()
    Loop
End Sub
