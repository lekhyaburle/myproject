    Dim fso
    Dim ObjFolder
    Dim ObjFiles
    Dim ObjFile
    Dim objExcel
	Dim Workbooks
    'MsgBox ("Going back In")
    'Creating File System Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    'Set stdout = fso.GetStandardStream (1)
    'stdout.WriteLine "This will go to standard output."
    'Getting the Folder Object
    Set ObjFolder = fso.GetFolder("C:\test")
    'stdout.WriteLine "This will go to standard output."
    'Getting the list of Files
    Set ObjFiles = ObjFolder.Files
        'On Error Resume Next
        For Each ObjFile In ObjFiles
            If LCase(Right(ObjFile.Name, 5)) = ".xlsx" Or LCase(Right(ObjFile.Name, 4)) = ".xls" Then
				Set xl  = CreateObject("Excel.Application")
				Set wb = xl.Workbooks.Open(ObjFile)
				wb.Activate
                'Workbooks.Open(ObjFile).Activate
                'stdout.WriteLine "This will go to standard output."
                wb.Sheets(1).Range("A1").Value = "rueyfueau"
                wb.RefreshAll
                wb.Save
                wb.Close
            End If
        Next


