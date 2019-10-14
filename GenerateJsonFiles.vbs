Option Explicit

On Error Resume Next

GenerateJsonFiles

Sub GenerateJsonFiles()
    
    Dim FSO
    Dim path
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    path = FSO.GetAbsolutePathName(".")
  
    Call RunThroughFolder(FSO, path)
  
End Sub

Sub RunThroughFolder(ByRef FSO, ByVal path)

    Dim oFolder
    Dim oSubFolders
    Dim oFiles
    Dim File
    Dim Item
    
    Set oFolder = FSO.GetFolder(path)
    Set oSubFolders = oFolder.SubFolders
    Set oFiles = oFolder.Files

    
    For Each File In oFiles
        If LCase(FSO.GetExtensionName(File.Name)) = "xlsm" Then
			IF Left(File.Name,1) <> "~" then
				Dim xlApp
				Dim xlBook
				Set xlApp = CreateObject("Excel.Application")
				Set xlBook = xlApp.Workbooks.Open(File, 0, True)
				xlApp.Run "ThisWorkbook.collectDataFortransform"
				xlApp.Quit
    
				Set xlBook = Nothing
				Set xlApp = Nothing
				
			End If
        End If
    Next
    
    For Each Item In oSubFolders
        Call RunThroughFolder(FSO, Item.path)
    Next

End Sub