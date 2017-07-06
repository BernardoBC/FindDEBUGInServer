Const ForReading = 1
Const ForWriting = 2
Set FSO = CreateObject("Scripting.FileSystemObject")
dim counter
dim minicounter
counter = 0
i = 0
p = 0
drive = "s:"
StartTime = Timer()

'Setting up files
strWriteFile = ".\output.txt"
Set objFSO=CreateObject("Scripting.FileSystemObject")
Set objOutFile = objFSO.CreateTextFile(strWriteFile,True)

Set Directory = FSO.GetFolder(drive)
Wscript.Echo now & ": Directory: " & Directory.Path
Set colFiles = Directory.Files

'Main Directory loop
For Each FileName In colFiles
    counter = counter + 1
    Wscript.Echo now & ": " & FileName.Name

    'Check if file is empty
    Set objFile = FSO.OpenTextFile(Directory.Path&"\"&FileName.Name, ForReading)
    IF FSO.GetFile(Directory.Path&"\"&FileName.Name).size <> 0 then    

        'File is larger than 0 (not empty)        
        strText = objFile.ReadAll

        'Count Instances
        arrLines = Split(strText, " ")
        For Each strLine in arrLines
            If InStr(strLine, "level=DEBUG") Then
                p = p + 1
            End If
        Next

        'Replace all instances
        'strNewText = Replace(strText, "level=DEBUG", "level=INFO")
        'Set objFile = FSO.OpenTextFile(Directory.Path&"\"&FileName.Name, ForWriting)
        'objFile.WriteLine strNewText

        'Wscript.Echo now & ": " & "Instances of 'level=DEBUG': " & p
        minicounter = minicounter + 1

    END IF
    
    objFile.Close            


Next

'Check subfolders
ShowSubfolders FSO.GetFolder(drive)

'Check SubFolder Procedures (recursive)
Sub ShowSubFolders(Folder)
    
    'Subfolder Loop
    For Each Subfolder in Folder.SubFolders
    	minicounter = 0
        Wscript.Echo now & ": " & Subfolder.Path
        Set Directory = FSO.GetFolder(Subfolder.Path)
        Set colFiles = Directory.Files
        
        'Check Files in directory'
        For Each FileName in colFiles
            'READ FILE HERE
            Wscript.Echo now & ": " & FileName.Name

            'Read File
            Set objFile = FSO.OpenTextFile(Directory.Path&"\"&FileName.Name, ForReading)
            IF FSO.GetFile(Directory.Path&"\"&FileName.Name).size <> 0 then    

                'Process text file then search for string
                strText = objFile.ReadAll

                'Count Instances
                arrLines = Split(strText, " ")
                For Each strLine in arrLines
                    If InStr(strLine, "level=DEBUG") Then
                        p = p + 1
                    End If
                Next

                'Replace all instances
                'strNewText = Replace(strText, "level=DEBUG", "level=INFO")
                'Set objFile = FSO.OpenTextFile(Directory.Path&"\"&FileName.Name, ForWriting)
                'objFile.WriteLine strNewText

                'Wscript.Echo now & ": " & "Instances of 'level=DEBUG': " & p
                minicounter = minicounter + 1

            END IF
            
            objFile.Close            

        Next

        'Call Recursively
        ShowSubFolders Subfolder
        counter = counter + minicounter
    Next
End Sub

'Script Summary
EndTime = Timer()
objOutFile.Write counter & "," & p & "," & FormatNumber(EndTime - StartTime, 2), "," &

On Error Resume Next 
pause
 If Err.Number <> 0 Then 
    strWriteFile = ".\output.txt"
    Set objOutFile = objFSO.CreateTextFile(strWriteFile,True)
    objOutFile.Write "Read files failed: " & Err.Number & ") - " & Err.Description 

    Err.Clear 
End If