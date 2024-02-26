Dim filePath
filePath = "HIER DATEIPFAD EINFÜGEN"

Dim outputFilePath
outputFilePath = Left(filePath, Len(filePath) - 4) & "_Formularfelder.txt"

Set objWord = CreateObject("Word.Application")
Set objDoc = objWord.Documents.Open(filePath, False, True)
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.CreateTextFile(outputFilePath)

fieldCounter = 1

For Each objField In objDoc.FormFields
    strFieldName = "Formularfeld " & fieldCounter
    objTextFile.WriteLine strFieldName & ": " & objField.Name
    fieldCounter = fieldCounter + 1
    Next

    objTextFile.Close
    objDoc.Close False
    objWord.Quit
