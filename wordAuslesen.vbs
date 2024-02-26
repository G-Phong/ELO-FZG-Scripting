' VBScript zum Auslesen und Durchnummerieren aller Formularfelder in einem Word-Dokument und Speichern in einer Textdatei

' Dateipfad zum Word-Dokument
Dim filePath
filePath = "C:\Users\giaph\AppData\Roaming\ELO Digital Office\FZG_test\405\temp\00_Vorhaben-Kurzbeschreibung.docx"

' Dateipfad für die Textdatei
Dim outputFilePath
outputFilePath = Left(filePath, Len(filePath) - 4) & "_Formularfelder.txt"

' Word-Anwendung erstellen
Set objWord = CreateObject("Word.Application")

' Dokument öffnen (nicht sichtbar)
Set objDoc = objWord.Documents.Open(filePath, False, True)

' Zählvariable für die Durchnummerierung der Formularfelder
fieldCounter = 1

' Textdatei erstellen und öffnen
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.CreateTextFile(outputFilePath)

' Alle Formularfelder durchlaufen und durchnummeriert in die Textdatei schreiben
For Each objField In objDoc.FormFields
    strFieldName = "Formularfeld " & fieldCounter
    objTextFile.WriteLine strFieldName & ": " & objField.Name
    fieldCounter = fieldCounter + 1
Next

' Textdatei schließen
objTextFile.Close

' Dokument schließen
objDoc.Close False

' Word-Anwendung beenden
objWord.Quit
