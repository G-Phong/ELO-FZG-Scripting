' VBScript zum Auslesen und Durchnummerieren aller Formularfelder in einem Word-Dokument und Speichern in einer Textdatei

' Dateipfad zum Word-Dokument
Dim filePath
filePath = "C:\Users\giaph\AppData\Roaming\ELO Digital Office\FZG_test\405\temp\00_Vorhaben-Kurzbeschreibung.docx"

' Dateipfad f�r die Textdatei
Dim outputFilePath
outputFilePath = Left(filePath, Len(filePath) - 4) & "_Formularfelder.txt"

' Word-Anwendung erstellen
Set objWord = CreateObject("Word.Application")

' Dokument �ffnen (nicht sichtbar)
Set objDoc = objWord.Documents.Open(filePath, False, True)

' Z�hlvariable f�r die Durchnummerierung der Formularfelder
fieldCounter = 1

' Textdatei erstellen und �ffnen
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.CreateTextFile(outputFilePath)

' Alle Formularfelder durchlaufen und durchnummeriert in die Textdatei schreiben
For Each objField In objDoc.FormFields
    strFieldName = "Formularfeld " & fieldCounter
    objTextFile.WriteLine strFieldName & ": " & objField.Name
    fieldCounter = fieldCounter + 1
Next

' Textdatei schlie�en
objTextFile.Close

' Dokument schlie�en
objDoc.Close False

' Word-Anwendung beenden
objWord.Quit
