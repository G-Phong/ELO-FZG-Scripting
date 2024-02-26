' Dateipfad zum Word-Dokument
Dim filePath
filePath = "C:\TUM_MASTER\6. + 7. SEMESTER (SS23) - MA\Hiwi Job TUM 2023\FZG_ELO_HIWI\JS_ELO_VSCODE_UMGEBUNG\00_Vorhaben-Kurzbeschreibung.docx"

' Dateipfad f�r die Textdatei
Dim outputFilePath
outputFilePath = Left(filePath, Len(filePath) - 4) & "_Formularfelder.txt"

' Word-Anwendung erstellen
Set objWord = CreateObject("Word.Application")

' Dokument �ffnen (nicht sichtbar)
Set objDoc = objWord.Documents.Open(filePath, False, True)

' Z�hlvariable f�r die Durchnummerierung der Formularfelder
fieldCounter = 1

' Z�hlvariable f�r die Checkboxen
checkboxCount = 0

' Textdatei erstellen und �ffnen
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.CreateTextFile(outputFilePath)

' Alle Formularfelder durchlaufen und durchnummeriert in die Textdatei schreiben
For Each objField In objDoc.FormFields
    strFieldName = "Formularfeld " & fieldCounter
    objTextFile.WriteLine strFieldName & ": " & objField.Name
    fieldCounter = fieldCounter + 1
    ' �berpr�fen, ob das Formularfeld eine Checkbox ist (FORMCHECKBOX)
    If objField.Type = 3 Then ' 3 entspricht FORMCHECKBOX
        ' Checkbox ankreuzen
        objField.CheckBox.Value = True
        checkboxCount = checkboxCount + 1
    End If
Next

' Anzahl der Checkboxen ausgeben
WScript.Echo "Das Formular hat insgesamt " & checkboxCount & " Checkboxen."

' Textdatei schlie�en
objTextFile.Close

' Dokument speichern
objDoc.Save

' Dokument schlie�en
objDoc.Close False

' Word-Anwendung beenden
objWord.Quit

' Alle WINWORD-Instanzen beenden
Set objShell = CreateObject("WScript.Shell")
objShell.Run "taskkill /f /im WINWORD.EXE", , True
