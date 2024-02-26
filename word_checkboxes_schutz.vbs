' VBScript zum �ffnen eines Word-Dokuments ohne Schreibschutz, Ankreuzen aller CheckBoxes und Einf�gen von "XYZ" in alle anderen Formularfelder

Dim filePath
filePath = "C:\TUM_MASTER\6. + 7. SEMESTER (SS23) - MA\Hiwi Job TUM 2023\FZG_ELO_HIWI\JS_ELO_VSCODE_UMGEBUNG\00_Vorhaben-Kurzbeschreibung.docx"

' Word-Anwendung erstellen
Set objWord = CreateObject("Word.Application")

' Dokument ohne Schreibschutz �ffnen (ReadOnly auf False setzen)
Set objDoc = objWord.Documents.Open(filePath, False, False) ' Der dritte Parameter (ReadOnly) ist auf False gesetzt, um das Dokument ohne Schreibschutz zu �ffnen

' Z�hlvariable f�r die gecheckten CheckBoxes
checkedCheckBoxes = 0

' Alle Formularfelder durchlaufen und CheckBoxes ankreuzen, oder "XYZ" in andere Felder einf�gen
For Each objField In objDoc.FormFields
    If objField.Type = 71 Then ' Type 71 entspricht einem CheckBox-Formularfeld
        objField.CheckBox.Value = True ' CheckBox ankreuzen
        checkedCheckBoxes = checkedCheckBoxes + 1 ' Z�hler erh�hen
    ElseIf objField.Type <> 71 Then ' Wenn es sich nicht um ein CheckBox-Formularfeld handelt
        objField.Result = "XYZ" ' "XYZ" in das Formularfeld einf�gen
    End If
Next

' Dokument speichern
objDoc.Save

' Dokument schlie�en
objDoc.Close

' Word-Anwendung beenden
objWord.Quit

' Freigeben der Objekte
Set objDoc = Nothing
Set objWord = Nothing

' Echo mit Anzahl der gecheckten CheckBoxes ausgeben
WScript.Echo "Es wurden " & checkedCheckBoxes & " CheckBoxes gecheckt. 'XYZ' wurde in alle anderen Formularfelder eingef�gt."
