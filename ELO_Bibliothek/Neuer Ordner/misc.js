// Globale JS variable "log"
let str = "1";
log.debug("Fehler" + str);
log.info("Fehler" + str);

// FeedbackMessage Ausgabe - muss man nicht bestätigen
let msg1 = "Projekt wird nun abgeschlossen";
workspace.setFeedbackMessage(msg1);

// InfoBox Ausgabe - muss man bestätigen (Titel, Hinweistext)
let msg2 = "Das ist eine InfoBox";
workspace.showInfoBox("Projektverwaltung", msg2);

// AlertBox Ausgabe - wie InfoBox, nur anderes Icon
let msg3 = "Projekt konnte nicht abgeschlossen werden";
workspace.showAlertBox("Projektverwaltung", msg3);

// QuestionBox Ausgabe - für Ja/Nein Abfragen
let msg4 = "Wollen Sie das Projekt jetzt abschließen";
if (workspace.showQuestionBox("Projektverwaltung", msg4)) {
  workspace.setFeedbackMessage("Projekt wird nun abgeschlossen");
}

// SimpleInputBox Ausgabe - Einfache Eingabeaufforderung ohne Titel
let input1 = workspace.showSimpleInputBox("Geben Sie den Projektnamen ein");
log.info("Eingegebener Projektnamen: " + input1);

// InputBox Ausgabe - Eingabeaufforderung mit Titel und Hinweistext
let input2 = workspace.showInputBox(
  "Projektverwaltung",
  "Geben Sie die Projektbeschreibung ein"
);
log.info("Eingegebene Projektbeschreibung: " + input2);

//DirectoriesAdapter-Objekt
workspace.directories;
workspace.directories.baseDir;
workspace.directories.checkOutDir;
workspace.directories.tempDir;
workspace.directories.trashDir;

//ViewAdapter
workspace.activeView;
workspace.activeView.refresh();
workspace.activeView.firstSelected;
workspace.activeView.firstSelected.loadSord();
workspace.activeView.firstSelected.loadSord().parentId;

//Archive-Adapter
archive.getElement(number id);
archive.getElementByGuid(String guid);
archive.getElementByArcpath(String path); //Bei Pfaden immer Trennsymbol (Pilcrow-Zeichen ¶) nutzen 
archive.name;
archive.getUserNames(boolean, boolean);

//Verschlagwortungsdialog indexDialog-Objekt
indexDialog.name;
indexDialog.docMaskId;
indexDialog.objKeyValue;
indexDialog.setObjKeyValue;
indexDialog.sord; //aktuelles Verschlagwortungsobjekt holen

//Aufgabenansicht - TasksAdapter-Objekt "tasks"
tasks.firstSelected.taskElement.archiveElement;

/*
Startet einen neuen Workflow aus einer Workflowvorlage heraus. Die neu erzeugte Workflow-ID wird als Rückgabewert geliefert. Als Parameter werden die ELO-Objekt-ID, der Name des neuen Workflows und die ID des Workflowtempla-
tes angegeben. Die Template-ID kann über den Workflowdesigner ermittelt wer-
den.
*/
startWorkflow(String objId, String name, int templateId); 

/*
Liest einen Workflow ein. An einigen Stellen steht nicht der komplette Workflow sondern nur die Workflow-ID und die Knoten-ID zur Verfügung. Über diese Funk-
tion erhalten Sie die vollständigen Workflow-lnformationen zur Bearbeitung.
*/
getWorkflow(int workflowId);

//Dateifunktionen - Datie kopieren
let source = new File("C:\\temp\\sourcefile.txt");
let destination = new File("C:\\temp\\newfile.txt");
utils.copyFile(source,destination);
