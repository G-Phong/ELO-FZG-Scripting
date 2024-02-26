// Ueber Bestaetigungsdialoge kann man Code waehrend des Workflows (zu bestimmten Workflow-Knoten) ausfuehren, und nicht nur zu Anfang
function eloFlowConfirmDialogStart() {
  var dialog = dialogs.flowConfirmDialog;
  var node = dialog.currentNode;

  //Abfrage der einzelnen Workflow-Knoten -> je nach dem in welchem Knoten des Workflows man sich befindet soll etwas anderes passieren -> if-else Abfrage
  if (node.name == "Eingabe der Basisdaten des Forschungsvorhabens") {
    //Nutzer wird zur Überprüfung der Daten gebeten
  } else if (node.name == "Ergänzung von Grobdaten und Projektbeschreibung") {
    //Nutzer wird zur Eingabe der Grobdaten / Projektbeschreibung gebeten
  } else if (node.name == "PDF-Version exportieren") {
    //Word wird zu einer PDF exportiert
    //Nutzer muss überprüfen und bestätigen
  } else if (node.name == "Formular weiterleiten per Mail") {
    //E-Mail-Vorlage wird erstellt
    //Nutzer muss auf "Senden klicken"
  }
}
