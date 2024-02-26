/*
*Auswahl eines Archiveintrags
*/

let title = "Projekte";
let msg = "W�hlen Sie das Projektverzeichnis aus.";
let projects = 18; //ELO Object ID des Projekt Ordners als Start
let id = workspace.showTreeSelectDialog(
  title,
  msg,
  projects,
  false,
  true,
  false
);
if (id > 0) {
  workspace.setFeedbackMessage(id);
}
