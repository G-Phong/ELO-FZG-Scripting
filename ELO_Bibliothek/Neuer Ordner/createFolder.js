//CreateFolder.js
var message = "Hier bitte Kundennamen eingeben";

/*
* @params: Fenstername, Nachricht im Fenster, Default-Wert in der Input-Maske
*/
var value = workspace.showSimpleInputBox("ELO", message, "");


//Archivstartknoten mit ELO Objekt ID 1
var root = archive.getElement(1);

//Basispfad einmalig berechnen, das ¶ Zeichen ist wie ein \\
var basePath = "¶Kunden¶" + value + "¶";

//AddPath legt einen Ordner an und prüft davor, ob es den Ordner bereits gibt
root.addPath(basePath + "Bestellungen", 1);

//Aktualisieren der Ordner
workspace.activeView.refreshArchive();

//User-Feedback
workspace.setFeedbackMessage("Kundenordner angelegt");
