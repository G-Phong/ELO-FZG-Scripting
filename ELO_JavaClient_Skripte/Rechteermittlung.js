//@include lib_utility
//@include lib_const
//@include lib_buttons

// ELO
importPackage(Packages.de.elo.ix.jscript);
importPackage(Packages.de.elo.client);
importPackage(Packages.de.elo.client);

// .NET-Framework / COM-Schnittstelle
//importPackage(Packages.de.elo.client);
importPackage(Packages.de.jacob.com);
importPackage(Packages.de.jacob.activeX);
importPackage(Packages.de.ms.activeX);
importPackage(Packages.de.ms.com);

// Java Lib
importPackage(Packages.java.io);
importPackage(Packages.java.util);
importPackage(Packages.java.nio.file);
importPackage(Packages.java.lang);
importPackage(Packages.java.net);
var importNames = JavaImporter();
importNames.importPackage(Packages.com.ms.com);
importNames.importPackage(Packages.com.ms.activeX);
importClass(Packages.com.jacob.activeX.ActiveXComponent);
importClass(Packages.com.jacob.com.Dispatch);
importClass(Packages.com.jacob.com.Variant);

function eloExpandRibbon() {
  createRibbon();
  newButton(BUTTON_RECHTEERMITTLUNG, Rechteermittlung);
}

function addEntry(dialog, name, value, line) {
  dialog.addLabel(1, line, 1, name);
  dialog.addLabel(2, line, 1, value);
}

function Rechteermittlung() {
  //Test Anwenderrechte
  var dialog = workspace.createGridDialog("Anwenderrechte", 2, 25);
  var panel = dialog.gridPanel;
  panel.setBackground(255, 255, 240);

  var label = dialog.addLabel(1, 1, 2, "Anwenderrechte");
  label.bold = true;
  label.fontSize = 18;

  var line = 3;
  var access = workspace.userRights;

  addEntry(dialog, "Ordner anlegen", access.hasAddStructureRight(), line++);
  addEntry(dialog, "Hauptadministrator", access.hasAdminRight(), line++);
  addEntry(
    dialog,
    "Verfallsdatum ändern",
    access.hasChangeDelDateRight(),
    line++
  );
  addEntry(
    dialog,
    "Verschlagwortungsmaske ändern",
    access.hasChangeMaskRight(),
    line++
  );
  addEntry(dialog, "Passwort ändern", access.hasChangePasswordRight(), line++);
  addEntry(
    dialog,
    "Berechtigungseinstellungen ändern",
    access.hasChangePermissionsRight(),
    line++
  );
  addEntry(
    dialog,
    "Revisionslevel ändern",
    access.hasChangeRevLevelRight(),
    line++
  );
  addEntry(
    dialog,
    "Dokumente löschen",
    access.hasDeleteDocumentRight(),
    line++
  );
  addEntry(
    dialog,
    "Dokumentenversionen löschen",
    access.hasDeleteDocVersionRight(),
    line++
  );

  addEntry(
    dialog,
    "Archivstruktur bearbeiten",
    access.hasEditArchiveRight(),
    line++
  );
  addEntry(
    dialog,
    "Dokumente bearbeiten",
    access.hasEditDocumentRight(),
    line++
  );
  addEntry(
    dialog,
    "Stichwortlisten bearbeiten",
    access.hasEditKeywordListsRight(),
    line++
  );
  addEntry(
    dialog,
    "Replikationsvereise bearbeiten",
    access.hasEditReplSetsRight(),
    line++
  );
  addEntry(
    dialog,
    "Scannereinstellungen bearbeiten",
    access.hasEditScanRight(),
    line++
  );
  addEntry(dialog, "Skripte bearbeiten", access.hasEditScriptRight(), line++);

  dialog.show();

}


