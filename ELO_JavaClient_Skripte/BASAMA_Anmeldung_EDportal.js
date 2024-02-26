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
  newButton(BUTTON_BASAMA, AnmeldungBASAMA);
}

function AnmeldungBASAMA() {
  prnt(
    "Neue Studienarbeit wird angemeldet. Bitte geben Sie im Folgenden die Basisdaten ein."
  );

  // Erstellen eines Sords als Vorlage fuer den Workflow
  var root = archive.getElementByArcpath(AnmeldungBASAMA_DIR_OPEN);
  var editInfo = ixc.createSord(root.id, 14, EditInfoC.mbAll); //MASK ID 14 (AnmeldungBASAMA_MASK)
  var sord = editInfo.sord;

  //Maske für Dateneingabe, IndexDialog anzeigen
  if (!indexDialog.editSord(sord, true, "Basisdaten eingeben")) {
    return;
  }

  //Basisdaten aus Verschlagwortung auslesen (direkt, ohne lib_const)
  var vorname_autor = indexDialog.getObjKeyValue("LIT_FIRST_NAME");
  var nachname_autor = indexDialog.getObjKeyValue("LIT_SURNAME");
  var abschlussarbeit_typ = indexDialog.getObjKeyValue("LIT_TYPE");
  var studi_email = indexDialog.getObjKeyValue("LIT_EMAIL_STUD");

  //Ordner spezifisch benennen
  sord.name = vorname_autor.toString() + "_" + nachname_autor.toString();

  // Ordner in Archivbaum hinzufuegen + SordList aktualisieren
  var dest = archive.getElement(root.id);
  dest.addStructure(sord);
  prnt(
    "Ordner wird hinzugefügt: " + "217_Studienarbeiten\\" + sord.name.toString()
  );
  workspace.updateSordLists();
  //var sord_id = dest.id;

  //Je nach dem was für eine Arbeit es ist soll das entsprechende Template kopiert werden

  try {
    var doc_type_id = 0;

    if (abschlussarbeit_typ == "Bachelorarbeit") {
      doc_type_id = 2197; // Objekt ID des Templates für Bachelorarbeit (statisch)
      //prnt("in bachelorarbeit");
    } else if (abschlussarbeit_typ == "Masterarbeit") {
      doc_type_id = 1960; // Objekt ID des Templates für Masterarbeit (statisch)
      //prnt("in MA");
    } else if (abschlussarbeit_typ == "Semesterarbeit") {
      doc_type_id = 1961; // Objekt ID des Templates für Semesterarbeit (statisch)
      //prnt("in SA");
    } else {
      prnt("Fehler beim Auslesen des Abschlussarbeits-Typs!");
      prnt("Type-Variable: " + abschlussarbeit_typ);
    }

    //Kopieren des Word-Templates in das tempDir
    var source = archive.getElement(doc_type_id).getFile();
    var dest = utils.getUniqueFile(
      workspace.directories.tempDir,
      "Anmeldung_" +
        vorname_autor.toString() +
        "_" +
        nachname_autor.toString() +
        ".pdf"
    );
    utils.copyFile(source, dest);
  } catch (error) {
    prnt("Fehler beim Kopieren der Vorlage!");
  }

  //Aktualisieren
  workspace.updateSordLists();

  //Dokument in den richtigen ELO Ordner einhaengen
  //Pfad zum spezifischen Forschungsvorhaben Word File
  var anmeldung_dok = archive.getElementByArcpath(
    AnmeldungBASAMA_DIR_OPEN + "¶" + sord.name.toString()
  );

  var prepdoc = anmeldung_dok.prepareDocument(0);
  var docpath =
    workspace.directories.tempDir +
    "\\" +
    "Anmeldung_" +
    vorname_autor.toString() +
    "_" +
    nachname_autor.toString() +
    ".pdf";
  var elodoc = anmeldung_dok.addDocument(prepdoc, docpath);
  workspace.updateSordLists();
  prnt(
    "Anmeldungsformular (" +
      abschlussarbeit_typ +
      ") wurde ausgefüllt und wie folgt abgelegt:" +
      "200_Organisation" +
      "\\" +
      "217_Studienarbeiten" +
      "\\" +
      vorname_autor.toString() +
      "_" +
      nachname_autor.toString() +
      "\\" +
      "Anmeldung_" +
      vorname_autor.toString() +
      "_" +
      nachname_autor.toString() +
      ".pdf"
  );

  //TODO: PDF ausfüllen

  prnt("Starte Workflow: BASAMA_Anmeldung_EDportal");

  //Pfad zum Ordner der Studienarbeit
  var studienarbeit_ordner = archive.getElementByArcpath(
    "¶200_Organisation¶217_Studienarbeiten" + "¶" + sord.name.toString()
  );

  //Workflow "BASAMA_Anmeldung_ED"portal" starten
  var workflow = ixc.checkoutWorkFlow(
    "BASAMA_Anmeldung_EDportal",
    WFTypeC.TEMPLATE,
    WFDiagramC.mbAll,
    LockC.NO
  );
  workflow.type = WFTypeC.ACTIVE;
  workflow.id = -1;
  workflow.objId = studienarbeit_ordner.getId(); //TODO: ID des Anmeldungsdokuments angeben -> je nach art
  ixc.checkinWorkFlow(workflow, WFDiagramC.mbAll, LockC.NO);
}

function eloFlowConfirmDialogStart() {
  var dialog = dialogs.flowConfirmDialog;
  var node = dialog.currentNode;

  if (node.name == "Anmeldung der Studienarbeit im ED-Portal") {
    prnt(
      "Im Folgenden öffnet sich im Browser das ED-Portal. Bitte dort die Studienarbeit anmelden!"
    );

    //Browser öffnen mit Link
    //Java Runtime ermöglicht es externe Befehle auszuführen -> Standardbrowser öffnen und bestimmte URL laden
    var url = "https://portal.ed.tum.de/de/Theses/melden";
    var runtime = java.lang.Runtime.getRuntime();
    runtime.exec(["rundll32", "url.dll,FileProtocolHandler", url]);
  }

  if (node.name == "Unterlagen Studienarbeit an Studierenden versenden") {
    //Verschlagwortung neu auslesen
    var vorname_autor = indexDialog.getObjKeyValue("LIT_FIRST_NAME");
    var nachname_autor = indexDialog.getObjKeyValue("LIT_SURNAME");
    var abschlussarbeit_typ = indexDialog.getObjKeyValue("LIT_TYPE");
    var studi_email = indexDialog.getObjKeyValue("LIT_EMAIL_STUD");

    //Dokument IDs (Starterpaket) in ELO für Studi - manuell ausgelesen (statisch)
    var dok_formeln_id = 16391;
    var dok_zitierrichtlinie_id = 2212;
    var dok_präsentationsvorlage_id = 2213;
    var dok_wissFehlverh_id = 16389;

    //Info
    prnt(
      "Starterpaket mit Studienarbeitsunterlagen  an Studierenden wird zusammengestellt. Outlook wird nun geöffnet..."
    );

    // Variablen erstellen
    var subject = "Unterlagen für die " + abschlussarbeit_typ.toString();
    var emailText =
      "Hallo " +
      vorname_autor +
      ",<BR>" +
      "<BR>" +
      "anbei befinden sich die Unterlagen für deine Studienarbeit.<BR>" +
      "<BR>" +
      "Viele Grüße,<BR>" +
      "FZG";

    // Anhänge: ein .ecd File (eine ELO-Verlinkung) -> dieses soll in ELO erstellt, im tempDir ordner abgelegt werden und mit in den Anhang gepackt werden
    var attachments = [
      attachFile(archive.getElement(dok_formeln_id)),
      attachFile(archive.getElement(dok_zitierrichtlinie_id)),
      attachFile(archive.getElement(dok_präsentationsvorlage_id)),
      attachFile(archive.getElement(dok_wissFehlverh_id)),
    ];

    //E-Mail an E-Mail-Verteiler senden
    mail(studi_email, subject, emailText, attachments, undefined);

    prnt("Bitte Outlook-Mail überprüfen und dann manuell absenden!");
  }
}
