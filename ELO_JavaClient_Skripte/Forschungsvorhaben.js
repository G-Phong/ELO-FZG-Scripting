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

//------- Button im entsprechenden Ribbon-Band hinzufuegen ----------
function eloExpandRibbon() {
  createRibbon();
  newButton(BUTTON_FORSCHUNGSVORHABEN, Forschungsvorhaben);
}

//------- Funktion Forschungsvorhaben ----------
function Forschungsvorhaben() {
  prnt(
    "Neues Forschungsvorhaben wird angelegt. Bestätigen Sie mit 'OK' und geben Sie dann die Basisdaten ein."
  );

  // Erstellen eines Sords als Vorlage fuer den Workflow
  var root = archive.getElementByArcpath(FORSCHUNGSVORHABEN_DIR_OPEN); //"¶200_Organisation¶216_Forschungsvorhaben"
  var editInfo = ixc.createSord(
    root.id,
    FORSCHUNGSVORHABEN_MASK,
    EditInfoC.mbAll
  ); //MASK ID 552
  var sord = editInfo.sord;

  //Maske für Dateneingabe, IndexDialog anzeigen
  if (!indexDialog.editSord(sord, true, "Basisdaten eingeben")) {
    return;
  }

  //Basisdaten aus Verschlagwortung auslesen (via Konstanten aus lib_const.js)
  var KBF_nummer = indexDialog.getObjKeyValue(FORSCHUNGSVORHABEN_KBF_NUMMER);
  var bezeichnung = indexDialog.getObjKeyValue(FORSCHUNGSVORHABEN_BEZEICHNUNG);
  var FZG_nummer = indexDialog.getObjKeyValue(FORSCHUNGSVORHABEN_FZG_NR);
  var sachbearbeiter = indexDialog.getObjKeyValue(
    FORSCHUNGSVORHABEN_SACHBEARBEITER
  );
  var AL = indexDialog.getObjKeyValue(FORSCHUNGSVORHABEN_AL);
  var laufzeit = indexDialog.getObjKeyValue(FORSCHUNGSVORHABEN_LAUFZEIT);
  var anz_pruefkoerper = indexDialog.getObjKeyValue(
    FORSCHUNGSVORHABEN_PRUEFKOERPER
  );
  var leistung_werkstatt = indexDialog.getObjKeyValue(
    FORSCHUNGSVORHABEN_L_WERKSTATT
  );
  var leistung_prueffeld = indexDialog.getObjKeyValue(
    FORSCHUNGSVORHABEN_L_PRUEFFELD
  );
  var leistung_elabor = indexDialog.getObjKeyValue(FORSCHUNGSVORHABEN_L_ELABOR);
  var leistung_labor = indexDialog.getObjKeyValue(FORSCHUNGSVORHABEN_L_LABOR);

  //Ordner spezifisch benennen
  sord.name = KBF_nummer.toString() + "_FV_" + bezeichnung.toString();

  // Ordner in Archivbaum hinzufuegen + SordList aktualisieren
  var dest = archive.getElement(root.id);
  dest.addStructure(sord);
  prnt(
    "Ordner wird hinzugefügt: " +
      "216_Forschungsvorhaben\\" +
      sord.name.toString()
  );
  workspace.updateSordLists();

  //Pfad zum spezifischen Forschungsvorhaben Word File
  var forschungsDokument = archive.getElementByArcpath(
    FORSCHUNGSVORHABEN_DIR_OPEN + "¶" + sord.name.toString()
  );

  try {
    //Kopieren des Word-Templates in das tempDir
    var source = archive.getElement(14824).getFile(); //Word Vorlage aus Vorlagenordner (ueber Objekt ID), Vorlage darf nicht gelöscht werden!
    var dest = utils.getUniqueFile(
      workspace.directories.tempDir,
      KBF_nummer.toString() + "_Vorhaben-" + bezeichnung.toString() + ".docx"
    );
    utils.copyFile(source, dest);
  } catch (error) {
    prnt("Fehler beim Kopieren der Vorlage!");
  }

  //Aktualisieren
  workspace.updateSordLists();

  //______Word-Datei ausfüllen mit Verschlagwortungsdaten (Basisdaten)______
  prnt(
    "Word-Formular (Laufzettel) wird nun mit den eben eingegebenen Basisdaten automatisch gefüllt!"
  );

  //Ausgabe der Basisdaten als ELO Textdialog
  var infoMessage = "<HTML>";
  infoMessage =
    infoMessage + "Die Word-Vorlage wird nun mit folgenden Daten gefüllt: <BR>";
  infoMessage = infoMessage + "KBF-Nummer: " + KBF_nummer + "<BR>";
  infoMessage = infoMessage + "Bezeichnung:  " + bezeichnung + "<BR>";
  infoMessage = infoMessage + "FZG Nr.:  " + FZG_nummer + "<BR>";
  infoMessage = infoMessage + "Sachbearbeiter:  " + sachbearbeiter + "<BR>";
  infoMessage = infoMessage + "Abteilungsleiter: " + AL + "<BR>";
  infoMessage = infoMessage + "Laufzeit: " + laufzeit + "<BR>";
  infoMessage = infoMessage + "Anzahl Prüfkörper: " + anz_pruefkoerper + "<BR>";
  infoMessage =
    infoMessage + "Werkstatt Leistung: " + leistung_werkstatt + "<BR>";
  infoMessage =
    infoMessage + "Prüffeld Leistung: " + leistung_prueffeld + "<BR>";
  infoMessage = infoMessage + "E-Labor Leistung: " + leistung_elabor + "<BR>";
  infoMessage = infoMessage + "Labor Leistung: " + leistung_labor + "<BR>";

  //Ausgabe zur Ueberpruefung
  prnt(infoMessage);

  //Formularfelder aus "00_Vorhaben-Kurzbeschreibung.docx, wurde extern ausgelesen via VB-Skript"
  //Word-Formular darf nicht geaendert werden, sonst muss das Skript auch geaendert werden
  var formularfelder = {
    Formularfeld1: "Text16", //Forschungsvorhaben-Name
    Formularfeld2: "Text17", //Kurzwort
    Formularfeld3: "Text18", // FZG-Nr
    Formularfeld4: "Text19", // Leer?
    Formularfeld5: "Text29", // Sachbearbeiter
    Formularfeld6: "Text21", // AL
    Formularfeld7: "Text30", // Laufzeit

    //Prüfstand
    Formularfeld8: "Kontrollkästchen1", //Vorhanden?
    Formularfeld9: "Kontrollkästchen2", //umgebaut?
    Formularfeld10: "Kontrollkästchen3", //neu?
    Formularfeld11: "Kontrollkästchen2", //Mechanik?
    Formularfeld12: "Kontrollkästchen3", //Elektrik/Messtechnik?

    Formularfeld13: "Text25", //Art und Anzahl Prüfkörper

    //Vermessung
    Formularfeld14: "Kontrollkästchen4", //Geometrie?
    Formularfeld15: "",
    Formularfeld16: "Kontrollkästchen4", //Rauheit?
    Formularfeld17: "Kontrollkästchen4", //Gewicht?
    Formularfeld18: "Text23", //Freies Textfeld (Andere)
    Formularfeld19: "",
    Formularfeld20: "",
    Formularfeld21: "",
    Formularfeld22: "Kontrollkästchen4",
    Formularfeld23: "Text31", //Freies Textfeld (Andere)
    Formularfeld24: "Text26", //Werkstatt
    Formularfeld25: "Text27", //Prüffeld
    Formularfeld26: "Text13", //E-Labor
    Formularfeld27: "Text14", //Labor
    Formularfeld28: "Text28", //Sachbearbeiter
  };

  //Daten definieren für Word Vorlage
  var fieldnames = [];
  var entries = [];

  var formularfelderKeys = Object.keys(formularfelder);

  for (var i = 0; i < formularfelderKeys.length; i++) {
    var formularfeld = formularfelderKeys[i];
    fieldnames[i] = formularfelder[formularfeld];
  }

  for (var i = 0; i < formularfelderKeys.length; i++) {
    entries[i] = "";
  }

  entries[0] = bezeichnung;
  entries[1] = KBF_nummer;
  entries[2] = FZG_nummer;
  entries[3] = "";
  entries[4] = sachbearbeiter;
  entries[5] = AL;
  entries[6] = laufzeit;
  entries[12] = anz_pruefkoerper;
  entries[23] = leistung_werkstatt;
  entries[24] = leistung_prueffeld;
  entries[25] = leistung_elabor;
  entries[26] = leistung_labor;

  var docPath =
    workspace.directories.tempDir +
    "\\" +
    KBF_nummer.toString() +
    "_Vorhaben-" +
    bezeichnung.toString() +
    ".docx";

  //Ruft eine Funktion zum Ausfuellen des Word-Dokuments auf -> Funktion gibt File zurück
  var elodoc = fillWordTemplate(
    docPath,
    fieldnames,
    entries,
    sord.name,
    KBF_nummer,
    bezeichnung
  );

  prnt("Word wird nun geoeffnet...");

  //____________

  //_____Word Datei öffnen_____
  //Navigation zum Ordner im Java Client über gotoId()
  workspace.updateSordLists();
  workspace.getActiveView().refresh();
  workspace.gotoId(sord.id);

  var ueberpruefen_msg = "<HTML>";
  ueberpruefen_msg =
    ueberpruefen_msg +
    "Bitte Word-Datei überprüfen und dann einchecken!" +
    "<BR>";
  ueberpruefen_msg =
    ueberpruefen_msg +
    "Für die Bearbeitung der Kontrollkästchen muss ggf. der Dokumentenschutz aktiviert werden:" +
    "<BR>";
  ueberpruefen_msg =
    ueberpruefen_msg +
    "Überprüfen -> Schützen -> Bearbeitung einschr. -> 'Ja, Schutz jetzt anwenden' -> Ok";
  prnt(ueberpruefen_msg);

  //Dokument auschecken -> wird im lokalen temporären checkout Ordner abgelegt
  var elodoc_checkout = elodoc.checkOut();
  utils.editFile(checkout.getLastDocument().getDocumentFile());

  var checkIn_msg =
    "Haben Sie das Word-Dokument überprüft? Klicken Sie auf 'Ja' um das Dokument automatisch einzuchecken";

  if (workspace.showQuestionBox("Word-Dokument überprüft?", checkIn_msg)) {
    prnt("Das Word-Dokument wird nun eingecheckt!");
    //CheckIn des dokuments nach Abfrage
    elodoc_checkout.checkIn("2", "Basisdaten eingegeben");
  } else {
    prnt("Bitte checken Sie das Dokument nach der Bearbeitung selbst ein!");
  }

  //Fehlerpruefung
  if (forschungsDokument.isLocked()) {
    prnt(
      "Fehler beim Einchecken des Dokuments. Bitte checken Sie das Dokument manuell ein!"
    );
  }

  //____________

  //Workflow "Forschungsvorhaben" starten
  var workflow = ixc.checkoutWorkFlow(
    "Forschungsvorhaben_confirmDialog",
    WFTypeC.TEMPLATE,
    WFDiagramC.mbAll,
    LockC.NO
  );
  workflow.type = WFTypeC.ACTIVE;
  workflow.id = -1;
  workflow.objId = forschungsDokument.getId();
  workflow.name = "Forschungsvorhaben";
  ixc.checkinWorkFlow(workflow, WFDiagramC.mbAll, LockC.NO);

  /*  var objId = forschungsDokument.id;
    var name = forschungsDokument.name;
    tasks.startWorkflow(objId, "Beispiel-WF" + name, 19); */
}

// Ueber Bestaetigungsdialoge kann man Code waehrend des Workflows (zu bestimmten Workflow-Knoten) ausfuehren, und nicht nur zu Anfang
function eloFlowConfirmDialogStart() {
  var dialog = dialogs.flowConfirmDialog;
  var node = dialog.currentNode;

  //Hier keine Verschlagwortung auslesen!

  //Abfrage der einzelnen Workflow-Knoten -> je nach dem in welchem Knoten des Workflows man sich befindet soll etwas anderes passieren -> if-else Abfrage
  if (node.name == "Eingabe der Basisdaten des Forschungsvorhabens") {
    //Basisdaten aus Verschlagwortung auslesen (via Konstanten aus lib_const.js)
    var KBF_nummer = indexDialog.getObjKeyValue(FORSCHUNGSVORHABEN_KBF_NUMMER);
    var bezeichnung = indexDialog.getObjKeyValue(
      FORSCHUNGSVORHABEN_BEZEICHNUNG
    );
    var FZG_nummer = indexDialog.getObjKeyValue(FORSCHUNGSVORHABEN_FZG_NR);
    var sachbearbeiter = indexDialog.getObjKeyValue(
      FORSCHUNGSVORHABEN_SACHBEARBEITER
    );
    var AL = indexDialog.getObjKeyValue(FORSCHUNGSVORHABEN_AL);
    var laufzeit = indexDialog.getObjKeyValue(FORSCHUNGSVORHABEN_LAUFZEIT);
    var anz_pruefkoerper = indexDialog.getObjKeyValue(
      FORSCHUNGSVORHABEN_PRUEFKOERPER
    );
    var leistung_werkstatt = indexDialog.getObjKeyValue(
      FORSCHUNGSVORHABEN_L_WERKSTATT
    );
    var leistung_prueffeld = indexDialog.getObjKeyValue(
      FORSCHUNGSVORHABEN_L_PRUEFFELD
    );
    var leistung_elabor = indexDialog.getObjKeyValue(
      FORSCHUNGSVORHABEN_L_ELABOR
    );
    var leistung_labor = indexDialog.getObjKeyValue(FORSCHUNGSVORHABEN_L_LABOR);

    var sordName = KBF_nummer.toString() + "_FV_" + bezeichnung.toString();

    var forschungsDokument = archive.getElementByArcpath(
      FORSCHUNGSVORHABEN_DIR_OPEN +
        "¶" +
        KBF_nummer.toString() +
        "_FV_" +
        bezeichnung.toString() +
        "¶" +
        KBF_nummer.toString() +
        "_Vorhaben-" +
        bezeichnung.toString()
    );

    //Ausgabe der Basisdaten als ELO Textdialog
    var infoMessage = "<HTML>";
    infoMessage =
      infoMessage + "Die Word-Vorlage wurde mit folgenden Daten gefüllt: <BR>";
    infoMessage = infoMessage + "KBF-Nummer: " + KBF_nummer + "<BR>";
    infoMessage = infoMessage + "Bezeichnung:  " + bezeichnung + "<BR>";
    infoMessage = infoMessage + "FZG Nr.:  " + FZG_nummer + "<BR>";
    infoMessage = infoMessage + "Sachbearbeiter:  " + sachbearbeiter + "<BR>";
    infoMessage = infoMessage + "Abteilungsleiter: " + AL + "<BR>";
    infoMessage = infoMessage + "Laufzeit: " + laufzeit + "<BR>";
    infoMessage =
      infoMessage + "Anzahl Prüfkörper: " + anz_pruefkoerper + "<BR>";
    infoMessage =
      infoMessage + "Werkstatt Leistung: " + leistung_werkstatt + "<BR>";
    infoMessage =
      infoMessage + "Prüffeld Leistung: " + leistung_prueffeld + "<BR>";
    infoMessage = infoMessage + "E-Labor Leistung: " + leistung_elabor + "<BR>";
    infoMessage = infoMessage + "Labor Leistung: " + leistung_labor + "<BR>";

    prnt(infoMessage);

    //Navigation zum Ordner
    workspace.updateSordLists();
    workspace.getActiveView().refresh();
    workspace.gotoId(forschungsDokument.id);

    prnt("Sie können den Workflow nun weiterleiten!");

    return;
  } else if (node.name == "Ergänzung von Grobdaten und Projektbeschreibung") {
    //Basisdaten aus Verschlagwortung auslesen (via Konstanten aus lib_const.js)
    var KBF_nummer = indexDialog.getObjKeyValue(FORSCHUNGSVORHABEN_KBF_NUMMER);
    var bezeichnung = indexDialog.getObjKeyValue(
      FORSCHUNGSVORHABEN_BEZEICHNUNG
    );
    var FZG_nummer = indexDialog.getObjKeyValue(FORSCHUNGSVORHABEN_FZG_NR);
    var sachbearbeiter = indexDialog.getObjKeyValue(
      FORSCHUNGSVORHABEN_SACHBEARBEITER
    );
    var AL = indexDialog.getObjKeyValue(FORSCHUNGSVORHABEN_AL);
    var laufzeit = indexDialog.getObjKeyValue(FORSCHUNGSVORHABEN_LAUFZEIT);
    var anz_pruefkoerper = indexDialog.getObjKeyValue(
      FORSCHUNGSVORHABEN_PRUEFKOERPER
    );
    var leistung_werkstatt = indexDialog.getObjKeyValue(
      FORSCHUNGSVORHABEN_L_WERKSTATT
    );
    var leistung_prueffeld = indexDialog.getObjKeyValue(
      FORSCHUNGSVORHABEN_L_PRUEFFELD
    );
    var leistung_elabor = indexDialog.getObjKeyValue(
      FORSCHUNGSVORHABEN_L_ELABOR
    );
    var leistung_labor = indexDialog.getObjKeyValue(FORSCHUNGSVORHABEN_L_LABOR);

    var sordName = KBF_nummer.toString() + "_FV_" + bezeichnung.toString();

    var forschungsDokument = archive.getElementByArcpath(
      FORSCHUNGSVORHABEN_DIR_OPEN +
        "¶" +
        KBF_nummer.toString() +
        "_FV_" +
        bezeichnung.toString() +
        "¶" +
        KBF_nummer.toString() +
        "_Vorhaben-" +
        bezeichnung.toString()
    );
    //Info
    prnt(
      "Bitte in das vorausgefüllte Word-Formular die Grobdaten und die Beschreibung des Projekts eintragen. " +
        "Dies geschieht aktuell durch das manuelle Bearbeiten des Word-Dokuments. " +
        "Das Formular befindet sich im Ordner: " +
        "\\216_Forschungsvorhaben" +
        "\\" +
        sordName.toString() +
        "\\" +
        KBF_nummer.toString() +
        "_Vorhaben-" +
        bezeichnung.toString() +
        ".docx"
    );

    //Prüfen ob das Dokument schon ausgecheckt ist!
    if (forschungsDokument.isLocked()) {
      prnt(
        "Dokument wird gerade von jemandem bearbeitet. Bitte Bearbeiter/-in informieren oder spaeter nochmals versuchen."
      );
      return -1;
    }
    try {
      var word_dok_checkout = forschungsDokument.checkOut();
      utils.editFile(word_dok_checkout.getDocumentFile());
      prnt("Word wird nun geoeffnet...");
      prnt(
        "Bitte ergänzen Sie die Grobdaten und die Projektbeschreibung im Word-Dokument."
      );
    } catch (e) {
      prnt(
        "Beim Auschecken ist ein Fehler aufgetreten. Bitte überprüfen Sie den Bearbeitungsstatus des Dokuments. " +
          e.toString()
      );
    }

    var checkIn_msg =
      "Haben Sie Ihre Daten fertig eingetragen? Klicken Sie auf 'Ja' um das Dokument automatisch einzuchecken!";

    //Navigation zum Ordner
    workspace.updateSordLists();
    workspace.getActiveView().refresh();
    workspace.gotoId(forschungsDokument.id);

    if (workspace.showQuestionBox("Word-Dokument ergänzt?", checkIn_msg)) {
      prnt("Das Word-Dokument wird nun eingecheckt!");
      //CheckIn nach Abfrage
      word_dok_checkout.checkIn("3", "Grobdaten ergänzt");
      prnt(
        "Gehen Sie nun zu ihrem Aufgabenbereich und leiten Sie den Workflow weiter."
      );
    } else {
      prnt("Bitte checken Sie das Dokument nach der Bearbeitung selbst ein!");
    }

    //Fehlerpruefung
    if (forschungsDokument.isLocked()) {
      prnt(
        "Fehler beim Einchecken des Dokuments. Bitte checken Sie das Dokument manuell ein!"
      );
    }

    //Wird dafür ein Feld bereitgestellt?
  } else if (node.name == "PDF-Version exportieren") {
    //Basisdaten aus Verschlagwortung auslesen (via Konstanten aus lib_const.js)
    var KBF_nummer = indexDialog.getObjKeyValue(FORSCHUNGSVORHABEN_KBF_NUMMER);
    var bezeichnung = indexDialog.getObjKeyValue(
      FORSCHUNGSVORHABEN_BEZEICHNUNG
    );
    var FZG_nummer = indexDialog.getObjKeyValue(FORSCHUNGSVORHABEN_FZG_NR);
    var sachbearbeiter = indexDialog.getObjKeyValue(
      FORSCHUNGSVORHABEN_SACHBEARBEITER
    );
    var AL = indexDialog.getObjKeyValue(FORSCHUNGSVORHABEN_AL);
    var laufzeit = indexDialog.getObjKeyValue(FORSCHUNGSVORHABEN_LAUFZEIT);
    var anz_pruefkoerper = indexDialog.getObjKeyValue(
      FORSCHUNGSVORHABEN_PRUEFKOERPER
    );
    var leistung_werkstatt = indexDialog.getObjKeyValue(
      FORSCHUNGSVORHABEN_L_WERKSTATT
    );
    var leistung_prueffeld = indexDialog.getObjKeyValue(
      FORSCHUNGSVORHABEN_L_PRUEFFELD
    );
    var leistung_elabor = indexDialog.getObjKeyValue(
      FORSCHUNGSVORHABEN_L_ELABOR
    );
    var leistung_labor = indexDialog.getObjKeyValue(FORSCHUNGSVORHABEN_L_LABOR);

    var sordName = KBF_nummer.toString() + "_FV_" + bezeichnung.toString();

    var forschungsDokument = archive.getElementByArcpath(
      FORSCHUNGSVORHABEN_DIR_OPEN +
        "¶" +
        KBF_nummer.toString() +
        "_FV_" +
        bezeichnung.toString() +
        "¶" +
        KBF_nummer.toString() +
        "_Vorhaben-" +
        bezeichnung.toString()
    );
    //Info
    prnt(
      "Das Word Dokument wird nun automatisch in eine PDF exportiert. Die PDF-Version befindet sich im selben Ordner unter: " +
        "\\216_Forschungsvorhaben" +
        "\\" +
        sordName.toString() +
        "\\" +
        KBF_nummer.toString() +
        "_Vorhaben-" +
        bezeichnung.toString() +
        ".docx"
    );

    //_____Word-Formular als PDF exportieren_____
    var word_dok = archive.getElementByArcpath(
      FORSCHUNGSVORHABEN_DIR_OPEN +
        "¶" +
        KBF_nummer.toString() +
        "_FV_" +
        bezeichnung.toString() +
        "¶" +
        KBF_nummer.toString() +
        "_Vorhaben-" +
        bezeichnung.toString()
    );

    //File kopieren in tempDir
    var tmp = utils.getUniqueFile(
      workspace.directories.tempDir,
      KBF_nummer.toString() + "_Vorhaben-" + bezeichnung.toString() + ".docx"
    );
    utils.copyFile(word_dok.getFile(), tmp);

    //_____PDF-Export via Funktion_____
    //Paramter für Funktionsaufruf definieren
    try {
      var inputFilePath =
        workspace.directories.tempDir +
        "\\" +
        KBF_nummer.toString() +
        "_Vorhaben-" +
        bezeichnung.toString() +
        ".docx";

      var root = archive.getElementByArcpath(FORSCHUNGSVORHABEN_DIR_OPEN); //"¶200_Organisation¶216_Forschungsvorhaben"
      var editInfo = ixc.createSord(
        root.id,
        FORSCHUNGSVORHABEN_MASK,
        EditInfoC.mbAll
      ); //MASK ID 552
      var sord = editInfo.sord;

      //Parameter für Funktionsaufruf definieren
      var outputFilePath =
        workspace.directories.tempDir +
        "\\" +
        KBF_nummer.toString() +
        "_Vorhaben-" +
        bezeichnung.toString() +
        ".pdf";

      var pdf_file = exportWordAsPDF(
        inputFilePath,
        outputFilePath,
        sord.name +
          "¶" +
          KBF_nummer.toString() +
          "_FV_" +
          bezeichnung.toString(),
        KBF_nummer,
        bezeichnung
      );
      prnt(
        "PDF Export (" +
          KBF_nummer.toString() +
          "_Vorhaben_" +
          bezeichnung.toString() +
          ".pdf) erfolgreich!"
      );
    } catch (e) {
      prnt("PDF-Export ist fehlgeschlagen. Fehler: " + e.toString());
    }
    //__________

    //Navigation zum Ordner
    workspace.updateSordLists();
    workspace.getActiveView().refresh();
    workspace.gotoId(forschungsDokument.id);

    //__________
  } else if (node.name == "Formular weiterleiten per Mail") {
    //Basisdaten aus Verschlagwortung auslesen (via Konstanten aus lib_const.js)
    var KBF_nummer = indexDialog.getObjKeyValue(FORSCHUNGSVORHABEN_KBF_NUMMER);
    var bezeichnung = indexDialog.getObjKeyValue(
      FORSCHUNGSVORHABEN_BEZEICHNUNG
    );
    var FZG_nummer = indexDialog.getObjKeyValue(FORSCHUNGSVORHABEN_FZG_NR);
    var sachbearbeiter = indexDialog.getObjKeyValue(
      FORSCHUNGSVORHABEN_SACHBEARBEITER
    );
    var AL = indexDialog.getObjKeyValue(FORSCHUNGSVORHABEN_AL);
    var laufzeit = indexDialog.getObjKeyValue(FORSCHUNGSVORHABEN_LAUFZEIT);
    var anz_pruefkoerper = indexDialog.getObjKeyValue(
      FORSCHUNGSVORHABEN_PRUEFKOERPER
    );
    var leistung_werkstatt = indexDialog.getObjKeyValue(
      FORSCHUNGSVORHABEN_L_WERKSTATT
    );
    var leistung_prueffeld = indexDialog.getObjKeyValue(
      FORSCHUNGSVORHABEN_L_PRUEFFELD
    );
    var leistung_elabor = indexDialog.getObjKeyValue(
      FORSCHUNGSVORHABEN_L_ELABOR
    );
    var leistung_labor = indexDialog.getObjKeyValue(FORSCHUNGSVORHABEN_L_LABOR);

    var sordName = KBF_nummer.toString() + "_FV_" + bezeichnung.toString();

    var forschungsDokument = archive.getElementByArcpath(
      FORSCHUNGSVORHABEN_DIR_OPEN +
        "¶" +
        KBF_nummer.toString() +
        "_FV_" +
        bezeichnung.toString()
    );

    //Info
    var infoMessage = "<HTML>";
    infoMessage =
      infoMessage +
      "<p>Folgende Personen sollen das neue Forschungsvorhaben zur Kenntnis nehmen:</p>";
    infoMessage = infoMessage + "<ul>";
    infoMessage = infoMessage + "<li>Prof. Stahl</li>";
    infoMessage = infoMessage + "<li>alle AL-Sachbearbeiter</li>";
    infoMessage = infoMessage + "<li>Labor</li>";
    infoMessage = infoMessage + "<li>Reiner</li>";
    infoMessage = infoMessage + "<li>Werkstatt</li>";
    infoMessage = infoMessage + "<li>E-Labor</li>";
    infoMessage = infoMessage + "<li>Verwaltung</li>";
    infoMessage = infoMessage + "<li>Homepageverwalter</li>";
    infoMessage = infoMessage + "</ul>";
    infoMessage = infoMessage + "<BR>";
    infoMessage =
      infoMessage + "Für diese Personen gibt es einen E-Mail-Verteiler: ";
    infoMessage = infoMessage + "forschungsvorhaben.fzg@ed.tum.de" + "<BR>";
    infoMessage = infoMessage + "</HTML>";

    prnt(infoMessage);

    // Variablen erstellen
    var subject = "Neues Forschungsvorhaben '" + bezeichnung.toString() + "'";
    var emailText =
      "Sehr geehrte Kolleginnen und Kollegen,<BR>" +
      "<BR>" +
      "hiermit möchten wir Ihnen ein neues Forschungsvorhaben vorstellen.<BR>" +
      "<BR>" +
      "Öffnen Sie dazu das angehängte ELO-Lesezeichen, um das Forschungsvorhaben in ELO anzuzeigen.<BR>" +
      "<BR>" +
      "Mit freundlichen Grüßen,<BR>" +
      "Konny";

    //ECD Datei erstellen
    //Ein .ecd File ist ein ELO-Lesezeichen, das Anklicken eines solchen Files öffnet den Java Client und führt zum entsprechenden File
    var ecdObjectName = createECDLink(forschungsDokument); //legt ein ECD File im temp ordner ab, Achtung: Dateiname wird neu durch die Funktion festgelegt

    // Anhänge: ein .ecd File (eine ELO-Verlinkung) -> dieses soll in ELO erstellt, im tempDir ordner abgelegt werden und mit in den Anhang gepackt werden
    var attachments = [
      workspace.directories.tempDir + "\\" + ecdObjectName.toString() + ".ecd",
    ];

    //E-Mail an E-Mail-Verteiler senden
    mail(
      "forschungsvorhaben.fzg@ed.tum.de",
      subject,
      emailText,
      attachments,
      undefined
    );

    prnt("Outlook wird nun geoffnet...");
    prnt("Ein ELO-Lesezeichen (.ecd-Datei) wird an die E-Mail angehaengt.");
    prnt("Bitte Outlook-Mail überprüfen und dann manuell absenden!");
  }
}

/**
 * Füllt ein Word-Dokument mit den angegebenen Feldern und Einträgen und speichert es im ELO-Ordner.
 *
 * @param {string} file - Der Pfad zum Word-Dokument, das bearbeitet wird.
 * @param {Array} fieldnames - Die Feldnamen im Word-Dokument.
 * @param {Array} entries - Die Einträge, die in die entsprechenden Felder eingefügt werden.
 * @param {string} sordName - Der Name des ELO-Sords.
 * @param {string} KBF_nummer - Die KBF-Nummer des Forschungsvorhabens.
 * @param {string} bezeichnung - Die Bezeichnung des Forschungsvorhabens.
 * @returns {Object} - Das ELO-Dokument, das dem Ordner hinzugefügt wurde.
 */
function fillWordTemplate(
  file,
  fieldnames,
  entries,
  sordName,
  KBF_nummer,
  bezeichnung
) {
  ComThread.InitSTA();

  try {
    // Dokument oeffnen
    var word = new ActiveXComponent("Word.Application");
    var documents = Dispatch.get(word, "Documents").toDispatch();
    var doc = Dispatch.call(documents, "Open", file).toDispatch();

    //Error message vorbereiten
    var err = "";
    var counter71 = 0;

    for (var i = 0; i < fieldnames.length; i++) {
      try {
        var field = Dispatch.call(
          doc,
          "FormFields",
          fieldnames[i]
        ).toDispatch();

        var fieldType = Dispatch.get(field, "Type"); //Feldtyp auslesen, Type 71 steht für CheckBox
        if (fieldType == 71) {
          counter71++;
          var checkBox = Dispatch.get(field, "CheckBox").toDispatch();
          Dispatch.put(checkBox, "Checked", true);
        }

        Dispatch.put(field, "Result", entries[i]);
      } catch (e) {
        err += fieldnames[i] + "   " + e + "<br>";
      }
    }
  } catch (e) {
    workspace.showInfoBox("Error", e + "<br>Nested Errors:" + err);
  } finally {
    // Dokument speichern und schließen
    Dispatch.call(doc, "SaveAs", file);
    var aw = Dispatch.get(doc, "ActiveWindow").toDispatch();
    Dispatch.call(aw, "Close");

    //Dokument in den richtigen ELO Ordner einhaengen
    var forschungsDokument = archive.getElementByArcpath(
      FORSCHUNGSVORHABEN_DIR_OPEN + "¶" + sordName.toString()
    );

    var prepdoc = forschungsDokument.prepareDocument(0);
    var docpath =
      workspace.directories.tempDir +
      "\\" +
      KBF_nummer.toString() +
      "_Vorhaben-" +
      bezeichnung.toString() +
      ".docx";
    var elodoc = forschungsDokument.addDocument(prepdoc, docpath);
    workspace.updateSordLists();
    prnt(
      "Word-Formular wurde ausgefüllt und wie folgt abgelegt: 216_Forschungsvorhaben" +
        "\\" +
        sordName.toString() +
        "\\" +
        KBF_nummer.toString() +
        "_Vorhaben-" +
        bezeichnung.toString() +
        ".docx"
    );
  }
  ComThread.Release();

  return elodoc;
}

/**
 * Exportiert ein Word-Dokument als PDF und fügt es dem entsprechenden ELO-Ordner hinzu.
 *
 * @param {string} inputFilePath - Der Pfad zum Eingabe-Word-Dokument.
 * @param {string} outputFilePath - Der Pfad zum exportierten PDF-Dokument.
 * @param {string} sordName - Der Name des ELO-Sords.
 * @param {string} KBF_nummer - Die KBF-Nummer des Forschungsvorhabens.
 * @param {string} bezeichnung - Die Bezeichnung des Forschungsvorhabens.
 * @returns {Object} - Das ELO-Dokument, das dem Ordner hinzugefügt wurde.
 */
function exportWordAsPDF(
  inputFilePath,
  outputFilePath,
  sordName,
  KBF_nummer,
  bezeichnung
) {
  ComThread.InitSTA();

  try {
    // Word-Anwendung öffnen
    var word = new ActiveXComponent("Word.Application");
    var documents = Dispatch.get(word, "Documents").toDispatch();

    // Dokument öffnen
    var doc = Dispatch.call(documents, "Open", inputFilePath).toDispatch();

    // Dokument als PDF exportieren
    Dispatch.call(doc, "ExportAsFixedFormat", outputFilePath, 17); // Der zweite Parameter 17 entspricht dem PDF-Dateiformat

    // Dokument schließen
    var aw = Dispatch.get(doc, "ActiveWindow").toDispatch();
    Dispatch.call(aw, "Close");

    /*  prnt(
      "Das Word-Dokument wurde erfolgreich als PDF exportiert. Die PDF-Version befindet sich unter: " +
        outputFilePath
    ); */

    //Dokument in den richtigen ELO Ordner einhaengen
    var forschungsDokument = archive.getElementByArcpath(
      FORSCHUNGSVORHABEN_DIR_OPEN + "¶" + sordName.toString()
    );

    var prepdoc = forschungsDokument.prepareDocument(0);
    var docpath =
      workspace.directories.tempDir +
      "\\" +
      KBF_nummer.toString() +
      "_Vorhaben-" +
      bezeichnung.toString() +
      ".pdf";
    var elodoc = forschungsDokument.addDocument(prepdoc, docpath);
    workspace.updateSordLists();
    prnt(
      "Word-Formular wurde als PDF exportiert und wie folgt abgelegt: 216_Forschungsvorhaben" +
        "\\" +
        sordName.toString() +
        "\\" +
        KBF_nummer.toString() +
        "_Vorhaben-" +
        bezeichnung.toString() +
        ".pdf"
    );
  } catch (e) {
    workspace.showInfoBox("Error", e);
    prnt("Fehler beim PDF-Export!");
  } finally {
    ComThread.Release();
  }
  return elodoc;
}

/**
 * Erstellt ein ELO-Lesezeichen (ECD-Link) für ein Element und legt ein Lesezeichen an.
 * Ein ELO-Lesezeichen ermöglicht durch Doppelklick den direkten Zugriff auf das entsprechende Element im ELO System.
 *
 * @param {Object} item - Das Element, für das das ELO-Lesezeichen erstellt wird.
 * @returns {string} - Der erstellte ELO-Lesezeichen-Dateiname.
 */
function createECDLink(item) {
  // Extrahiert den Archivnamen aus dem Archiv-Objekt.
  var arcName = archive.archiveName;

  // Extrahiert die Objekt-ID des Elements.
  var objId = item.id;

  // Erstellt einen bereinigten Dateinamen für das ELO-Lesezeichen.
  var ecdName = buildName(item);

  // Kombiniert den temporären Dateipfad mit dem bereinigten Dateinamen und der Erweiterung ".ecd".
  var tempName = workspace.directories.tempDir + "\\" + ecdName + ".ecd";

  // Erstellt den Inhalt des ELO-Lesezeichens mit Archivnamen und Objekt-ID.
  var t = "EP\r\nA" + arcName + "\r\nI" + objId + "\r\n";

  // Erstellt eine Datei im temporären ELO Ordner
  var file = new Packages.java.io.File(tempName);

  // Schreibt den Inhalt in die Datei.
  Packages.org.apache.commons.io.FileUtils.writeStringToFile(file, t);

  // Setzt eine Feedback-Nachricht für das Workspace-Objekt.
  workspace.setFeedbackMessage("ELO-Lesezeichen angelegt");

  // Gibt den erstellten ELO-Lesezeichen-Dateinamen zurück.
  return ecdName;
}

/**
 * Erstellt einen Dateinamen und entfernt unzulaessige Sonderzeichen.
 *
 * @param {Object} item - Die Datei, für die der Dateiname erstellt wird.
 * @returns {string} - Der bereinigte Dateiname ohne Sonderzeichen.
 */
function buildName(item) {
  // Konvertiert den Namen in einen String.
  var name = String(item.name);

  // Entfernt unzulässige Sonderzeichen im Dateinamen.
  var cleanName = name.replace(/\<|\>|\:|\\|\/|\*|\?|\||\~/g, "_");

  // Gibt den bereinigten Dateinamen zurück.
  return cleanName;
}
