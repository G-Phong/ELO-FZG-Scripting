//@include lib_utility
//@include lib_const
//@include lib_buttons

//------- Button im entsprechenden Ribbon-Band hinzufÃ¼gen ----------
function eloExpandRibbon() {
    createRibbon();
    newButton(BUTTON_FORSCHUNGSVORHABEN, Forschungsvorhaben);
  }
  
  //------- Funktion Forschungsvorhaben ----------
  function Forschungsvorhaben() {
    // Erstellen eines Sords als Vorlage fÃ¼r den Workflow
    var root = archive.getElementByArcpath(FORSCHUNGSVORHABEN_DIR_OPEN);
    var editInfo = ixc.createSord(
      root.id,
      FORSCHUNGSVORHABEN_MASK,
      EditInfoC.mbAll
    );
    var sord = editInfo.sord;
    sord.name = "Forschungsvorhaben 01";
  
    // Ordner in Archivbaum einhÃ¤ngen
    var dest = archive.getElement(root.id);
    //dest.addStructure(sord); //Ordner hinzufügen, testweise auskommentiert
    // workspace.showInfoBox("destination",dest.toString());
  
    //------- Datei Forschungsvorhaben (Word) erstellen -------
    workspace.updateSordLists();
  
    var litElem = archive.getElementByArcpath(
      FORSCHUNGSVORHABEN_DIR_OPEN + "¶Forschungsvorhaben 01"
    );
  
    /*   var rootTemplate = archive.getElementByArcpath(FORSCHUNGSVORHABEN_DIR_TEMPLATE);
    var templatePath = archive.getElement(rootTemplate.id);
    workspace.showInfoBox("templatePath", templatePath.toString()); */
    try {
      var prepdoc = litElem.prepareDocument(0);
  
      var docPath =
        workspace.directories.tempDir + "\\00_Vorhaben-Kurzbeschreibung.docx";
      //var docPath = templatePath + "\\00_Vorhaben-Kurzbeschreibung.docx";
  
      var elodoc = litElem.addDocument(prepdoc, docPath);
    } catch (errror) {
      workspace.showInfoBox(
        "Fehler",
        "Dokument 00_Vorhaben-Kurzbeschreibung.docx liegt nicht im temporären Ordner bereit (checkout)!"
      );
    }
  
    /*  var string = String(getIndexValueByName(sord, FORSCHUNGSVORHABEN_KBF_NUMMER)); //so liest man werte aus dem indexserver aus (maske im admin console)
     */
    /*   workspace.showInfoBox("Fehler", string); */
  
    // Word in Forschungsvorhaben Ordner einhängen
    // Word-Dok wird aus dem Template Ordner reinkopiert
  
    //------- Maske für Dateneingabe -------
    //IndexDialog anzeigen
    if (!indexDialog.editSord(sord, true, "Basisdaten eingeben")) {
      return;
    }
  
    var KBF_nummer = indexDialog.getObjKeyValue(FORSCHUNGSVORHABEN_KBF_NUMMER);
    var bezeichnung = indexDialog.getObjKeyValue("FORSCH_BEZEICHNUNG");
    var FZG_nummer = indexDialog.getObjKeyValue("FORSCH_FZG_NR");
    var sachbearbeiter = indexDialog.getObjKeyValue("FORSCH_SACHBEARBEITER");
    var AL = indexDialog.getObjKeyValue("FORSCH_AL");
    var laufzeit = indexDialog.getObjKeyValue("FORSCH_LAUFZEIT");
  
    //KBF Nummer darf nicht leer sein -> Test
    if (KBF_nummer == "") {
      return;
    } else {
      workspace.showInfoBox(
        "Maskeneingabe",
        "Eingegebener Wert für KBF-Nummer: " + KBF_nummer
      );
    }
  
    workspace.updateSordLists();
  
    var workflow = ixc.checkoutWorkFlow(
      "Forschungsvorhaben_confirmDialog",
      WFTypeC.TEMPLATE,
      WFDiagramC.mbAll,
      LockC.NO
    );
  
    workflow.type = WFTypeC.ACTIVE;
    workflow.id = -1;
    workflow.objId = litElem.getId();
    workflow.name = "Forschungsvorhaben";
    ixc.checkinWorkFlow(workflow, WFDiagramC.mbAll, LockC.NO);
  }
  
  // Funktionen innerhalb des Workflows
  function eloFlowConfirmDialogStart() {
    //Diese Funktion wird bei JEDER WEITERLEITUNG eines Knotens ausgeführt. -> Problematisch!
    //Kann ich das an eine Bedingung knüpfen? Bei jedem Knoten soll was anderes gemacht werden. if(this.Knoten == "Knoten1") ? so in die Richtung?
    //Ja, das geht mit Successors!
  
    var dialog = dialogs.flowConfirmDialog;
    var node = dialog.currentNode;
  
    if (node.name == "PDF-Version des Laufzettels an Kornelia") {
      workspace.showInfoBox(
        "confirmDialog-Skriptfunktion wurde autom. aufgerufen",
        "Richtiger Knoten -> Schreibe nun die Daten in das Word-Template"
      );
  
      // prnt("Weiterleitung");
  
      /*   workspace.showInfoBox(
      "Weiterleitung",
      "eloFlowConfirmDialogStart -> Jetzt könnten die eingegebenen Daten der Verschlagwortungsmaske ausgelesen und in die Word-Vorlage reinkopiert werden!"
    ); */
  
      /*   workspace.showInfoBox(
        "Bei welchem Workflow Knoten sind wir?",
        "Name des aktuellen Knotens: " + node.name
      ); */
  
      //Prüfung bei welchem Knoten der Workflow gerade ist
      if (node.name == "Eingabe der Basisdaten in den Laufzettel") {
        workspace.showInfoBox(
          "ELO-Knoten",
          "(1) Der aktuelle Knoten ist " + node.name
        );
      } else if (node.name == "Ergänzung von Grobdaten und Projektbeschreibung") {
        workspace.showInfoBox(
          "ELO-Knoten",
          "(2) Der aktuelle Knoten ist " + node.name
        );
      } else {
        workspace.showInfoBox(
          "ELO-Knoten",
          "(3) Der aktuelle Knoten ist " + node.name
        );
      }
  
      //_______DATEN IN KNOTEN SCHREIBEN_________
      var KBF_nummer = indexDialog.getObjKeyValue(FORSCHUNGSVORHABEN_KBF_NUMMER);
      var bezeichnung = indexDialog.getObjKeyValue("FORSCH_BEZEICHNUNG");
      var FZG_nummer = indexDialog.getObjKeyValue("FORSCH_FZG_NR");
      var sachbearbeiter = indexDialog.getObjKeyValue("FORSCH_SACHBEARBEITER");
      var AL = indexDialog.getObjKeyValue("FORSCH_AL");
      var laufzeit = indexDialog.getObjKeyValue("FORSCH_LAUFZEIT");
  
      var infoMessage = "<HTML>";
      infoMessage =
        infoMessage +
        "Die Word-Vorlage wird nun mit folgenden Daten gefüllt: <BR>";
      infoMessage = infoMessage + "KBF-Nummer: " + KBF_nummer + "<BR>";
      infoMessage = infoMessage + "bezeichnung:  " + bezeichnung + "<BR>";
      infoMessage = infoMessage + "FZG_nummer:  " + FZG_nummer + "<BR>";
      infoMessage = infoMessage + "sachbearbeiter:  " + sachbearbeiter + "<BR>";
      infoMessage = infoMessage + "AL: " + AL + "<BR>";
      infoMessage = infoMessage + "laufzeit: " + laufzeit + "<BR>";
  
      workspace.showInfoBox("ELO-Verschlagwortungsdaten", infoMessage);
  
      //TODO: Hier Word-Dok füllen über COM-Schnittstelle
      /* initCOM(
        "C:\x20Users\x20giaph\x20AppData\x20Roaming\x20ELO Digital Office\x20FZG_test\x20405\x20checkout\x2000_Vorhaben-Kurzbeschreibung.docx"
      );
   */
      var item = workspace.getActiveView().getFirstSelected();
      var file = item.file;
      var litElem = archive.getElementByArcpath(
        FORSCHUNGSVORHABEN_DIR_OPEN + "¶Forschungsvorhaben 01"
      );
      file = litElem.prepareDocument(0);
  
      ComThread.InitSTA();
      try {
        var word = new ActiveXComponent("Word.Application");
        var documents = Dispatch.get(word, "Documents").toDispatch();
        var doc = Dispatch.call(documents, "Open", file).toDispatch();
  
        /* var name = readProperty(doc, "Name"); */
        var aw = Dispatch.get(doc, "ActiveWindow").toDispatch();
        Dispatch.call(aw, "Close");
        item.name = item.name + " : " + name;
        prnt(item.name);
        item.saveSord();
      } catch (e) {
        prnt(e);
        log.info("Error reading word properties (Vorlage): " + e);
        prnt("error trying to access the COM bridge");
      }
      ComThread.Release();
  
      var xl = new ActiveXComponent("Word.Application");
      Dispatch.put(xl,"Visible",1);
  
      //________ENDE: DATEN IN KNOTEN SCHREIBEN______
  
      //________Personalisierter Dialog___________
      var panel = dialog.addGridPanel(2, 3);
      var ltxt = "hier ist ein eingabefeld: ";
      var label = panel.addLabel(1, 1, 2, ltxt);
      label.fontSize = 16;
      label.bold = true;
      label.setColor(200, 30, 30);
  
      inputField = panel.addTextField(1, 2, 1);
      var desc = "<html> Weiterleitung nach Übertragen in die Word-Vorlage. <br>";
      var remark = panel.addLabel(1, 3, 2, desc);
      remark.fontSize = 12;
      //________ENDE: Personalisierter Dialog___________
    } else {
      return;
    }
  }
  
  /* function eloFlowConfirmDialogOkStart() {
    workspace.showInfoBox("Weiterleitung", "eloFlowConfirmDialogOkStart");
  }
  
  function eloFlowConfirmDialogEnd() {
    workspace.showInfoBox("Weiterleitung", "eloFlowConfirmDialogEnd");
  } */
  