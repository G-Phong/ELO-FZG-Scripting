//@include lib_utility
//@include lib_const
//@include lib_buttons

function eloExpandRibbon() {
  createRibbon();
  newButton(BUTTON_BASAMA_Benotung, BenotungBASAMA);
}

function BenotungBASAMA() {
  prnt(
    "Studienarbeit wird benotet, richtiges File muss vorher ausgewählt sein!"
  );

  // Felder auslesen
  // Ausgewaehlte Anmeldung
  // Ausgewaehltes Item
  var item = workspace.activeView.firstSelected;
  // Abfangen, falls nichts ausgewaehlt
  if (item == null) {
    workspace.showInfoBox(
      "Fehler",
      "Bitte Beurteilung der BASAMA auswaehlen, damit die Benotung weitergeleitet werden kann."
    );
    return;
  }
  // Abfangen, wenn Verschlagwortungsmaske nicht Anmeldung BASAMA
  var sord = item.loadSord();
  if (sord.getMask() != MASK_LITERATURE) {
    workspace.showInfoBox(
      "Fehler",
      "Bitte Beurteilung der BASAMA auswaehlen, damit die Benotung weitergeleitet werden kann."
    );
    return;
  }

  //Workflow "BASAMA_Anmeldung_EDportal" starten
  var workflow = ixc.checkoutWorkFlow(
    "BASAMA_Benotung_EDportal",
    WFTypeC.TEMPLATE,
    WFDiagramC.mbAll,
    LockC.NO
  );
  workflow.type = WFTypeC.ACTIVE;
  workflow.id = -1;
  workflow.objId = 15350; //TODO: ID des Benotungsdokuments angeben
  ixc.checkinWorkFlow(workflow, WFDiagramC.mbAll, LockC.NO);
}

function eloFlowConfirmDialogStart() {
  var dialog = dialogs.flowConfirmDialog;
  var node = dialog.currentNode;

  if (node.name == "Bewertung der Studienarbeit prüfen") {
    //TODO
    prnt("Selbständig Bewertung prüfen");
  }

  if (node.name == "Kenntnisnahme der Note durch Prof. Stahl") {
    //TODO
    prnt("TODO: Automatisierte E-Mail an Prof. Stahl");
  }

  if (node.name == "Eintragung der Note in das ED-Portal") {
    //TODO
    prnt("User soll nun auf das ED-Portal gehen.");
  }
}
