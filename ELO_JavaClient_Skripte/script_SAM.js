function eloExpandRibbon() {
  createRibbon();
  newButton(BUTTON_SAM, SAMSicherheitsunterweisung);
}

//------- Funktion SAM Sicherheitsunterweisung goes here ----------
function SAMSicherheitsunterweisung() {

// --------- Datum ----------	
  const zeit = new Date();
  var jahr = zeit.getFullYear();
  jahr = jahr.toString();
  jahr = jahr.substr(2,2);
  var monat = zeit.getMonth() + 1;
  monat = monat.toString();
  if (monat.length < 2) monat = "0" + monat;
  var tag = zeit.getDate();
  tag = tag.toString();
  if (tag.length < 2) tag = "0" + tag;


//------- Sord erstellen -------	
var root = archive.getElementByArcpath(SAM_STUDI_DIR_OPEN);
  var editInfo = ixc.createSord(root.id, SAM_STUDI_MASK, EditInfoC.mbAll);
  var sord = editInfo.sord;
sord.name = LIT_MASK_DUMMY_NAME; 

//------- Maske für Dateneingabe -------	
//IndexDialog anzeigen
  if (!indexDialog.editSord(sord, true, SAM_STUDI_SORD_TITLE)) {
      return;
  }
var surname = indexDialog.getObjKeyValue(SAM_STUDI_MASK_NACHNAME);
if (surname == '') {
      return;
  }
var first_name = indexDialog.getObjKeyValue(SAM_STUDI_MASK_VORNAME);
if (first_name == '') {
      return;
  }

//------- Datei SAM Unterweisung erstellen -------
workspace.updateSordLists();	

//  ELO-Benutzername aus LDAP auf benötigtes Format bringen
var name = workspace.getUserName();
  var nameformatted = getNameFormatted(name,LST_NAME_FST_NAME);
var dateiname = "\\SAM_Unterweisung_" + surname + "_" + first_name + "_" + jahr + monat + tag + ".pdf"


try {
  
  var litElem = archive.getElementByArcpath(SAM_STUDI_DIR_OPEN + "¶" + nameformatted + "¶" + surname + "_" + first_name );
  // PDF aus Templates kopieren
  var sam_studi_template = archive.getElementByArcpath(SAM_STUDI_DIR_TEMPLATE);
  var docpath = workspace.directories.tempDir + dateiname;	
  
  // PDF ausfüllen 
  openRegistrationForm(sam_studi_template, sord, docpath, name);
  
  // Pdf in Studi Ordner einhängen
  var prepdoc = litElem.prepareDocument(0);
  var elodoc = litElem.addDocument(prepdoc, docpath);
  workspace.updateSordLists();

} catch (error) {

  // ELO Ordner des Studenten anlegen
  var sam_studi_dir = archive.getElementByArcpath(SAM_STUDI_DIR_OPEN + "¶" + nameformatted );
  sord.name = surname + "_" + first_name ; //sord.name = sam_studi_num;
  
  // Ordner in Archivbaum einhängen
  sord.parentId = sam_studi_dir.id;
  var dest = archive.getElement(sord.parentId);
  var litElem = dest.addStructure(sord);
  workspace.updateSordLists();

  // PDF aus Templates kopieren
  var sam_studi_template = archive.getElementByArcpath(SAM_STUDI_DIR_TEMPLATE);
  var docpath = workspace.directories.tempDir + dateiname;	
  
  // PDF ausfüllen 
  openRegistrationForm(sam_studi_template, sord, docpath, name);

  // Pdf in Studi Ordner einhängen
  var prepdoc = litElem.prepareDocument(0);
  var elodoc = litElem.addDocument(prepdoc, docpath);
  workspace.updateSordLists();
}

//------- Meldung an User -------
 workspace.showInfoBox("Unterlagen für SAM Sicherheitsunterweisung angelegt", "Die Sicherheitsunterweisung Ihres Studierenden wurde im Ordner " + SAM_STUDI_DIR_OPEN + "¶" + nameformatted + "¶" + surname + "_" + first_name + " angelegt.<br><br>Bitte fuellen Sie diese nachfolgend in der PDF (das sich automatisch oeffnet) aus und checken das Dokument danach ein. Leiten Sie dann den entsprechenden Workflow an den Sicherheitsbeauftragten weiter.");
 workspace.updateSordLists();
 workspace.getActiveView().refresh();
 workspace.gotoId(litElem.id);
 
 // PDF Auschecken
 elodoc.checkOut();
 
 // PDF-File oeffnen
 utils.editFile(checkout.getLastDocument().getDocumentFile());
  
 // File nicht einchecken (muss vom Nutzer nach Eingabe gemacht werden)
   //coDoc.checkIn("0", "Test");
 
 
//------- Start des Workflows -------
  var wf = ixc.checkoutWorkFlow("SAM", WFTypeC.TEMPLATE, WFDiagramC.mbAll, LockC.NO);
  wf.type = WFTypeC.ACTIVE;
  //workspace.showInfoBox("SAM", BASAMA_BENOT_INFO1);
//workspace.showInfoBox("SAM", BASAMA_BENOT_WF_CONFIRM_Ot_sen); 
wf.id = -1;
wf.objId = elodoc.getId(); //wf.objId = litElem.getId();
wf.name = "SAM";
  ixc.checkinWorkFlow(wf, WFDiagramC.mbAll, LockC.NO);
}



function openRegistrationForm(elopdf, sord, storepath, name) {
  // Pdf in pdfbox öffnen
  var pdf = PDDocument.load(new File(elopdf.getFile()));
  var pdfCat = pdf.getDocumentCatalog();
  var acroForm = pdfCat.getAcroForm();
  if (acroForm == null) {
      return;
  }

  // Felder ausfüllen für Namen
var vorname_user = getNameFormatted(name,FST_NAME_ONLY);
var nachname_user = getNameFormatted(name,LST_NAME_ONLY);
var mitarbeiter_kuerzel = [getUserEntry(ELO_USER_SHORT)]
  acroForm.getField(SAM_STUDI_REG_FORM_STUD_SURNAME).setValue(getIndexValueByName(sord, SAM_STUDI_MASK_NACHNAME));
  acroForm.getField(SAM_STUDI_REG_FORM_STUD_FSTNAME).setValue(getIndexValueByName(sord, SAM_STUDI_MASK_VORNAME));
  acroForm.getField(SAM_STUDI_REG_FORM_ASS_SURNAME).setValue(nachname_user);
  acroForm.getField(SAM_STUDI_REG_FORM_ASS_FSTNAME).setValue(vorname_user);
  acroForm.getField("KuerzelAssistent").setValue(mitarbeiter_kuerzel);
   
  // Datum ausfüllen
  var cal = Calendar.getInstance();
  var date = formatDateAsString(cal.get(Calendar.DAY_OF_MONTH), cal.get(Calendar.MONTH) + 1, cal.get(Calendar.YEAR), false);
  acroForm.getField(SAM_STUDI_REG_FORM_DATUM).setValue(date);

  // Speichern
  pdf.save(storepath);
  pdf.close();
}




// Funktionen innerhalb des Workflows
function eloFlowConfirmDialogOkStart() {
  var dialog = dialogs.flowConfirmDialog;
  var successors = dialog.selectedSuccessors;
  retval = 0;
  if (successors == null)
      return;

  var sucName = successors[0].getName().split("");
  var sucName = successors[0].getName();
  var dialog = dialogs.flowConfirmDialog;
  var item = tasksViews.getTasksViewForWorkflow(dialog.getFlowId()).getWorkflow(dialog.getFlowId()).getArchiveElement();
  var sord = item.loadSord();
  var uname = sord.getOwnerName();

//  Datei-Besitzername(=Antragsteller) aus LDAP auf benötigtes Format bringen
  var unameformatted = getNameFormatted(uname,LST_NAME_FST_NAME);  
  var date = sord.getIDateIso()

// Überprüfen ob alle benoetigten Felder der Sicherheitsunterweisung.pdf ausgefuellt sind, sonst nicht weiterleiten
SAM_WORKFLOWTRIGGER_CHECKPDF = "Eintragung in TUMOnline"
if (sucName.equals(SAM_WORKFLOWTRIGGER_CHECKPDF)){
var fieldnames_read_last_line = ["LRZ-Kennung", "Startdatum", "Nachname_Stud", "Vorname_Stud", "Nachname_Assist", "Vorname_Assist"];
var entries_read_last_line = [];
  var verbose = true;
var elopdf = item;

readPDFFields(elopdf, fieldnames_read_last_line, entries_read_last_line, verbose)

// Wenn ein Feld leer ist, dann benachrichtigen, dass erneut auszufüllen ist.
 if ((entries_read_last_line[0].isEmpty()) || (entries_read_last_line[1].isEmpty()) || (entries_read_last_line[2].isEmpty()) || (entries_read_last_line[3].isEmpty()) || (entries_read_last_line[4].isEmpty()) || (entries_read_last_line[5].isEmpty())) {
      workspace.showInfoBox("SAM-Sicherheitsunterweisung", "Das Antragsformular ist noch nicht vollständig ausgefüllt! Bitte alle Namen, LRZ-Kennung und Startdatum ausfüllen. Anschließend den Workflow erneut weiterleiten.");
      return -1;	
 }
 
 if (elopdf.isLocked()){
      workspace.showInfoBox("SAM-Sicherheitsunterweisung", "Das Antragsformular ist noch nicht eingechecket. Bitte einchecken! Anschließend den Workflow erneut weiterleiten.");
      return -1;
 }
    
 }
}




// Felder aus PDF Datei auslesen
  function readPDFFields(elopdf, fieldnames, entries, verbose) {
  if (verbose == undefined) { verbose = true; }

  // Pdf öffnen
  var pdf = PDDocument.load(elopdf.getFile());
  var pdfCat = pdf.getDocumentCatalog();
  var temppath = utils.getUniqueFile(workspace.directories.tempDir, "Urlaubsblatt_neuer_Antrag.pdf");
  var acroForm = pdfCat.getAcroForm();
  if (acroForm == null) {
      return false;
  }

  // Felder befüllen
  for (var i = 0; i < fieldnames.length; i++) {
      try {
          entries[i] = acroForm.getField(fieldnames[i]).getValueAsString();
      } catch (e) {
          if (verbose) { workspace.showInfoBox("", e); }
      }
  }
}

//@include lib_utility
//@include lib_const
//@include lib_buttons