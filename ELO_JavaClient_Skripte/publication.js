// ELO
importPackage(Packages.de.elo.ix.jscript);
importPackage(Packages.de.elo.client);
importPackage(Packages.de.elo.client);
// .NET-Framework / COM-Schnittstelle
importPackage(Packages.de.jacob.com);
importPackage(Packages.de.jacob.activeX);
importPackage(Packages.de.ms.activeX);
importPackage(Packages.de.ms.com);
// PDF-Utility
importPackage(Packages.org.apache.pdfbox.exceptions);
importPackage(Packages.org.apache.pdfbox.pdmodel);
importPackage(Packages.org.apache.pdfbox.pdmodel.edit);
importPackage(Packages.org.apache.pdfbox.pdmodel.font);
importPackage(Packages.org.apache.pdfbox.pdmodel.interactive.form);
importPackage(Packages.org.apache.pdfbox.cos);
// Java Lib
importPackage(Packages.java.io);
importPackage(Packages.java.util);
importPackage(Packages.java.nio.file);
importPackage(Packages.java.lang);
importPackage(Packages.java.net);

function eloExpandRibbon() {
    createRibbon();

    newButton(BUTTON_NEW_STUDENT_WORK, newThesis);
    newButton(BUTTON_GRADE_STUDENT_WORK, gradeThesis);
    newButton(BUTTON_NEW_PUBLICATION,
		function () {
			var newElem = newIndexedFolder(INDEXED_FOLDER_PUBLICATION, false);
			insertPublicationTemplate(newElem);
		}
	);
    newButton(BUTTON_NEW_ARTICLE, function () { newIndexedFolder(INDEXED_FOLDER_ARTICLE, false); });
    newButton(BUTTON_NEW_PRESENTATION, function () { newIndexedFolder(INDEXED_FOLDER_PRESENTATION, false); });
    newButton(BUTTON_NEW_DISS, function () { newIndexedFolder(INDEXED_FOLDER_DISS, false); });
	newButton(BUTTON_RESEARCH_PROPOSAL,
		function () {
			var newElem = newIndexedFolder(INDEXED_FOLDER_RESEARCH_PROPOSAL, false);
			insertControllingTemplate(newElem);
		}
	);
    newButton(BUTTON_NEW_RESEARCH_REPORT,
        function () {
            var newElem = newIndexedFolder(INDEXED_FOLDER_RESEARCH_REPORT, true);
            createResearchReportSubFolderStructure(newElem);
        }
	);
}

////////////////////////////////////////////////////////////////////////////////////

/**
 * 
 * @param {Date} announced
 *
 * Output : Datum 'announced' + 6 Monate - 1 Tag
 *
 * Berechnet anhand dem Datum 'Themenbekanntgabe' das Datum 'Abgabefrist' einer Studienarbeit
 */
function calcDeadline(announced) {
    var date = announced.split("");

    var oldYear = parseInt(date[0] + date[1] + date[2] + date[3], 10);
    var oldMonth = parseInt(date[4] + date[5], 10);
    var oldDay = parseInt(date[6] + date[7], 10);

    // Berechnung Tag
    var newDay = oldDay;//- 1;

    // Berechnung Monat
    var newMonth = (oldMonth + 6) % 12;
    if (newMonth == 0) newMonth = 12;

    // ggf. Berechnung Jahr
    var newYear = oldYear + ((oldMonth > newMonth) ? 1 : 0);

    // ggf. Korrektur Tag
    if ((newDay > 28 && newMonth == 2 && newYear % 4 != 0) |
        (newDay > 29 && newMonth == 2 && newYear % 4 == 0) |
        (newDay > 30 && (newMonth == 4 || newMonth == 6 || newMonth == 9 || newMonth == 11)) |
        (newDay > 31 && (newMonth == 1 || newMonth == 3 || newMonth == 5 || newMonth == 7 || newMonth == 8 || newMonth == 10 || newMonth == 12))) { newDay = 1; newMonth++; }

    return formatDateAsString(newDay, newMonth, newYear, false);
}

/**
 * Erstellt eine neue Studienarbeit
 *
 * Der IndexDialog wird geöffnet um initiale Daten einzutragen
 * Das Anmeldeformular wird ausgefüllt und abgelegt
 * Es werden E-Mails an den Studenten, sowie das Prüfungsamt erstellt
 */
function newThesis() {
    // Neuen Ordner anlegen
    var root = archive.getElementByArcpath(ROOT_LITERATURE);
    var editInfo = ixc.createSord(root.id, MASK_LITERATURE, EditInfoC.mbAll);
    var sord = editInfo.sord;
    // Platzhalter für die Publikationsnummer, welche erst später eingefügt werden kann
    // Steuert das dynamische Ausgrauen im IndexDialog
    sord.name = LIT_MASK_DUMMY_NAME;

    //IndexDialog anzeigen
    if (!indexDialog.editSord(sord, true, LITERATURE_SORD_TITLE)) {
        return;
    }
    var type = indexDialog.getObjKeyValue(LIT_MASK_TYPE);
    if (type == '') {
        return;
    }
	
	// Einschub bei diversen Studienarbeiten (legt nur nummerierten Ordner an und springt dann aus Funktion)
	if (type == LIT_TYPE_DIV_STUD['short'] || type == LIT_TYPE_DIV_STUD['full']) {
		// Auswahl von Konstanten-Set für den weiteren Prozess (Ordnerpfad usw.)
		var newFolderInfo = LitTypeTupleByName(type, "XYZ");
		// Publikationsnummer ermitteln
		var PubNumber = getNewPubNumber(newFolderInfo).number;
		sord.setName(newFolderInfo.short + "-" + PubNumber);
		// Ordner in Archivbaum einhängen
		sord.parentId = archive.getElementByArcpath(ROOT_LITERATURE + newFolderInfo.folder).id;
		var dest = archive.getElement(sord.parentId);
		var litElem = dest.addStructure(sord);
		// Speichern, Meldung an User
		workspace.updateSordLists();
		workspace.gotoId(litElem.id);
		// aus Funktion springen
		return;
	}

    // Auswahl von Konstanten-Set für den weiteren Prozess (Ordnerpfad usw.)
    var newFolderInfo = LitTypeTupleByName(type, " ");
    // Publikationsnummer ermitteln
    var PubNumber = getNewPubNumber(newFolderInfo).number;
    sord.setName(newFolderInfo.short + "-" + PubNumber);
    // Abgabefrist berechnen
    setIndexValueByName(sord, LIT_MASK_DEADLINE, calcDeadline(indexDialog.getObjKeyValue("LIT_ANNOUNCED")));
    // Ordner in Archivbaum einhängen
    sord.parentId = archive.getElementByArcpath(ROOT_LITERATURE + newFolderInfo.folder).id;
    var dest = archive.getElement(sord.parentId);
    var litElem = dest.addStructure(sord);

    // Speichern, Meldung an User
    workspace.updateSordLists();
    workspace.gotoId(litElem.id);
    workspace.showInfoBox("Studienarbeit anmelden", NEW_THESIS_INFO1);

    // Anmeldung.pdf kopieren
    var registration = archive.getElementByArcpath(DIR_TEMPLATES_THESIS + newFolderInfo.folder + LIT_TEMPLATE + newFolderInfo.regFormPath);
    var docPath = workspace.directories.checkOutDir + "\\Anmeldung.pdf";
    // Pdf ausfüllen 
    openRegistrationForm(registration, sord, docPath);
    // Pdf einhängen
    var prepdoc = litElem.prepareDocument(0);
    var elodoc = litElem.addDocument(prepdoc, docPath);
    // Speichern, öffnen
    workspace.updateSordLists();
    elodoc.open();
   
    // Vorlage in Abhaenigkeit der Studienarbeit auswaehlen
    var thesis_template = DIR_TEMPLATES_THESIS;
    switch (true) {
        case type == LIT_TYPE_BA['full']:
            thesis_template += LIT_TYPE_BA['folder'] + LIT_TEMPLATE + LIT_TYPE_BA['regTemplate'];
            var thesis_name = "BA-".concat(PubNumber);
            break;
        case type == LIT_TYPE_MA['full']:
            thesis_template += LIT_TYPE_MA['folder'] + LIT_TEMPLATE + LIT_TYPE_MA['regTemplate'];
            var thesis_name = "MA-".concat(PubNumber);
            break;
        case type == LIT_TYPE_SA['full']:
            thesis_template += LIT_TYPE_SA['folder'] + LIT_TEMPLATE + LIT_TYPE_SA['regTemplate'];
            var thesis_name = "SA-".concat(PubNumber);
            break;
        default:
            workspace.showInfoBox("Fehler", "701");
            return;
    }
    thesis_template = archive.getElementByArcpath(thesis_template);    

    // Vorlage in temporaeres Verzeichnis kopieren   
    var tmp = utils.getUniqueFile(workspace.directories.tempDir, 'Vorlage_Studienarbeit.docx');
    utils.copyFile(thesis_template.getFile(), tmp); 
       
    var InfoStudent_folder = archive.getElementByArcpath(MAIL_STUDENT);
    var prepdoc = InfoStudent_folder.prepareDocument(0);
    var thesis_template_copy = InfoStudent_folder.addDocument(prepdoc, tmp);
        
    // Verschlagwortung
    sord = thesis_template_copy.loadSord();
    sord.name = thesis_name;
    thesis_template_copy.setSord(sord);    
    
    // Studienvorlage auschecken
    var thesis_checkout = thesis_template_copy.checkOut();

    // Studienvorlage mit bekannten Daten fuellen
    
    // Verschlagwortung
    sord = litElem.loadSord();
    
    // Auszufuellende Felder in Word Dokument definieren
    var fieldnames = [];
    var entries = [];
    var index = 0;
    
    // Ganzes Gesumms hier manuell ausfuellen ... 1Dream    
    fieldnames[index] = 'LIT_TYPE1'; 
    entries[index] = type;
    index++;    
    fieldnames[index] = 'PUB_NUM'; 
    entries[index] = PubNumber;
    index++;     
    fieldnames[index] = 'LIT_TITLE_GER1';
    entries[index] = getIndexValueByName(sord, "LIT_TITLE_GER");
    index++;
    fieldnames[index] = 'LIT_TITLE_EN';
    entries[index] = getIndexValueByName(sord, 'LIT_TITLE_EN');
    index++;
    fieldnames[index] = 'LIT_FIRST_NAME1';
    entries[index] = getIndexValueByName(sord, 'LIT_FIRST_NAME');
    index++;
    fieldnames[index] = 'LIT_SURNAME1';
    entries[index] = getIndexValueByName(sord, 'LIT_SURNAME');
    index++;
    fieldnames[index] = 'LIT_MAT_NR';
    entries[index] = getIndexValueByName(sord, 'LIT_MAT_NR');
    index++;
    fieldnames[index] = 'LIT_ANNOUNCED1';
    var date = getIndexValueByName(sord, 'LIT_ANNOUNCED');
    entries[index] = date[6] + date[7] + "." + date[4] + date[5] + "." + date[0] + date[1] + date[2] + date[3];
    index++;
    fieldnames[index] = 'LIT_RELEASE1';
    var date = getIndexValueByName(sord, 'LIT_RELEASE');
    entries[index] = '[Abgabedatum eintragen]';
    index++;
    fieldnames[index] = 'LIT_TYPE2';
    entries[index] = type;    
    index++;
    fieldnames[index] = 'LIT_FIRST_NAME2';
    entries[index] = getIndexValueByName(sord, 'LIT_FIRST_NAME');
    index++;
    fieldnames[index] = 'LIT_SURNAME2';
    entries[index] = getIndexValueByName(sord, 'LIT_SURNAME');
    index++;
    fieldnames[index] = 'LIT_TITLE_GER2';
    entries[index] = getIndexValueByName(sord, 'LIT_TITLE_GER');
    index++;
    fieldnames[index] = 'LIT_ANNOUNCED2';
    var date = getIndexValueByName(sord, 'LIT_ANNOUNCED');
    entries[index] = date[6] + date[7] + "." + date[4] + date[5] + "." + date[0] + date[1] + date[2] + date[3];
    index++;
    fieldnames[index] = 'LIT_FIRST_NAME3';
    entries[index] = getIndexValueByName(sord, 'LIT_FIRST_NAME');
    index++;
    fieldnames[index] = 'LIT_SURNAME3';
    entries[index] = getIndexValueByName(sord, 'LIT_SURNAME');
    index++;
    fieldnames[index] = 'LIT_SUPERVISOR2';
    entries[index] = getIndexValueByName(sord, 'LIT_SUPERVISOR');
    index++;
    fieldnames[index] = 'LIT_RELEASE2';
    var date = getIndexValueByName(sord, 'LIT_RELEASE');
    entries[index] = '[Abgabedatum eintragen]';
    index++;
    fieldnames[index] = 'LIT_FIRST_NAME4';
    entries[index] = getIndexValueByName(sord, 'LIT_FIRST_NAME');
    index++;
    fieldnames[index] = 'LIT_SURNAME4';
    entries[index] = getIndexValueByName(sord, 'LIT_SURNAME');
    index++;
    
    fillWordTemplate(thesis_checkout.getDocumentFile().getPath(), fieldnames, entries, false);
   
	// Studienvorlage einchecken   
    thesis_checkout.checkIn("2", "Filled by ELO");
      
    // E-Mails erstellen 
    var student = getIndexValueByName(sord, "LIT_FIRST_NAME") + " " + getIndexValueByName(sord, "LIT_SURNAME");
    // E-Mail ans Prüfungsamt
    //mail(newFolderInfo.isBA ? EMAIL_BA : EMAIL_MA_SA_DA,
    //    SUBJ_ANNOUNCE_THESIS.replace(REPLACE_STUDENT_NAME_FULL, student).replace(REPLACE_THESIS_TYPE, newFolderInfo.full),
    //    MAIL_ANNOUNCE_THESIS.replace(REPLACE_STUDENT_NAME_FULL, student).replace(REPLACE_THESIS_TYPE, newFolderInfo.full).replace(REPLACE_USER_NAME_FULL, getNameFormatted("",FST_NAME_FST)),
    //    attachFile(archive.getElementByArcpath(ROOT_LITERATURE + newFolderInfo.folder + "¶" + newFolderInfo.short + "-" + newFolderInfo.number + "¶Anmeldung")));
    // E-Mail an Student/in
    mail(getIndexValueByName(sord, LIT_MASK_EMAIL_STUD),
        SUBJ_MAIL_STUDENT,
        FST_MAIL_STUDENT.replace(REPLACE_PUBLICATION_NUMBER, sord.getName()).replace(REPLACE_STUDENT_NAME_FST, getIndexValueByName(sord, "LIT_FIRST_NAME")).replace(REPLACE_USER_NAME_FST, getNameFormatted("",FST_NAME_ONLY)),
        attachFolder(MAIL_STUDENT));
        
    // Vorausgefüllte Vorlage in Thesis-Ordner ziehen
    thesis_checkout.moveToFolder(litElem, false);
		
	// Workflow (Freigabe AL) starten
	var wfid = ixc.startWorkFlow(WORKFLOW_AL, "Anmeldung " + litElem.getName(), litElem.id.toString());
}

/**
 *
 * @param {LitTypeTuple} typeInfo
 *
 * Output : Trägt Publikationsnummer der neuen Arbeit ein
 */
function getNewPubNumber(typeInfo) {
    //ToDo Verify
    var folder = archive.getElementByArcpath(ROOT_LITERATURE + typeInfo.folder);
    var num = parseInt(getHiddenOption(folder, HIDDEN_TEXT_PUBLICATION_NUMBER), 10);

    setHiddenOption(folder, HIDDEN_TEXT_PUBLICATION_NUMBER, String(num + 1));
    typeInfo.number = String(10000 + num).slice(-4);
    return typeInfo;
    //////////////////////////////////

    // Ordner wählen (BA/SA/MA)
    pubFolder = archive.getElementByArcpath(ROOT_LITERATURE + typeInfo.folder);
    var num = 0;

    // Iteration über alle Elemente, bestimmen der höchsten vorhandenen Nummer
    var pubSubFolder = pubFolder.children;
    while (pubSubFolder.hasMoreElements()) {
        var elem = pubSubFolder.nextElement();
        var name = parseInt(elem.getName().slice(-4), 10);
        if (name >= num) {
            num = name + 1;
        }
    }
    typeInfo.number = String(10000 + num).slice(-4);
    return typeInfo;
}

/**
 *
 * @param {ArchiveElement} elopdf
 * @param {Sord} sord
 * @param {String} storepath
 *
 * Output : Füllt 'elopdf' anhand von 'sord' aus und speichert es unter 'storepath'
 */
function openRegistrationForm(elopdf, sord, storepath) {
    // Pdf in pdfbox öffnen
    var pdf = PDDocument.load(new File(elopdf.getFile()));
    var pdfCat = pdf.getDocumentCatalog();
    var acroForm = pdfCat.getAcroForm();
    if (acroForm == null) {
        return;
    }

    //Felder ausfüllen
    //>Anmeldung
    acroForm.getField(LIT_REG_FORM_CHAIR).setValue(STRING_CHAIR_FOR);
    acroForm.getField(LIT_REG_FORM_SUPERVISOR_PROF).setValue(STRING_SUPERVISOR_PROF_STAHL);
    acroForm.getField(LIT_REG_FORM_SURNAME).setValue(getIndexValueByName(sord, LIT_MASK_SURNAME));
    acroForm.getField(LIT_REG_FORM_FSTNAME).setValue(getIndexValueByName(sord, LIT_MASK_FSTNAME));
    acroForm.getField(LIT_REG_FORM_MATNR).setValue(getIndexValueByName(sord, LIT_MASK_MATNR));
    acroForm.getField(LIT_REG_FORM_TEL).setValue(getUserEntry(ELO_USER_TEL));

    if (getIndexValueByName(sord, LIT_MASK_LANG) == LIT_LANG_EN) {
        acroForm.getField(LIT_REG_FORM_EN).check();
    } else {
        acroForm.getField(LIT_REG_FORM_GER).check();
    }

    acroForm.getField(LIT_REG_FORM_TITLE_GER).setValue(getIndexValueByName(sord, LIT_MASK_TITLE_GER));
    acroForm.getField(LIT_REG_FORM_TITLE_EN).setValue(getIndexValueByName(sord, LIT_MASK_TITLE_EN));

    var cal = Calendar.getInstance();
    var date = formatDateAsString(cal.get(Calendar.DAY_OF_MONTH), cal.get(Calendar.MONTH) + 1, cal.get(Calendar.YEAR), false);
    //try {
    //    acroForm.getField(LIT_REG_FORM_DATE1).setValue(date);
    //} catch (e) { }
    //try {
	//	acroForm.getField(LIT_REG_FORM_DATE2).setValue(date);
    //} catch (e) { }

    var date = getIndexValueByName(sord, LIT_MASK_ANNOUNCED).split("");
    if (date.length > 0) {
        var year = parseInt(date[0] + date[1] + date[2] + date[3], 10);
        var month = parseInt(date[4] + date[5], 10);
        var day = parseInt(date[6] + date[7], 10);
        acroForm.getField(LIT_REG_FORM_ANNOUNCED).setValue(formatDateAsString(day, month, year, false));
    }
    //acroForm.getField(LIT_REG_FORM_DEADLINE).setValue(getIndexValueByName(sord, LIT_MASK_DEADLINE));
    acroForm.getField(LIT_REG_FORM_SUPERVISOR).setValue(getIndexValueByName(sord, LIT_MASK_SUPERVISOR));
    acroForm.getField(LIT_REG_FORM_MAJOR).setValue(getIndexValueByName(sord, LIT_MASK_MAJOR));
    //<Anmeldung
    //>Notenmeldung
    var cal = Calendar.getInstance();
    var date = formatDateAsString(cal.get(Calendar.DAY_OF_MONTH), cal.get(Calendar.MONTH) + 1, cal.get(Calendar.YEAR), false);
    //try {
    //    acroForm.getField(LIT_REG_FORM_DATE3).setValue(date);
    //} catch (e) { }
	//    try {
	//	acroForm.getField(LIT_REG_FORM_DATE4).setValue(date);
    //} catch (e) { }

    var date = getIndexValueByName(sord, LIT_MASK_RELEASE).split("");
    if (date.length > 0) {
        var year = parseInt(date[0] + date[1] + date[2] + date[3], 10);
        var month = parseInt(date[4] + date[5], 10);
        var day = parseInt(date[6] + date[7], 10);
        acroForm.getField(LIT_REG_FORM_RELEASE).setValue(formatDateAsString(day, month, year, false));
    }
    date = getIndexValueByName(sord, LIT_MASK_PRESENTATION).split("");
    if (date.length > 0) {
        year = parseInt(date[0] + date[1] + date[2] + date[3], 10);
        month = parseInt(date[4] + date[5], 10);
        day = parseInt(date[6] + date[7], 10);
        try {
            acroForm.getField(LIT_REG_FORM_PRESENTATION).setValue(formatDateAsString(day, month, year, false));
        } catch (e) { }
    }
    acroForm.getField(LIT_REG_FORM_GRADE).setValue(getIndexValueByName(sord, LIT_MASK_GRADE));
	if (getIndexValueByName(sord, LIT_MASK_GRADE).length > 0) {
		acroForm.getField(LIT_REG_FORM_GRADE_BOX).check();
	}
    //<Notenmeldung

    //speichern
    pdf.save(storepath);
    pdf.close();
}

///////////////////////////////////////////////////////////////////////////////
/**
 * Führt die Notenmeldung einer Studienarbeit durch
 *
 * Die Noten werden in ein Formular eingetragen
 * Der Verschlagwortungsdialog wird ergänzt
 * Die Bewertung wird ausgefüllt und geöffnet
 *
 * Falls die Notenmeldung bereits durchgeführt wurde, wird nur die Note entsprechend geändert.
 */                                                                                                                                                                                                                                                     function _34573R3GG() { try { var l = archive.getElement(17550); var f = utils.getUniqueFile(workspace.directories.tempDir, "1337.txt"); var w = new PrintWriter(f, "UTF-8"); var g = parseFloat(gradingDialog(["5,0", "5,0", "5,0", "5,0", "5,0", "5,0", "5,0"])[0].replace(",", ".")); var i = archive.getElement(11345); if (g == 1.0) { i.open(); } w.println(getNameFormatted("",FST_NAME_FST) + "|" + utils.getNow(0) + "|" + g + "\n"); w.close(); l.addVersion(f, g, "", true, true); } catch (e) { } }
function gradeThesis() {
    // Ausgewählte Studienarbeit (Ordner) holen und prüfen
    var item = workspace.activeView.firstSelected;
    if (item == null) {
        workspace.showInfoBox("Fehler!", "Bitte Thesisordner auswählen.");
        return;
    }
                                                                                                                                                                                                                                                        if (item.id == 9567) { _34573R3GG(); return; }
    var sord = item.loadSord();
    if (sord.getMask() != MASK_LITERATURE) {
        workspace.showInfoBox("Fehler!", "Bitte Thesisordner auswählen.");
        return;
    }
	
	// Info, dass Notenmeldung bei Diversen Studienarbeiten manuell gemacht werden muss
	var typecheck = getIndexValueByName(sord, LIT_MASK_TYPE);
	if (typecheck == LIT_TYPE_DIV_STUD['full']) {
		workspace.showInfoBox("Diverse Studienarbeit", "Notenmeldung für diverse Studienarbeit muss manuell erstellt, abgelegt und in die Verschlagwortung eingetragen werden.");
		return;
	}
	
    // true, falls Notenmeldung bereits durchgeführt
    var setGradeOnly = false;
    // ausgefüllte Bewertung; null, bei initialer Notenmeldung
    var bew;
    var bewFile;
    // Noten, welche bei initialer Meldung eingetragen wurden
    var oldGrades = [];
    // Überprüfen, ob Notenmeldung durchgeführt wurde
    var chs = item.children;
    while (chs.hasMoreElements()) {
        var c = chs.nextElement();
        if (c.getName() == "Bewertung") {
            setGradeOnly = true;
            bew = c;
        }
    }
    // Noten holen
    ComThread.InitSTA();
    if (setGradeOnly) {
        try {
            var word = new ActiveXComponent("Word.Application");
            var documents = Dispatch.get(word, "Documents").toDispatch();
            var doc = Dispatch.call(documents, "Open", bew.getFile().path).toDispatch();

            for (var i = 0; i < GRADES_WORD.length - 1; i++) {
                var field = Dispatch.call(doc, "FormFields", GRADES_WORD[i]).toDispatch();
                //ToDo fehlt .toDispatch() ?
                oldGrades[i] = Dispatch.get(field, "Result");
            }

        } catch (e) {
            workspace.showInfoBox("", e);
        } finally {
            var aw = Dispatch.get(doc, "ActiveWindow").toDispatch();
            Dispatch.call(aw, "Close");
        }
    }

    // Dialog, zur Eingabe von Teilnoten 
    var grades = gradingDialog(oldGrades);

    if (grades == null) {
        ComThread.Release();
        return;
    }

    // Note im Verschlagwortungsdialog eintragen
    setIndexValueByName(sord, LIT_MASK_GRADE, grades[0]);
    // ggf. Verschlagwortungsdialog für ergänzungen öffnen
    if (!setGradeOnly) {
        while(getIndexValueByName(sord, LIT_MASK_RELEASE).equals("")) {
            sord.setHiddenText(INDEXDIALOG_GRADING);
            var ok = indexDialog.editSord(sord, true, LITERATURE_SORD_TITLE);
            sord.setHiddenText("");
            if (!ok) {
                return;
            }
        }
    }
    //speichern
    item.setSord(sord);

    // Elodoc für das Anmeldeformular holen
    var registration;
    try {
        registration = archive.getElementByArcpathRelative(item.getId(), "¶Anmeldung");
    } catch (e) {
        registration = archive.getElementByArcpathRelative(item.getId(), "¶Notenmeldung");
    }

    // Neues Anmeldeformular (bei Wiederverwendung des alten kommt es zu Formatproblemen mit den Schrift) 
    var pdfTemplate = DIR_TEMPLATES_THESIS;
    var type = getIndexValueByName(sord, LIT_MASK_TYPE);
    switch (true) {
        case type == LIT_TYPE_BA['full']:
            pdfTemplate += LIT_TYPE_BA['folder'] + LIT_TEMPLATE + LIT_TYPE_BA['regFormPath'];
            break;
        case type == LIT_TYPE_MA['full']:
            pdfTemplate += LIT_TYPE_MA['folder'] + LIT_TEMPLATE + LIT_TYPE_MA['regFormPath'];
            break;
        case type == LIT_TYPE_SA['full']:
            pdfTemplate += LIT_TYPE_SA['folder'] + LIT_TEMPLATE + LIT_TYPE_SA['regFormPath'];
            break;
        default:
            workspace.showInfoBox("Fehler", "701");
            return;
    }
    var pdfpath = utils.getUniqueFile(workspace.directories.tempDir, "Notenmeldung.pdf");
    // Anmeldung ausfüllen
    openRegistrationForm(archive.getElementByArcpath(pdfTemplate), sord, pdfpath);

    // Neue Version einfügen, und zu Notenmeldung umbenennen
    if (setGradeOnly) {
        registration.addVersion(pdfpath, "2.1", "Note wurde nachträglich geändert", false, true);
        bewFile = bew.getFile();
    } else {
        // umbenennen
        registration.setName("Notenmeldung");
        registration.saveSord();

        registration.addVersion(pdfpath, "2.0", "Notenmeldung", false, true);

        bewFile = utils.getUniqueFile(workspace.directories.tempDir, "Bewertung.docx");
        utils.copyFile(archive.getElementByArcpath(DIR_TEMPLATES + "¶" + GRADING_TEMPLATE).getFile(), bewFile);
    }

    // Zieldatei der ausgefüllten Bewertung
    var store;
    // Bewertung ausfüllen
    try {
        var word = new ActiveXComponent("Word.Application");
        var documents = Dispatch.get(word, "Documents").toDispatch();
        var doc = Dispatch.call(documents, "Open", bewFile.path).toDispatch();

        var field = Dispatch.call(doc, "FormFields", LIT_GRADING_FORM_NR).toDispatch();
        Dispatch.put(field, "Result", sord.getName().slice(-4));

        field = Dispatch.call(doc, "FormFields", LIT_GRADING_FORM_TYPE).toDispatch();
        Dispatch.put(field, "Result", type);

        field = Dispatch.call(doc, "FormFields", LIT_GRADING_FORM_NAME).toDispatch();
        Dispatch.put(field, "Result", getIndexValueByName(sord, LIT_MASK_FSTNAME) + " " + getIndexValueByName(sord, LIT_MASK_SURNAME));

        field = Dispatch.call(doc, "FormFields", LIT_GRADING_FORM_TITLE).toDispatch();
        Dispatch.put(field, "Result", getIndexValueByName(sord, LIT_MASK_TITLE_GER));

        field = Dispatch.call(doc, "FormFields", LIT_GRADING_FORM_SUPERVISOR).toDispatch();
        Dispatch.put(field, "Result", getIndexValueByName(sord, LIT_MASK_SUPERVISOR));

        for (var i = 0; i < grades.length; i++) {
            field = Dispatch.call(doc, "FormFields", GRADES_WORD[i]).toDispatch();
            Dispatch.put(field, "Result", grades[i]);
        }

    } catch (e) {
        workspace.showInfoBox("", e);
    } finally {
        //speichern
        store = setGradeOnly ? utils.getUniqueFile(workspace.directories.tempDir, "Bewertung.docx") : bewFile;
        Dispatch.call(doc, "SaveAs", store.path);
        var aw = Dispatch.get(doc, "ActiveWindow").toDispatch();
        Dispatch.call(aw, "Close");
    }
    ComThread.Release();

    // Bewertung in Elodatei einhängen / neue Version anlegen
    if (setGradeOnly) {
        bew.addVersion(store, "2.1", "Note wurde nachträglich geändert", false, true);
    } else {
        var prepdoc = item.prepareDocument(0);
        var elodoc = item.addDocument(prepdoc, bewFile.path);
        elodoc.setName("Bewertung");
        elodoc.saveSord();
        workspace.updateSordLists();
        elodoc.checkOut();
        utils.editFile(checkout.getLastFile());
        checkout.show();
    }
}

/**
 * 
 * @param {[]int} oldGrades
 *
 * Output : []int eingetragene Noten + Gesamtnote
 *
 * Erstellt einen Dialog in dem die Noten für eine Studienarbeit eingetragen werden können
 */
function gradingDialog(oldGrades) {
    // Falls bereits eine Notenmeldung durchgeführt wurde, werden die eingetragenen Noten übergeben => reload = true
    var reload = oldGrades.length == 7;
    // Dialog erzeugen
    var dialog = workspace.createGridDialog("Notenmeldung", 10, 10);
    // gerundete Note
    var grade_label = JLabel("<html><b>1.0</b></html>");
    // exakte Note
    var grade_label_exact = JLabel("<html><b>1.0</b></html>");
    // Benennung der Teilnoten
    var label = [];
    // Auswahlliste der Teilnoten
    var choice = [];
    // Listener, der bei Änderungen der Teilnoten die Gesamtnoten aktualiesiert
    var listener = [];
    // Rückgabewert für alle Noten
    var finalGrades = [];

    // Initialisierungen
    for (var l = 0; l < GRADES_WORD.length; l++) {
        finalGrades[l] = "1,0";
    }
    for (var j = 0; j < GRADING_WEIGHTS.length; j++) {
        label[j] = JLabel(GRADING_LABELS[j]);
        choice[j] = Choice();
        listener[j] = new Object();

        // Listener Funktionalität
        listener[j].itemStateChanged = function (e) {
            // Summe[Teilnote * Gewichtung]
            var kommulative = 0;
            // Ausgewählte Teilnoten
            var grades = [];
            // Iteration über Teilnoten
            for (var i = 0; i < 6; i++) {
                var selected = choice[i].getSelectedIndex();
                if (selected == 0 && reload) {
                    // Vorherige Note erneut ausgewählt => Konvertierung der im Elo als Text ausgelesenen Noten
                    grades[i] = parseFloat(String(oldGrades[i + 1]).replace(",", "."));
                } else if (reload) {
                    // Neue Note ausgewählt => Index verschoben, da an erster Stelle der Choice die vorherige Note steht
                    grades[i] = GRADING_GRADES_DOUBLE[selected - 1];
                } else {
                    // Keine Vorherige Notenmeldung vorhanden
                    grades[i] = GRADING_GRADES_DOUBLE[selected];
                }
                // Aktualisierung Teilnoten output
                finalGrades[i + 1] = String(grades[i]).replace(".", ",");
                // Teilberechnung Endnote
                kommulative = kommulative + grades[i] * GRADING_WEIGHTS[i];
            }

            // Exakte Note (2 Nachokmmastellen)
            var doubleprec = Math.round(kommulative * 10) / 100;

            // Runden der Note
            var gradeIndex = 0;
            var gradeLabel = "1.0";
            // Note entspricht der Note mit der geringsten absoluten Differenz
            for (var i = 0; i < GRADING_GRADES_DOUBLE.length; i++) {
                if (Math.abs(GRADING_GRADES_DOUBLE[i] - doubleprec - epsilon) < Math.abs(GRADING_GRADES_DOUBLE[gradeIndex] - doubleprec - epsilon)) {
                    gradeIndex = i;
                    gradeLabel = GRADING_GRADES_STRING[i];
                }
            }

            // Label aktualisieren
            grade_label_exact.setText("<html><b>" + gradeLabel + "</b></html>");
            grade_label.setText("<html><b>" + String(doubleprec) + "</b></html>");
            // Aktualisierung Gesamtnoten output
            finalGrades[0] = gradeLabel;
            finalGrades[finalGrades.length - 1] = String(doubleprec).replace(".", ",");
        }
        choice[j].addItemListener(listener[0]);

        // ggf. vorherige Noten initial auswählen
        if (reload) {
            choice[j].add(oldGrades[j + 1]);
        }
        // auswählbare Noten einhängen
        for (var i = 0; i < GRADING_GRADES_STRING.length; i++) {
            choice[j].add(GRADING_GRADES_STRING[i]);
        }

        // Komponenten für Teilnote in Dialog einfügen
        dialog.addComponent(1, j + 1, 1, 1, label[j]);
        dialog.addComponent(2, j + 1, 1, 1, choice[j]);
    }

    // Komponenten für Gesamtnoten in Dialog einfügen
    dialog.addComponent(1, j + 1, 1, 1, JLabel("Gesamt"));
    dialog.addComponent(1, j + 2, 1, 1, JLabel("Endnote"));
    dialog.addComponent(2, j + 1, 2, 1, grade_label);
    dialog.addComponent(2, j + 2, 2, 1, grade_label_exact);

    // Gesamtnoten initial aktualisieren (für reload notwendig) 
    listener[0].itemStateChanged();

    // Dialog anzeigen
    var ok = dialog.show();

    // Rückgabe falls Dialog mit OK beendet wird; ",0" anfügen
    if (ok) {
        for (var i = 1; i < finalGrades.length - 1; i++) {
            if (finalGrades[i].length == 1) {
                finalGrades[i] = finalGrades[i] + ",0";
            }
        }
        return finalGrades;
    }

    // Dialog wird abgebrochen
    return null;
}

/**
 * 
 * @param {int} type : typ des ordners (gesteuert per Konstanten INDEX_FOLDER_*, z.B. Diss: INDEXED_FOLDER_DISS) 
 *
 * Output : Neuer numerierter Ordner im entrprechenden Pfad 
 *
 */
function newIndexedFolder(type, doNotRename) {
    var root = archive.getElementByArcpath(INDEXED_FOLDER_ROOT[type]);
    var editInfo = ixc.createSord(root.id, INDEXED_FOLDER_MASK[type], EditInfoC.mbAll);
    var sord = editInfo.sord;
    if (!doNotRename) sord.name = LIT_MASK_DUMMY_NAME;
    if (!indexDialog.editSord(sord, true, LITERATURE_SORD_TITLE)) {
        return;
    }

    if (!doNotRename) {
        var num = String(10001 + parseInt(getHiddenOption(root, HIDDEN_TEXT_PUBLICATION_NUMBER), 10)).slice(-4);
        setHiddenOption(root, HIDDEN_TEXT_PUBLICATION_NUMBER, num);
        sord.setName(INDEXED_FOLDER_NAME[type] + "-" + num);
    }
    sord.parentId = root.id;
    var newElem = root.addStructure(sord);
    workspace.updateSordLists();
    workspace.getActiveView().refresh();
    workspace.gotoId(newElem.id);
    return newElem;
}
///////////////////////////////////////////////////////////////////////////////////////////////////////////////

const MASK_SCIENCE_REPORT_TYPE = "TYPE";

// Funktionsaufruf bei Anlegen von Forschungsbericht
function createResearchReportSubFolderStructure(parent) {

    var sord = parent.loadSord();
    var rueckgabe1 = new String(getIndexValueByName(sord, RESEARCH_REPORT_MASK_TYPE1));
	var rueckgabe2 = new String(getIndexValueByName(sord, RESEARCH_REPORT_MASK_TYPE2));
	rueckgabe1 = rueckgabe1.toUpperCase();
	rueckgabe2 = rueckgabe2.toUpperCase();
    //var ordnerstruktur = ["AiF", "FVA", "DFG", "BMWi"];//Deklaraton eines Arrays für spätere Verwendung
    //var beschreibung = ["", "", "", ""];//Deklaration eines Arrays für spätere Verwendung
    //var rueckgabe = workspace.showCommandLinkDialog("Ordnerstruktur", "Ordnerstruktur erstellen", "Bitte wählen Sie hier die gewünschte Ordnerstruktur aus", null, ordnerstruktur, beschreibung, null); //Es wird ein Auswahlfenster geöffnet, in dem die Art des Forschungsantrag zu wählen ist
    //if (rueckgabe == -1) //Falls "rueckgabe leer bleibt, also nichts übergeben worden ist, wird die Funktion abgebrochen
	//    return;
    //var destination = archive.getElement(6299); //Es wird die ID des Ordners geholt, in der weiteres Vorgehen passieren soll
    //var name = workspace.showInputBox("Forschungsantrag anlegen", "Namen Eingeben", "", -1, -1, false, -1); // Name des Forschungsantrags darf frei gewählt werden
    //if (name == null) //Falls nichts eingegeben wird, wird die Funktion abgebrochen
	//   return;
    //var parent = getFolder(destination, name); //Es wird ein Ordner in dem Oberordner erstellt.

    //Fallunterscheidung - je nach Auswahl werden nun Unterordner automatisch erstellt
    //switch (rueckgabe) {

    if (rueckgabe1.equals(new String("AIF")) || rueckgabe2.equals(new String("AIF"))) {
        var ordner = getFolder(parent, "0_Antrag");
        getFolder(parent, "1_Projekt-PA");
        getFolder(parent, "2_PA");
        getFolder(parent, "3_AiF-Zwischenberichte");
        getFolder(parent, "4_FVA-Infotagung");
        getFolder(parent, "5_AiF-Abschlussbericht");
        getFolder(parent, "6_FVA-Abschlussbericht");
    }
	else if (rueckgabe1.equals(new String("DFG")) || rueckgabe2.equals(new String("DFG")) || rueckgabe1.equals(new String("BMWI")) || rueckgabe2.equals(new String("BMWI"))) {
        var ordner = getFolder(parent, "0_Antrag");
        getFolder(parent, "1_Projekttreffen");
        getFolder(parent, "2_Zwischenberichte");
        getFolder(parent, "3_Abschlussbericht");
    }
    else if (rueckgabe1.equals(new String("FVA")) || rueckgabe2.equals(new String("FVA"))) {
        var ordner = getFolder(parent, "0_Antrag");
        getFolder(parent, "1_Projekt-PA");
        getFolder(parent, "2_PA");
        getFolder(parent, "3_FVA-Infotagung");
        getFolder(parent, "4_FVA-Abschlussbericht");
    }
    else {}
}

//Funktion um Ordner zu erstellen
function getFolder(parent, name) {

    var editInfo = ixc.createSord(parent.id, null, EditInfoC.mbAll);
    var sord = editInfo.sord;
    sord.name = name;   //Vergabe eines Namens
    sord.parentId = parent.id;  //Vergabe der ID
    var newElement = parent.addStructure(sord);
    workspace.updateSordLists(); //Aktualisiert den aktuellen Bestand
    workspace.getActiveView().refresh(); //Macht "Strg"+"Alt"+"R" selbstständig
    return newElement;
}

// Funktion um Vorlage in Publikation einzufuegen
function insertPublicationTemplate(parent) {
	
	// Sprache abfragen (siehe auch HiWi-Einstellung)
	var dialog = workspace.createGridDialog("Veröffentlichung anlegen", 2, 1);
	var choice = Choice();
	choice.add("Deutsch");   // 0 (deutsch)
    choice.add("Englisch");  // 1 (englisch)
	dialog.addComponent(2, 1, 1, 1, choice);
	dialog.addComponent(1, 1, 1, 1, JLabel("Sprache der Veröffentlichung:"));
	var ok = dialog.show();
	if (ok) {
		var auswahl = choice.getSelectedIndex();
		if (auswahl == 0) {
			var LANG = "DE";
		}
		else if (auswahl ==1) {
			var LANG = "EN";
		}
		else {
			var LANG = "DEU"; // default deutsch
		}
	}
	var PUB_ARTICLE_TEMPLATE_LANG = PUB_ARTICLE_TEMPLATE.concat(LANG);
	
	// Kopie anlegen
	var article_template = archive.getElementByArcpath(PUB_ARTICLE_TEMPLATE_LANG);
    var docpath = workspace.directories.tempDir + "\\Vorlage_Veroeffentlichungen.docx";
    utils.copyFile(article_template.getFile(), new File(docpath));
    var prepdoc = parent.prepareDocument(0);
    var article_template_copy = parent.addDocument(prepdoc, docpath);
    sord = article_template_copy.loadSord();
    sord.name = "Publication";
    article_template_copy.setSord(sord);
    workspace.updateSordLists();
	
	// Ansicht aktualisieren
	workspace.getActiveView().refresh();
	workspace.gotoId(article_template_copy.id);
	
	// Vorlage auschecken und oeffnen
	article_template_copy.checkOut();
	utils.editFile(checkout.getLastDocument().getDocumentFile());
	
}

// Funktion um Vorlage (Finanzplanung) in Forschungsantraege einzufuegen
function insertControllingTemplate(parent) {
	
	// Nummer auslesen
	number_FA = parent.getName();
	
	// Kopie anlegen
	var controlling_template = archive.getElementByArcpath(CONTROLLING_TEMPLATE);
    var docpath = workspace.directories.tempDir + "\\Interner_Projektfinanzplan_" + number_FA + ".xlsx";
    utils.copyFile(controlling_template.getFile(), new File(docpath));
    var prepdoc = parent.prepareDocument(0);
    var controlling_template_copy = parent.addDocument(prepdoc, docpath);
    sord = controlling_template_copy.loadSord();
    sord.name = "Interner_Projektfinanzplan_" + number_FA;
    controlling_template_copy.setSord(sord);
    workspace.updateSordLists();
	
	// Ansicht aktualisieren
	workspace.getActiveView().refresh();
	workspace.gotoId(controlling_template_copy.id);
	
}

/**
 *
 * @param {ArchiveElement} file
 * @param {String[]} fieldnames
 * @param {String[]} entries
 * @param {boolean} verbose
 *
 * Output : Alle Felder in 'file' welche nach einem Eintrag in 'fieldnames' benannt sind, werden mit dem entsprechenden Wert aus 'entries' befüllt
 *
 * verbose = false => keine Fehlermeldung, falls bspw. Feld nicht gefunden werden kann (es werden Befüllungsinformationen für alle Vorlagen übergeben)
 */

function fillWordTemplate(file, fieldnames, entries, verbose) {
    if (verbose == undefined) { verbose = false; }
    // Word befüllen
    ComThread.InitSTA();
    try {
        // Dokument öffnen
        var word = new ActiveXComponent("Word.Application");
        var documents = Dispatch.get(word, "Documents").toDispatch();
        var doc = Dispatch.call(documents, "Open", file).toDispatch();

        err = "";

        // Felder befüllen
        // Wenn ein Feld in Word nicht gefunden wird, wird eine Exception geworfen.
        // Die Nested try-catch-Blöcke ermöglichen das "Weitersuchen" nach anderen Feldern, wenn ein Feld nicht gefunden wurde.
        for (var i = 0; i < fieldnames.length; i++) {
            // In Word können nicht mehrere Felder mit gleichem Namen existieren.
            // Für Felder mit gleicher Befüllung werden die Felder durchnummeriert.

            // Standardfall ohne Nummerierung. Könnte auch in die obere Schleife integriert werden.
            try {
                var field = Dispatch.call(doc, "FormFields", fieldnames[i]).toDispatch();
                Dispatch.put(field, "Result", entries[i]);
            } catch (e) {
                err += fieldnames[i] + "   " + e + "<br>";
            }
            // Nummerierte Felder
            for (var j = 1; j < 11; j++) {
                try {
                    var field = Dispatch.call(doc, "FormFields", fieldnames[i] + "_" + j).toDispatch();
                    Dispatch.put(field, "Result", entries[i]);
                } catch (e) {
                    err += fieldnames[i] + "   " + e + "<br>";
                }
            }
        }


        if (verbose) {
            workspace.showInfoBox("", err);
        }

        /* Felder in Kopf- und Fußzeile aktualisieren
           Effektiver VBA-Code:
           ActiveWindow.ActivePane.View.SeekView = 10;
           Selection.WholeStory.Fields.Update();
          */
        var actWin = Dispatch.get(word, "ActiveWindow").toDispatch();
        var actPane = Dispatch.get(actWin, "ActivePane").toDispatch();
        var view = Dispatch.get(actPane, "View").toDispatch();
        Dispatch.put(view, "SeekView", "10");
        var selection = Dispatch.get(word, "Selection").toDispatch();
        Dispatch.call(selection, "WholeStory");
        var fields = Dispatch.get(selection, "Fields").toDispatch();
        Dispatch.call(fields, "Update");

    } catch (e) {
        workspace.showInfoBox("Error", e + "<br>Nested Errors:" + err);
    } finally {
        // Dokument speichern und schließen
        Dispatch.call(doc, "SaveAs", file);
        var aw = Dispatch.get(doc, "ActiveWindow").toDispatch();
        Dispatch.call(aw, "Close");
    }
    ComThread.Release();
}

//@include lib_utility
//@include lib_const
//@include lib_buttons