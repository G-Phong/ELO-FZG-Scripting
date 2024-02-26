// Zum Debuggen Funktion direkt Aufrufen (z.B. Testbutton) dann im Debugger DBG_RIBBON = true setzen.
DBG_RIBBON = false;
function createRibbon() {
    if (ribbon.getTab(TAB_FZG) != null && !DBG_RIBBON)
        return;

    // Tabs
    var tabFzg = ribbon.addTab(71, null, TAB_FZG);
    tabFzg.setTitle("FZG");
    var tabAdmin = ribbon.addTab(81, null, TAB_ADMIN);
    tabAdmin.setTitle("Admin");
    //tabAdmin.setVisibleCallback(isAdmin, this);

    if (!isAdmin())
        tabAdmin.setVisibleCallback(function () { return false; }, this);
    var tabProvisional = ribbon.addTab(91, null, TAB_PROVISIONAL);
    tabProvisional.setTitle("Vorläufig");
    //tabProvisional.setVisibleCallback(testFunctions, this);

    if (!testFunctions())
        tabProvisional.setVisibleCallback(function () { return false; }, this);

    /////////////////////////////////////////////////////////////////////////////////
    // Bands
    // FZG Tab
    var bandStudentProject = ribbon.addBand(tabFzg.getId(), "10", BAND_STUDENT_PROJECT);
    bandStudentProject.setTitle("Studienarbeiten");
    var bandStudierende = ribbon.addBand(tabFzg.getId(), "15", BAND_STUDIERENDE);
    bandStudierende.setTitle("Studierende");
	var bandPublication = ribbon.addBand(tabFzg.getId(), "20", BAND_PUBLICATION);
    bandPublication.setTitle("Publikationen");
    var bandVerwaltung = ribbon.addBand(tabFzg.getId(), "21", BAND_VERWALTUNG);
    bandVerwaltung.setTitle("Verwaltung");
    var bandSketchDepot = ribbon.addBand(tabFzg.getId(), "25", BAND_SKETCH_DEPOT);
    bandSketchDepot.setTitle("Zeichnungsarchiv");
	var bandAddress = ribbon.addBand(tabFzg.getId(), "30", BAND_ADDRESS);
    bandAddress.setTitle("Adressen");
	var bandPurchase = ribbon.addBand(tabFzg.getId(), "35", "BAND_PURCHASE");
    bandPurchase.setTitle("Bestellungen");
    var bandClear = ribbon.addBand(tabFzg.getId(), "37", "BAND_CLEAR");
    bandClear.setTitle("Genehmigungen");
    var bandOther = ribbon.addBand(tabFzg.getId(), "40", BAND_OTHER);
    bandOther.setTitle("Sonstiges");
    var bandInfo = ribbon.addBand(tabFzg.getId(), "45", BAND_INFO);
    bandInfo.setTitle("Info");

    // Admin Tab
    var bandAdmin = ribbon.addBand(tabAdmin.getId(), "10", BAND_ADMIN);
    bandAdmin.setTitle("Administration");

    // Vorlaeufig Tab
    var bandProvisional = ribbon.addBand(tabProvisional.getId(), "10", BAND_PROVISIONAL);
    bandProvisional.setTitle("Voräufig");

    /////////////////////////////////////////////////////////////////////////////////
    // Buttons
    // Achtung: variable 'button' wird wiederverwendet!
    var button;

    // FZG Tab
    // Studienarbeiten Band
    // Neue Studienarbeit
    button = ribbon.addButton(tabFzg.getId(), bandStudentProject.getId(), BUTTON_NEW_STUDENT_WORK);
    button.setOrdinal(2);
    button.setTitle("Studienarbeit anmelden");
    button.setIconName("ScriptButton2");
    // Notenmeldung
    button = ribbon.addButton(tabFzg.getId(), bandStudentProject.getId(), BUTTON_GRADE_STUDENT_WORK);
    button.setOrdinal(3);
    button.setTitle("Notenmeldung");
    button.setIconName("ScriptButton3");
// Anmeldung BASAMA beim Prüfungsamt
    button = ribbon.addButton(tabFzg.getId(), bandStudentProject.getId(), BUTTON_BASAMA);
    button.setOrdinal(43);
    button.setTitle("Studienarbeit anmelden (ED-Portal)");
    button.setIconName("ScriptButton43");
// BUTTON_BASAMA_Benotung
    button = ribbon.addButton(tabFzg.getId(), bandStudentProject.getId(), BUTTON_BASAMA_Benotung);
    button.setOrdinal(44);
    button.setTitle("Notenmeldung (ED-Portal)");
    button.setIconName("ScriptButton44");


    // Studierende Band
    // HiWi-Einstellung
    button = ribbon.addButton(tabFzg.getId(), bandStudierende.getId(), BUTTON_NEW_HIWI);
    button.setOrdinal(9);
    button.setTitle("HiWi-Einstellung");
    button.setIconName("ScriptButton9");
    // SAM-Sicherheitsunterweisung
    button = ribbon.addButton(tabFzg.getId(), bandStudierende.getId(), BUTTON_SAM);
    button.setOrdinal(40);
    button.setTitle("SAM Sicherheitsunterweisung");
    button.setIconName("ScriptButton40");


    // Verwaltungen Band
    // Forschungsvorhaben
    button = ribbon.addButton(tabFzg.getId(), bandVerwaltung.getId(), BUTTON_FORSCHUNGSVORHABEN);
    button.setOrdinal(45); //TODO: Checken, ob der Wert hier passt
    button.setTitle("Forschungsvorhaben anlegen");
    button.setIconName("ScriptButton27");
    //Callback funktion für die Sichtbarkeit des Buttons
    //button.setEnabledCallback(function () {return isLeitungskreis() || isAdmin() || isVerwaltung();}, this); //Zugriff nur für Admins, Leitungskreis oder Verwaltung
    //button.setEnabledCallback(function () {return isAdmin();}, this); //Zugriff nur für Admins
    button.setEnabledCallback(function () {return false;}, this); //Zugriff gesperrt (für debugging zwecke)
    button.setTooltip("Nur für Leitungskreis, Verwaltung oder Admins möglich")


	// Genehmigungen Band
    // Urlaub beantragen
    button = ribbon.addButton(tabFzg.getId(), bandClear.getId(), BUTTON_HOLIDAY);
    button.setOrdinal(39);
    button.setTitle("Urlaub beantragen");
    button.setIconName("ScriptButton39");
	// Mobiles Arbeiten-Dokumentation freigeben
    button = ribbon.addButton(tabFzg.getId(), bandClear.getId(), BUTTON_CLEAR_HOMEOFFICE);
    button.setOrdinal(41);
    button.setTitle("Mobiles Arbeiten dokumentieren");
    button.setIconName("ScriptButton41");
	// Zeiterfassung korrigieren
    button = ribbon.addButton(tabFzg.getId(), bandClear.getId(), BUTTON_CLEAR_GLEITZEIT);
    button.setOrdinal(42);
    button.setTitle("Zeiterfassung korrigieren");
    button.setIconName("ScriptButton42");

    // Sonstiges Band
    // Vorlagen
    button = ribbon.addButton(tabFzg.getId(), bandOther.getId(), BUTTON_TEMPLATES);
    button.setOrdinal(7);
    button.setTitle("Dokument aus Vorlage");
    button.setIconName("ScriptButton7");
    // Ordner auschecken
    button = ribbon.addButton(tabFzg.getId(), bandOther.getId(), BUTTON_CHECKOUT_FOLDER);
    button.setOrdinal(11);
    button.setTitle("Ordner auschecken und bearbeiten");
    button.setIconName("ScriptButton11");
    // Ordner einchecken
    button = ribbon.addButton(tabFzg.getId(), bandOther.getId(), BUTTON_CHECKIN_FOLDER);
    button.setOrdinal(12);
    button.setTitle("Ordner einchecken");
    button.setIconName("ScriptButton12");
    // Ordner herunterladen
    button = ribbon.addButton(tabFzg.getId(), bandOther.getId(), BUTTON_EXPORT_FOLDER);
    button.setOrdinal(14);
    button.setTitle("Ordner herunterladen");
    button.setIconName("ScriptButton14");
    // PDFs Zusammenfügen
    button = ribbon.addButton(tabFzg.getId(), bandOther.getId(), BUTTON_MERGE_PDF);
    button.setOrdinal(37);
    button.setTitle("PDFs zusammenfügen");
    button.setIconName("ScriptButton37");

    // Info Band
	// Support
    button = ribbon.addButton(tabFzg.getId(), bandInfo.getId(), BUTTON_SUPPORT);
    button.setOrdinal(1);
    button.setTitle("Support");
    button.setIconName("ScriptButton8");
    // FZG
    button = ribbon.addButton(tabFzg.getId(), bandInfo.getId(), BUTTON_FZG);
    button.setOrdinal(8);
    button.setTitle("FZG");
    button.setCallback(function () { workspace.showInfoBox("FZG", "<html>Servus " + getNameFormatted("",FST_NAME_FST) + ",<br><br>Herzlich Willkommen im Dokumentenmanagementsystem ELO der<br>Forschungsstelle für Zahnräder und Getriebesysteme.</html>"); }, this);
    button.setIconName("ScriptButton1");

    // Adressen Band
    // Durchsuchen
    button = ribbon.addButton(tabFzg.getId(), bandAddress.getId(), BUTTON_QUERY_ADB);
    button.setOrdinal(17);
    button.setTitle("Adressen durchsuchen");
    button.setIconName("ScriptButton17");
    // Einfuegen
    button = ribbon.addButton(tabFzg.getId(), bandAddress.getId(), BUTTON_INSERT_ADB);
    button.setOrdinal(16);
    button.setTitle("Neue Adresse anlegen");
    button.setIconName("ScriptButton16");
    button.setEnabledCallback(
        function () {
            try {
                readADBCredentialsFromFile(CREDENTIAL_FILES[0]);
            } catch (e) { return false; }
            return true;
        },
        this);
    // Querbearbeitung
    button = ribbon.addButton(tabFzg.getId(), bandAddress.getId(), BUTTON_CROSSPROCESSING_ADB);
    button.setOrdinal(24);
    button.setTitle("Querbearbeitung");
    button.setIconName("ScriptButton24");
    button.setEnabledCallback(
        function () {
            try {
                readADBCredentialsFromFile(CREDENTIAL_FILES[0]);
            } catch (e) {
                return false;
            }
            return true;
        },
        this);

    // Publikationen Band
    // Neue Publikation
    button = ribbon.addButton(tabFzg.getId(), bandPublication.getId(), BUTTON_NEW_PUBLICATION);
    button.setOrdinal(21);
    button.setTitle("Publikation anlegen");
    button.setIconName("ScriptButton21");
    // Neuer Presseartikel
    button = ribbon.addButton(tabFzg.getId(), bandPublication.getId(), BUTTON_NEW_ARTICLE);
    button.setOrdinal(22);
    button.setTitle("Presseartikel anlegen");
    button.setIconName("ScriptButton22");
    // Neuer Vortrag
    button = ribbon.addButton(tabFzg.getId(), bandPublication.getId(), BUTTON_NEW_PRESENTATION);
    button.setOrdinal(23);
    button.setTitle("Vortrag anlegen");
    button.setIconName("ScriptButton23");
    // Neue Diss
    button = ribbon.addButton(tabFzg.getId(), bandPublication.getId(), BUTTON_NEW_DISS);
    button.setOrdinal(25);
    button.setTitle("Dissertation anlegen");
    button.setIconName("ScriptButton25");
	// Forschungsantrag
    button = ribbon.addButton(tabFzg.getId(), bandPublication.getId(), BUTTON_RESEARCH_PROPOSAL);
    button.setOrdinal(26);
    button.setTitle("Forschungsantrag anlegen");
    button.setIconName("ScriptButton26");
    // Neuer Forschungsbericht
    button = ribbon.addButton(tabFzg.getId(), bandPublication.getId(), BUTTON_NEW_RESEARCH_REPORT);
    button.setOrdinal(27);
    button.setTitle("Forschungsbericht anlegen");
    button.setIconName("ScriptButton27");

    // Zeichnungsarchiv Band
    // Zeichnungsablage oeffnen
    button = ribbon.addButton(tabFzg.getId(), bandSketchDepot.getId(), BUTTON_OPEN_SKETCH_DEPOT);
    button.setOrdinal(30);
    button.setTitle("Zeichnungsablage öffnen");
    button.setIconName("ScriptButton30");
    // Zeichnungen pruefen
    button = ribbon.addButton(tabFzg.getId(), bandSketchDepot.getId(), BUTTON_CHECK_SKETCH_DEPOT);
    button.setOrdinal(32);
    button.setTitle("Zeichnungsablage prüfen lassen");
    button.setIconName("ScriptButton32");
	
	// Bestellungen Band
	// Bedarfsmeldung anlegen
    button = ribbon.addButton(tabFzg.getId(), bandPurchase.getId(), BUTTON_PURCH_CREATE);
    button.setOrdinal(35);
    button.setTitle("Bedarfsmeldung anlegen");
    button.setIconName("ScriptButton35");
	// Bedarfsmeldung freigeben
    button = ribbon.addButton(tabFzg.getId(), bandPurchase.getId(), BUTTON_PURCH_CLEAR);
    button.setOrdinal(36);
    button.setTitle("Bedarfsmeldung freigeben lassen");
    button.setIconName("ScriptButton36");

    // Admin Tab
    // Global Options
    button = ribbon.addButton(tabAdmin.getId(), bandAdmin.getId(), BUTTON_GLOBAL_OPTIONS);
    button.setOrdinal(10);
    button.setTitle("Global Options");
    button.setIconName("ScriptButton10");
    // Hidden Options
    button = ribbon.addButton(tabAdmin.getId(), bandAdmin.getId(), BUTTON_HIDDEN_OPTIONS);
    button.setOrdinal(15);
    button.setTitle("Hidden Options");
    button.setIconName("ScriptButton15");
    // Ordner CheckIn/Out sperre entfernen
    button = ribbon.addButton(tabAdmin.getId(), bandAdmin.getId(), BUTTON_FOLDER_CHECK_IN_OUT_UNLOCK);
    button.setOrdinal(13);
    button.setTitle("Ordner CheckIn/Out Sperre entfernen");
    button.setIconName("ScriptButton13");

    // Vorlaeufig Tab
    // Testbutton
    button = ribbon.addButton(tabProvisional.getId(), bandProvisional.getId(), BUTTON_TEST);
    button.setOrdinal(6);
    button.setTitle("Testbutton");
    button.setIconName("ScriptButton6");
    // Bibtex
    button = ribbon.addButton(tabProvisional.getId(), bandProvisional.getId(), BUTTON_BIBTEX);
    button.setOrdinal(34);
    button.setTitle("BibTex Ausgabe");
    button.setIconName("ScriptButton34");

    // In bestehende Tabs einfuegen
    // Workflow
    //startTab = ribbon.getTab("start");
    //workflowBand = ribbon.getBand("workflow");
    // WF
    button = ribbon.addButton("workflow", "start", BUTTON_START_WF);
    button.setOrdinal(29);
    button.setTitle("Workflow starten");
    button.setIconName("ScriptButton29");
    // Adhoc WF
    button = ribbon.addButton("workflow", "start", BUTTON_START_ADHOC_WF);
    button.setOrdinal(38);
    button.setTitle("Adhoc Workflow");
    button.setIconName("ScriptButton28");
}

//@include lib_utility
//@include lib_const