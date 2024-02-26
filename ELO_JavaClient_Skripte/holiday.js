importPackage(Packages.java.io);
importPackage(Packages.javax.swing);
// .NET-Framework / COM-Schnittstelle
//importPackage(Packages.de.elo.client);
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

// Button hinzufuegen
function eloExpandRibbon() {
    createRibbon();
    newButton(BUTTON_HOLIDAY, holiday);
}

// Dummy file und Workflow anlegen bei Klick auf Button
function holiday() {

//    DIALOG_THIS.close();
    var wf = ixc.checkoutWorkFlow("Urlaubsantrag", WFTypeC.TEMPLATE, WFDiagramC.mbAll, LockC.NO);
    var wf_ws = ixc.checkoutWorkFlow("Urlaubsantrag_Werkstatt_Pr�ffeld", WFTypeC.TEMPLATE, WFDiagramC.mbAll, LockC.NO);
    var wf_pv = ixc.checkoutWorkFlow("Urlaubsantrag_Personalverantwortlicher", WFTypeC.TEMPLATE, WFDiagramC.mbAll, LockC.NO);

    var name = workspace.getUserName();
//  ELO-Benutzername aus LDAP auf ben�tigtes Format bringen
    var nameformatted = getNameFormatted(name,LST_NAME_FST_NAME);
	
    var usrdata = archive.getElementByArcpathRelative(USERS_FOLDER_ID, "�" + name);
    var sord_usrdat = usrdata.loadSord();
    var position = getIndexValueByName(sord_usrdat, ELO_USER_POS);

// Datum
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
    
    var docpath = workspace.directories.tempDir + "\\Urlaub_" + jahr + monat + tag +".txt";
    var file = utils.getUniqueFile(workspace.directories.tempDir, "Urlaub_" + jahr + monat + tag +".txt");
    utils.copyFile(archive.getElement(HOLIDAY_DUMMY).getFile(), file);
	
    var dummyparent = archive.getElementByArcpathRelative(HOLIDAY_DUMMY_PARENT_ID, "�" + nameformatted);
    
    var prepdoc = dummyparent.prepareDocument(0);
    var dummy = dummyparent.addDocument(prepdoc, docpath);
    var sord = dummy.loadSord();
    dummy.sord = sord;
    dummy.saveSord();
    dummy.setMaskId(HOLIDAY_MASK);
    dummy.saveSord();
    sord = dummy.loadSord();
    var uname = sord.getOwnerName();

//  Datei-Besitzername(=Antragsteller) aus LDAP auf ben�tigtes Format bringen
    var unameformatted2 = getNameFormatted(uname,FST_NAME_FST);
  // Dummy loeschen, wenn Abbruch
  if (!indexDialog.editSord(sord, true, "Urlaubsantrag")) {
        dummy.del();
        return;
    }

    dummy.sord = sord;
    dummy.saveSord();
  // Bei OK Workflow erstellen
  if ((position.equals("Pr�ffeldmechaniker"))|(position.equals("Pr�ffeldmechanikerin"))|(position.equals("Werkstatt"))) {
    wf_ws.type = WFTypeC.ACTIVE;
    wf_ws.id = -1;
    wf_ws.objId = dummy.getId();
    wf_ws.name = "Urlaubsantrag "+unameformatted2+" (Werkstatt/Pr�ffeld)";
    ixc.checkinWorkFlow(wf_ws, WFDiagramC.mbAll, LockC.NO);
     }else if (unameformatted2.equals("Administrator Administrator")) {
    wf_pv.type = WFTypeC.ACTIVE;
    wf_pv.id = -1;
    wf_pv.objId = dummy.getId();
    wf_pv.name = "Urlaubsantrag "+unameformatted2;
    ixc.checkinWorkFlow(wf_pv, WFDiagramC.mbAll, LockC.NO);    
     }else{
    wf.type = WFTypeC.ACTIVE;
    wf.id = -1;
    wf.objId = dummy.getId();
    wf.name = "Urlaubsantrag "+unameformatted2;
    ixc.checkinWorkFlow(wf, WFDiagramC.mbAll, LockC.NO);
    }
    // Meldung an User
    workspace.showInfoBox("Urlaubsantrag", HOLIDAY_INFO1);
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

//  Datei-Besitzername(=Antragsteller) aus LDAP auf ben�tigtes Format bringen
    var unameformatted = getNameFormatted(uname,LST_NAME_FST_NAME);  

    var date = sord.getIDateIso()
    var vert  = getIndexValueByName(sord, HOLIDAY_NAME_VERT);

// �nderungen im Dialog des Workflowfensters gleich in Verschlagwortung abspeichern und verwenden
    var sord_dialog  = dialog.getSord();
    var start = getIndexValueByName(sord_dialog, HOLIDAY_DATE_START);
    var end   = getIndexValueByName(sord_dialog, HOLIDAY_DATE_END);
 if (start > 0){ 
//�berbraten
    setIndexValueByName(sord, HOLIDAY_DATE_START, start);
    setIndexValueByName(sord, HOLIDAY_DATE_END, end);
    item.setSord(sord);
 }



   if ((sucName.equals(HOLIDAY_WF_VIEW_BLADL))|(sucName.equals(HOLIDAY_WF_CONFIRM_AL))|(sucName.equals(HOLIDAY_WF_CONFIRM_PF))|(sucName.equals(HOLIDAY_WF_DECLINE))|(sucName.equals(HOLIDAY_WF_DECLINE_PF))|(sucName.equals(HOLIDAY_WF_CONFIRM_PF_FEST))){
//Pr�fen ob HolidayBladl (durch Konny) ausgecheckt ist.
    var holidayFolder = archive.getElementByArcpathRelative(HOLIDAY_ARCHIVE_ID, "�" + unameformatted);
    var holidayBladl = archive.getElementByArcpathRelative(holidayFolder.id, "�" + "Urlaub");
   if (holidayBladl.isLocked()){
        workspace.showInfoBox("Urlaubsantrag", "Das zugeh�rige Urlaubsblatt wird gerade von Frau G�th bearbeitet. Bitte Frau G�th informieren oder sp�ter erneut versuchen.");
        return -1;	
   }
//Pr�fen ob HolidayBladl bereits voll ist.
    var elopdf = holidayBladl;
    var fieldnames_read_last_line = ['vomamRow21'];
    var entries_read_last_line =[];
    var verbose = true;
readPDFFields(elopdf, fieldnames_read_last_line, entries_read_last_line, verbose)
//Pr�fen, ob neue Zeile ben�tigt wird, oder ob Urlaubszeitraum schon eingetragen ist (f�r Genehmigung Pf nach AL)     
    var ss = start.split('');
    var es = end.split('');
    var startformat = ss[6]+ss[7]+"."+ss[4]+ss[5]+"."+ss[0]+ss[1]+ss[2]+ss[3];
    var endformat   = es[6]+es[7]+"."+es[4]+es[5]+"."+es[0]+es[1]+es[2]+es[3];
    
var fieldnames_read_2 = ['bisRow21'];
    var entries_read_2 =[];
    var verbose = true;
    readPDFFields(elopdf, fieldnames_read_2, entries_read_2, verbose)
//Wenn Blatt voll dann Benachrichtigung, dass Konny informiert werden muss
   if ((!entries_read_last_line[0].isEmpty())&&
      (!((entries_read_last_line[0].equals(startformat))&&
         (entries_read_2[0].equals(endformat))))){
        workspace.showInfoBox("Urlaubsantrag", "Das zugeh�rige Urlaubsblatt ist voll. Bitte Frau G�th informieren, dass neues Urlaubsblatt angelegt werden muss und anschlie�end erneut versuchen.");
        return -1;	
   }
}

//Pr�fung, ob Resturlaub gleich 0. Sonst keine Best�tigung durch AL m�glich
   if (sucName.equals(HOLIDAY_WF_CONFIRM_AL)){


    var ss = start.split('');
    var es = end.split('');
    var startformat = ss[6]+ss[7]+"."+ss[4]+ss[5]+"."+ss[0]+ss[1]+ss[2]+ss[3];
    var endformat   = es[6]+es[7]+"."+es[4]+es[5]+"."+es[0]+es[1]+es[2]+es[3];

    var elopdf = holidayBladl;

//PDF Felder einlesen
    var fieldnames_read_1 = [];
    var fieldnames_read_2 = [];
    for (var i = 1; i < 22; i++) {
    fieldnames_read_1[i-1] = 'vomamRow'+i;
    fieldnames_read_2[i-1] = 'bisRow'+i;
    }
    var entries_read_1 =[];
    var entries_read_2 =[];
    var verbose = true;

    readPDFFields(elopdf, fieldnames_read_1, entries_read_1, verbose)
    readPDFFields(elopdf, fieldnames_read_2, entries_read_2, verbose)

//Pr�fen, ob Datumsbereich schon eingetragen.
	var i = 0;
	while ((!((entries_read_1[i].equals(startformat))&&
		 (entries_read_2[i].equals(endformat)))&&
		(i < entries_read_1.length))&&!(entries_read_1[i].isEmpty())) {i++}
//F�r Nummerierung von Feldern i um 1 erh�hen
//Hier nicht erh�hen, weil Zeile davor angeschaut wird
//    i = i+1;
//Pr�fen, ob Resturlaub gleich 0   
//Nur pr�fen wenn nicht erste Zeile
 if (i > 0){
    var fieldnames_read_rest = ['RestRow'+i];
    var entries_read_rest =[];
    var verbose = true;
    readPDFFields(elopdf, fieldnames_read_rest, entries_read_rest, verbose)
    if (!entries_read_rest[0].isEmpty()){
        if (entries_read_rest[0] == 0 ){
	var restisnull = true;
    }
    }
    }else{
	var restisnull = false;
    } 
//Wenn Resturlaub dann Benachrichtigung, dass Konny informiert werden muss
 if (restisnull){
        workspace.showInfoBox("Urlaubsantrag", "Der Resturlaub des zugeh�rigen Urlaubsblatts ist null. Bitte Frau G�th informieren, dass neues Urlaubsblatt angelegt werden muss und anschlie�end erneut versuchen.");
        return -1;
   }


}

// Urlaubsblatt ansehen
   if (sucName.equals(HOLIDAY_WF_VIEW_BLADL)) {



    var ds = date.split('');	
    var ss = start.split('');
    var es = end.split('');
    var dateformat =  ds[6]+ds[7]+"."+ds[4]+ds[5]+"."+ds[0]+ds[1]+ds[2]+ds[3];
    var startformat = ss[6]+ss[7]+"."+ss[4]+ss[5]+"."+ss[0]+ss[1]+ss[2]+ss[3];
    var endformat   = es[6]+es[7]+"."+es[4]+es[5]+"."+es[0]+es[1]+es[2]+es[3];
    var cal = Calendar.getInstance();
    var datetoday = formatDateAsString(cal.get(Calendar.DAY_OF_MONTH), cal.get(Calendar.MONTH) + 1, cal.get(Calendar.YEAR), false);
    var elopdf = holidayBladl;

//PDF Felder einlesen
    var fieldnames_read_1 = [];
    var fieldnames_read_2 = [];
    for (var i = 1; i < 22; i++) {
    fieldnames_read_1[i-1] = 'vomamRow'+i;
    fieldnames_read_2[i-1] = 'bisRow'+i;
    }
    var entries_read_1 =[];
    var entries_read_2 =[];
    var verbose = true;

    readPDFFields(elopdf, fieldnames_read_1, entries_read_1, verbose)
    readPDFFields(elopdf, fieldnames_read_2, entries_read_2, verbose)

//Pr�fen, ob Datumsbereich schon eingetragen.
	var i = 0;
	while ((!((entries_read_1[i].equals(startformat))&&
		 (entries_read_2[i].equals(endformat)))&&
		(i < entries_read_1.length))&&!(entries_read_1[i].isEmpty())) {i++}
//Wenn Urlaubszeitraum schon eingetragen, dann nur pdf �ffnen
	if (!(entries_read_1[i].isEmpty())){
	elopdf.open();
//Ansonsten eintragen
	} else {
//F�r Zeilennummerierung in PDF	
        i = i+1;
//Jetzt PDF ausf�llen
    var fieldnames = ['vomamRow'+i,'bisRow'+i,'VertreterinRow'+i,'AntragstellerinRow'+i];
    var entries = [startformat,endformat,vert,dateformat+", "+uname];
    var openTemp = true;
    var addVersion = false;
    var versionnumber = "0.0";
    var versiontitle = "Urlaubsblatt ge�ffnet �ber Workflow durch AL";
    var verbose = true;
    fillPDFTemplate(elopdf, fieldnames, entries, openTemp, addVersion, versionnumber, versiontitle, verbose)


}
        return -1;
}

// Bestaetigung durch Abteilungsleiter
   if (sucName.equals(HOLIDAY_WF_CONFIRM_AL)) {

	
    var ds = date.split('');	
    var ss = start.split('');
    var es = end.split('');
    var dateformat =  ds[6]+ds[7]+"."+ds[4]+ds[5]+"."+ds[0]+ds[1]+ds[2]+ds[3];
    var startformat = ss[6]+ss[7]+"."+ss[4]+ss[5]+"."+ss[0]+ss[1]+ss[2]+ss[3];
    var endformat   = es[6]+es[7]+"."+es[4]+es[5]+"."+es[0]+es[1]+es[2]+es[3];
    var cal = Calendar.getInstance();
    var datetoday = formatDateAsString(cal.get(Calendar.DAY_OF_MONTH), cal.get(Calendar.MONTH) + 1, cal.get(Calendar.YEAR), false);


    var elopdf = holidayBladl;

//PDF Felder einlesen
    var fieldnames_read_1 = [];
    var fieldnames_read_2 = [];
    for (var i = 1; i < 22; i++) {
    fieldnames_read_1[i-1] = 'vomamRow'+i;
    fieldnames_read_2[i-1] = 'bisRow'+i;
    }
    var entries_read_1 =[];
    var entries_read_2 =[];
    var verbose = true;

    readPDFFields(elopdf, fieldnames_read_1, entries_read_1, verbose)
    readPDFFields(elopdf, fieldnames_read_2, entries_read_2, verbose)

//Pr�fen, ob Datumsbereich schon eingetragen.
	var i = 0;
	while ((!((entries_read_1[i].equals(startformat))&&
		 (entries_read_2[i].equals(endformat)))&&
		(i < entries_read_1.length))&&!(entries_read_1[i].isEmpty())) {i++}
//Wenn Datumsbereich schon eingetragen, dann AL Genehmigung in dieser Zeile erg�nzen.
//Jetzt PDF ausf�llen
//F�r Nummerierung von Feldern i um 1 erh�hen
    i = i+1;
    var fieldnames = ['vomamRow'+i,'bisRow'+i,'VertreterinRow'+i,'AntragstellerinRow'+i,'VorgesetzterRow'+i];
    var entries = [startformat,endformat,vert,dateformat+", "+uname, datetoday+", "+getUserEntry(ELO_USER_SHORT)];
    var openTemp = false;
    var addVersion = true;
    var versionnumber = "1.1";
    var versiontitle = "Genehmigung durch AL";
    var verbose = true;

    fillPDFTemplate(elopdf, fieldnames, entries, openTemp, addVersion, versionnumber, versiontitle, verbose)

}

// Bestaetigung durch Dr. Pflaum
   if ((sucName.equals(HOLIDAY_WF_CONFIRM_PF))|(sucName.equals(HOLIDAY_WF_CONFIRM_PF_FEST))) {


	
    var ds = date.split('');	
    var ss = start.split('');
    var es = end.split('');
    var dateformat =  ds[6]+ds[7]+"."+ds[4]+ds[5]+"."+ds[0]+ds[1]+ds[2]+ds[3];
    var startformat = ss[6]+ss[7]+"."+ss[4]+ss[5]+"."+ss[0]+ss[1]+ss[2]+ss[3];
    var endformat   = es[6]+es[7]+"."+es[4]+es[5]+"."+es[0]+es[1]+es[2]+es[3];
    var cal = Calendar.getInstance();
    var datetoday = formatDateAsString(cal.get(Calendar.DAY_OF_MONTH), cal.get(Calendar.MONTH) + 1, cal.get(Calendar.YEAR), false);

    var elopdf = holidayBladl;

//PDF Felder einlesen
    var fieldnames_read_1 = [];
    var fieldnames_read_2 = [];
    for (var i = 1; i < 22; i++) {
    fieldnames_read_1[i-1] = 'vomamRow'+i;
    fieldnames_read_2[i-1] = 'bisRow'+i;
    }
    var entries_read_1 =[];
    var entries_read_2 =[];
    var verbose = true;

    readPDFFields(elopdf, fieldnames_read_1, entries_read_1, verbose)
    readPDFFields(elopdf, fieldnames_read_2, entries_read_2, verbose)

//Pr�fen, ob Datumsbereich schon eingetragen.
	var i = 0;
	while ((!((entries_read_1[i].equals(startformat))&&
		 (entries_read_2[i].equals(endformat)))&&
		(i < entries_read_1.length))&&!(entries_read_1[i].isEmpty())) {i++}
//Wenn Datumsbereich schon eingetragen, dann Pf Genehmigung in dieser Zeile erg�nzen.

//Jetzt PDF ausf�llen
//F�r Nummerierung von Feldern i um 1 erh�hen
    i = i+1;
//Zuerst pr�fen, ob ein AL schon freigegeben hat
    var fieldnames_read_al = ['VorgesetzterRow'+i];
    var entries_read_al =[];
    var verbose = true;

    readPDFFields(elopdf, fieldnames_read_al, entries_read_al, verbose)
//Wenn kein AL freigegeben hat dann alles ausf�llen und heutiges Datum und Pf K�rzel eintragen
    if (entries_read_al[0].isEmpty()){

    var fieldnames = ['vomamRow'+i,'bisRow'+i,'VertreterinRow'+i,'AntragstellerinRow'+i,'VorgesetzterRow'+i];
    var entries = [startformat,endformat,vert,dateformat+", "+uname, datetoday+", "+getUserEntry(ELO_USER_SHORT)];
    var openTemp = false;
    var addVersion = true;
    var versionnumber = "1.3";
    var versiontitle = "Freigabe durch PV und Version versendet an MA";
    var verbose = true;

    fillPDFTemplate(elopdf, fieldnames, entries, openTemp, addVersion, versionnumber, versiontitle, verbose)

//Wenn AL bereits Urlaub genehmigt hat dann nur K�rzel pf erg�nzen.
}else{
    var fieldnames = ['PFGenehmigungRow'+i];
    var entries = [getUserEntry(ELO_USER_SHORT)];
    var openTemp = false;
    var addVersion = true;
    var versionnumber = "1.3";
    var versiontitle = "Freigabe durch PV und Version versendet an MA";
    var verbose = true;

    fillPDFTemplate(elopdf, fieldnames, entries, openTemp, addVersion, versionnumber, versiontitle, verbose)

}

// E-Mail an Mitarbeiter/in mit Anhang des Urlaubsblatts
//Mail f�r WiMA
if (sucName.equals(HOLIDAY_WF_CONFIRM_PF)) {
    var usrdata = archive.getElementByArcpathRelative(USERS_FOLDER_ID, "�" + uname);
    var sord = usrdata.loadSord();
    email = getIndexValueByName(sord, ELO_USER_MAIL) + ";";
    var body = "<html><BODY style=font-size:11pt;>Servus,<br><br>hier ist Ihr ELO.<br>";
    body += "Ihr Urlaub wurde genehmigt! Anbei finden Sie das Urlaubsblatt als .pdf f�r Ihre Unterlagen.<br>";
    body += "Der zugeh�rige Workflow befindet sich jetzt wieder bei Ihnen und kann in ELO beendet werden.";
    body += "</body></html>";
    var name = holidayBladl.getName()
    var datetoday2 = formatDateAsString2(cal.get(Calendar.DAY_OF_MONTH), cal.get(Calendar.MONTH) + 1, cal.get(Calendar.YEAR));
    holidayBladl.setName(name+"_"+datetoday2)
    mail(email, WORKFLOW_MAIL_SUBJECT, body, [attachFile(holidayBladl)], true);

// Mail f�r Festangestellte
}else{
// E-Mail an Mitarbeiter/in mit Anhang des Urlaubsblatts
    var usrdata = archive.getElementByArcpathRelative(USERS_FOLDER_ID, "�" + uname);
    var sord = usrdata.loadSord();
    email = getIndexValueByName(sord, ELO_USER_MAIL) + ";";
    var body = "<html><BODY style=font-size:11pt;>Servus,<br><br>hier ist Ihr ELO.<br>";
    body += "Ihr Urlaub wurde genehmigt! Anbei finden Sie das Urlaubsblatt als .pdf f�r Ihre Unterlagen.<br>";
    body += "Ihr Urlaub wird jetzt noch in Zeus eingetragen. Sie erhalten eine E-Mail, sobald Ihr Urlaub in Zeus eingetragen ist.";
    body += "</body></html>";
    var name = holidayBladl.getName()
    var datetoday2 = formatDateAsString2(cal.get(Calendar.DAY_OF_MONTH), cal.get(Calendar.MONTH) + 1, cal.get(Calendar.YEAR));
    holidayBladl.setName(name+"_"+datetoday2)
    mail(email, WORKFLOW_MAIL_SUBJECT, body, [attachFile(holidayBladl)], true);
}

 if (sucName.equals(HOLIDAY_WF_CONFIRM_PF)) {
//Pr�fen, ob Blatt voll   
    var elopdf = holidayBladl;
    var fieldnames_read_last_line = ['vomamRow21'];
    var entries_read_last_line =[];
    var verbose = true;
    readPDFFields(elopdf, fieldnames_read_last_line, entries_read_last_line, verbose)
//Pr�fen, ob Resturlaub gleich 0   
    var fieldnames_read_rest = ['RestRow'+i];
    var entries_read_rest =[];
    var verbose = true;
    readPDFFields(elopdf, fieldnames_read_rest, entries_read_rest, verbose)
    if (!entries_read_rest[0].isEmpty()){
        if (entries_read_rest[0] == 0 ){
	var restisnull = true;
    }
    }else{
	var restisnull = false;
    } 
//Wenn Blatt voll oder Resturlaub gleich 0 dann Mail an Konny, dass neues Blatt angelegt werden muss
 if ((!entries_read_last_line[0].isEmpty())|(restisnull)){
    email = "zander@fzg.mw.tum.de;" //Hier Mailaddresse von Frau Gueth eintragen
    var body = "<html><BODY style=font-size:11pt;>Servus,<br><br>hier ist Ihr ELO.<br>";
    body += "Erinnerung: Das Urlaubsblatt eines Mitarbeiters ist voll! Bitte volles Urlaubsblatt archivieren und neues Urlaubsblatt anlegen f�r: "+ uname;
    body += "</body></html>";
    mail(email, WORKFLOW_MAIL_SUBJECT, body, [], true);
   }
 }
}

// Urlaubsblatt ansehen durch Andrea Baur (nur oeffnen, nichts eintragen)
   if (sucName.equals(HOLIDAY_WF_OPEN_FEST_BLADL)) {

    var holidayFolder = archive.getElementByArcpathRelative(HOLIDAY_ARCHIVE_ID, "�" + unameformatted);
    var holidayBladl = archive.getElementByArcpathRelative(holidayFolder.id, "�" + "Urlaub");
    var elopdf = holidayBladl;
    elopdf.open();
    return -1;
}

// E-Mail an Mitarbeiter/in durch Andrea Baur, dass Urlaub in Zeus eingetragen wurde
   if (sucName.equals(HOLIDAY_WF_CONFIRM_ZEUS)) {
    var usrdata = archive.getElementByArcpathRelative(USERS_FOLDER_ID, "�" + uname);
    var sord = usrdata.loadSord();
    email = getIndexValueByName(sord, ELO_USER_MAIL) + ";";
    var body = "<html><BODY style=font-size:11pt;>Servus,<br><br>hier ist Ihr ELO.<br>";
    body += "Ihr genehmigter Urlaub wurde soeben in Zeus eingetragen. Ihr Urlaubsantrag ist jetzt vollst�ndig abgeschlossen. Der zugeh�rige Workflow befindet sich jetzt wieder bei Ihnen und kann in ELO beendet werden.";
    body += "</body></html>";
    mail(email, WORKFLOW_MAIL_SUBJECT, body, [], true);

// Pr�fen ob Blatt voll ist oder Resturlaub gleich 0. Falls ja, Erinnerungs-Mail an Konny
    var holidayFolder = archive.getElementByArcpathRelative(HOLIDAY_ARCHIVE_ID, "�" + unameformatted);
    var holidayBladl = archive.getElementByArcpathRelative(holidayFolder.id, "�" + "Urlaub");
    var elopdf = holidayBladl;

    var ds = date.split('');	
    var ss = start.split('');
    var es = end.split('');
    var dateformat =  ds[6]+ds[7]+"."+ds[4]+ds[5]+"."+ds[0]+ds[1]+ds[2]+ds[3];
    var startformat = ss[6]+ss[7]+"."+ss[4]+ss[5]+"."+ss[0]+ss[1]+ss[2]+ss[3];
    var endformat   = es[6]+es[7]+"."+es[4]+es[5]+"."+es[0]+es[1]+es[2]+es[3];
    var cal = Calendar.getInstance();
    var datetoday = formatDateAsString(cal.get(Calendar.DAY_OF_MONTH), cal.get(Calendar.MONTH) + 1, cal.get(Calendar.YEAR), false);


//PDF Felder einlesen
    var fieldnames_read_1 = [];
    var fieldnames_read_2 = [];
    for (var i = 1; i < 22; i++) {
    fieldnames_read_1[i-1] = 'vomamRow'+i;
    fieldnames_read_2[i-1] = 'bisRow'+i;
    }
    var entries_read_1 =[];
    var entries_read_2 =[];
    var verbose = true;

    readPDFFields(elopdf, fieldnames_read_1, entries_read_1, verbose)
    readPDFFields(elopdf, fieldnames_read_2, entries_read_2, verbose)

//Pr�fen, ob Datumsbereich schon eingetragen.
	var i = 0;
	while ((!((entries_read_1[i].equals(startformat))&&
		 (entries_read_2[i].equals(endformat)))&&
		(i < entries_read_1.length))&&!(entries_read_1[i].isEmpty())) {i++}
//Wenn Datumsbereich schon eingetragen, dann Pf Genehmigung in dieser Zeile erg�nzen.

//F�r Nummerierung von Feldern i um 1 erh�hen
    i = i+1;

//Pr�fen, ob Blatt voll   
    var elopdf = holidayBladl;
    var fieldnames_read_last_line = ['vomamRow21'];
    var entries_read_last_line =[];
    var verbose = true;
    readPDFFields(elopdf, fieldnames_read_last_line, entries_read_last_line, verbose)
//Pr�fen, ob Resturlaub gleich 0   
    var fieldnames_read_rest = ['RestRow'+i];
    var entries_read_rest =[];
    var verbose = true;
    readPDFFields(elopdf, fieldnames_read_rest, entries_read_rest, verbose)
    if (!entries_read_rest[0].isEmpty()){
        if (entries_read_rest[0] == 0 ){
	var restisnull = true;
    }
    }else{
	var restisnull = false;
    } 
//Wenn Blatt voll oder Resturlaub gleich 0 dann Mail an Konny, dass neues Blatt angelegt werden muss
 if ((!entries_read_last_line[0].isEmpty())|(restisnull)){
    email = "zander@fzg.mw.tum.de;" //Hier Mailaddresse von Frau Gueth eintragen
    var body = "<html><BODY style=font-size:11pt;>Servus,<br><br>hier ist Ihr ELO.<br>";
    body += "Erinnerung: Das Urlaubsblatt eines Mitarbeiters ist voll! Bitte volles Urlaubsblatt archivieren und neues Urlaubsblatt anlegen f�r: "+ uname;
    body += "</body></html>";
    mail(email, WORKFLOW_MAIL_SUBJECT, body, [], true);
   }
}

// Urlaub ablehnen
  if ((sucName.equals(HOLIDAY_WF_DECLINE))|(sucName.equals(HOLIDAY_WF_DECLINE_PF))) {


	
	
    var ds = date.split('');	
    var ss = start.split('');
    var es = end.split('');
    var dateformat =  ds[6]+ds[7]+"."+ds[4]+ds[5]+"."+ds[0]+ds[1]+ds[2]+ds[3];
    var startformat = ss[6]+ss[7]+"."+ss[4]+ss[5]+"."+ss[0]+ss[1]+ss[2]+ss[3];
    var endformat   = es[6]+es[7]+"."+es[4]+es[5]+"."+es[0]+es[1]+es[2]+es[3];
    var cal = Calendar.getInstance();
    var datetoday = formatDateAsString(cal.get(Calendar.DAY_OF_MONTH), cal.get(Calendar.MONTH) + 1, cal.get(Calendar.YEAR), false);

    var elopdf = holidayBladl;

//PDF Felder einlesen
    var fieldnames_read_1 = [];
    var fieldnames_read_2 = [];
    for (var i = 1; i < 22; i++) {
    fieldnames_read_1[i-1] = 'vomamRow'+i;
    fieldnames_read_2[i-1] = 'bisRow'+i;
    }
    var entries_read_1 =[];
    var entries_read_2 =[];
    var verbose = true;

    readPDFFields(elopdf, fieldnames_read_1, entries_read_1, verbose)
    readPDFFields(elopdf, fieldnames_read_2, entries_read_2, verbose)

//Pr�fen, ob Datumsbereich schon eingetragen.
	var i = 0;
	while ((!((entries_read_1[i].equals(startformat))&&
		 (entries_read_2[i].equals(endformat)))&&
		(i < entries_read_1.length))&&!(entries_read_1[i].isEmpty())) {i++}
//Wenn Datumsbereich schon eingetragen, dann Pf Genehmigung in dieser Zeile erg�nzen.

//Jetzt PDF ausf�llen
//F�r Nummerierung von Feldern i um 1 erh�hen
    i = i+1;

//"Nicht genehmigt!" dr�berbuttern
    var fieldnames = ['vomamRow'+i,'bisRow'+i,'VertreterinRow'+i,'AntragstellerinRow'+i,'VorgesetzterRow'+i];
    var entries = [startformat,endformat,vert,dateformat+", "+uname, "Nicht genehmigt!"];
    var openTemp = false;
    var addVersion = true;
    var versionnumber = "1.3";
    var versiontitle = "Urlaub nicht genehmigt";
    var verbose = true;

    fillPDFTemplate(elopdf, fieldnames, entries, openTemp, addVersion, versionnumber, versiontitle, verbose)
//Pr�fen, ob Blatt voll   
    var elopdf = holidayBladl;
    var fieldnames_read_last_line = ['vomamRow21'];
    var entries_read_last_line =[];
    var verbose = true;
    readPDFFields(elopdf, fieldnames_read_last_line, entries_read_last_line, verbose)
// Hier wird keine Pr�fung, ob Resturlaub gleich 0 ist, ben�tigt. Weil Resturlaub beim ablehnen sich ja nicht verringert.
//Wenn Blatt voll dann Mail an Konny, dass neues Blatt angelegt werden muss
 if (!entries_read_last_line[0].isEmpty()){
    var usrdata = archive.getElementByArcpathRelative(USERS_FOLDER_ID, "�" + "Urlaub");
    var sord = usrdata.loadSord();
    email = "gueth@fzg.mw.tum.de;" //Hier Mailaddresse von Frau Gueth eintragen
    var body = "<html><BODY style=font-size:11pt;>Servus,<br><br>hier ist Ihr ELO.<br>";
    body += "Erinnerung: Das Urlaubsblatt eines Mitarbeiters ist voll! Bitte volles Urlaubsblatt archivieren und neues Urlaubsblatt anlegen f�r: "+ uname;
    body += "</body></html>";
    mail(email, WORKFLOW_MAIL_SUBJECT, body, [], true);
   }

//Dummy file l�schen
    var dialog = dialogs.flowConfirmDialog;
    var item = tasksViews.getTasksViewForWorkflow(dialog.getFlowId()).getWorkflow(dialog.getFlowId()).getArchiveElement();
    item.del();
}


   if (sucName.equals(HOLIDAY_WF_OPEN_BLADL)) {
	
    var holidayFolder = archive.getElementByArcpathRelative(HOLIDAY_ARCHIVE_ID, "�" + unameformatted);
    var holidayBladl = archive.getElementByArcpathRelative(holidayFolder.id, "�" + "Urlaub");
    workspace.updateSordLists();

   if (!holidayBladl.isLocked()){

    holidayBladl.checkOut();
    utils.editFile(checkout.getLastDocument().getDocumentFile());
//    workspace.gotoId(holidayFolder.id);	
}else{
    utils.editFile(holidayBladl.getDocumentFile());
//    workspace.gotoId(holidayFolder.id);	
}
return -1;
}

//Pr�fen, ob alle Zeilen durch Konny ausgef�llt sind
   if ((sucName.equals(HOLIDAY_WF_CHECKIN_BLADL))|(sucName.equals(HOLIDAY_WF_FEST_BLADL))|(sucName.equals(HOLIDAY_WF_REMOVE_HOL))|(sucName.equals(HOLIDAY_WF_TOOMUCH))) {

    var holidayFolder = archive.getElementByArcpathRelative(HOLIDAY_ARCHIVE_ID, "�" + unameformatted);
    var holidayBladl = archive.getElementByArcpathRelative(holidayFolder.id, "�" + "Urlaub");
    workspace.updateSordLists();

    var ss = start.split('');
    var es = end.split('');
    var startformat = ss[6]+ss[7]+"."+ss[4]+ss[5]+"."+ss[0]+ss[1]+ss[2]+ss[3];
    var endformat   = es[6]+es[7]+"."+es[4]+es[5]+"."+es[0]+es[1]+es[2]+es[3];

    var elopdf = holidayBladl;

//PDF Felder einlesen
    var fieldnames_read_1 = [];
    var fieldnames_read_2 = [];
    for (var i = 1; i < 22; i++) {
    fieldnames_read_1[i-1] = 'vomamRow'+i;
    fieldnames_read_2[i-1] = 'bisRow'+i;
    }
    var entries_read_1 =[];
    var entries_read_2 =[];
    var verbose = true;

    readPDFFields(elopdf, fieldnames_read_1, entries_read_1, verbose)
    readPDFFields(elopdf, fieldnames_read_2, entries_read_2, verbose)

//Pr�fen, ob Datumsbereich schon eingetragen.
	var i = 0;
	while ((!((entries_read_1[i].equals(startformat))&&
		 (entries_read_2[i].equals(endformat)))&&
		(i < entries_read_1.length))&&!(entries_read_1[i].isEmpty())) {i++}
//F�r Nummerierung von Feldern i um 1 erh�hen
    i = i+1;
    var fieldnames_read_konny = [];
    fieldnames_read_konny[0] = ['ArbeitstageRow'+i];
    fieldnames_read_konny[1] = ['UrlaubskarteiRow'+i];
    fieldnames_read_konny[2] = ['RestRow'+i];

    var entries_read_konny =[];
    var verbose = true;
   if (holidayBladl.isLocked()){
       readPDFFieldsCheckout(elopdf, fieldnames_read_konny, entries_read_konny, verbose)
    }else{
       readPDFFields(elopdf, fieldnames_read_konny, entries_read_konny, verbose)
    }
//Wenn nicht alle Felder ausgef�llt dann Meldung an Konny und Abbruch.
    if ((entries_read_konny[0].isEmpty())|(entries_read_konny[1].isEmpty())|(entries_read_konny[2].isEmpty())){
        workspace.showInfoBox("Urlaubsantrag", "F�r diesen Urlaub sind nicht alle erforderlichen Felder (Urlaubstage, Eintragung in die Urlaubskartei, Resturlaub) ausgef�llt. Bitte zuerst alle erforderlichen Felder ausf�llen und dann erneut versuchen.");
        return -1;
//Sonst einchecken und weiterleiten
    }else{
    holidayBladl.checkIn("1.2", "Urlaubsblatt durch Frau G�th bearbeitet");
    workspace.updateSordLists();
//Wenn Resturlaub zu wenig dann noch Mail an Mitarbeiter
    if (sucName.equals(HOLIDAY_WF_TOOMUCH)) {
    var usrdata = archive.getElementByArcpathRelative(USERS_FOLDER_ID, "�" + uname);
    var sord = usrdata.loadSord();
    email = getIndexValueByName(sord, ELO_USER_MAIL) + ";";
    var body = "<html><BODY style=font-size:11pt;>Servus,<br><br>hier ist Ihr ELO.<br>";
    body += "Ihr Urlaub konnte nicht eingetragen werden. Es wurde mehr Urlaub beantragt als Resturlaub auf diesem Urlaubsblatt verf�gbar ist. Bitte Zeitraum korrigieren und neuen Antrag stellen. Falls zus�tzlicher Urlaub mit einem neuen Urlaubsblatt genehmigt werden soll, bitte zuerst Urlaub aus aktuellem Urlaubsblatt verbrauchen und nach Genehmigung zweiten Urlaubsantrag stellen. Anbei finden Sie das Urlaubsblatt als .pdf f�r Ihre Unterlagen.";
    body += "</body></html>";
    var name = holidayBladl.getName()
    var cal = Calendar.getInstance();
    var datetoday2 = formatDateAsString2(cal.get(Calendar.DAY_OF_MONTH), cal.get(Calendar.MONTH) + 1, cal.get(Calendar.YEAR));
    holidayBladl.setName(name+"_"+datetoday2)
    mail(email, WORKFLOW_MAIL_SUBJECT, body, [attachFile(holidayBladl)], true);
	
//Dummy file l�schen
    var dialog = dialogs.flowConfirmDialog;
    var item = tasksViews.getTasksViewForWorkflow(dialog.getFlowId()).getWorkflow(dialog.getFlowId()).getArchiveElement();
    item.del();

}

}

}

// Dummy file loeschen
  if (sucName.equals(HOLIDAY_WF_END)) {

    var dialog = dialogs.flowConfirmDialog;
    var item = tasksViews.getTasksViewForWorkflow(dialog.getFlowId()).getWorkflow(dialog.getFlowId()).getArchiveElement();
    item.del();

}
    return retval;
}

// Felder in PDF Datei ausefuellen
    function fillPDFTemplate(elopdf, fieldnames, entries, openTemp, addVersion, versionnumber, versiontitle, verbose) {
    if (verbose == undefined) { verbose = true; }

    // Pdf �ffnen
    var pdf = PDDocument.load(elopdf.getFile());
    var pdfCat = pdf.getDocumentCatalog();
    var temppath = utils.getUniqueFile(workspace.directories.tempDir, "Urlaubsblatt_neuer_Antrag.pdf");
    var acroForm = pdfCat.getAcroForm();
    if (acroForm == null) {
        return false;
    }

    // Felder bef�llen
    for (var i = 0; i < fieldnames.length; i++) {
        try {
            acroForm.getField(fieldnames[i]).setValue(entries[i]);
        } catch (e) {
            if (verbose) { workspace.showInfoBox("", e); }
        }
    }

    // speichern und schlie�en
    pdf.save(temppath);
    if (openTemp){
    Desktop.getDesktop().open(temppath);
}
    if (addVersion){
    elopdf.addVersion(temppath, versionnumber, versiontitle, false, true);
}
}

// Felder aus PDF Datei auslesen
    function readPDFFields(elopdf, fieldnames, entries, verbose) {
    if (verbose == undefined) { verbose = true; }

    // Pdf �ffnen
    var pdf = PDDocument.load(elopdf.getFile());
    var pdfCat = pdf.getDocumentCatalog();
    var temppath = utils.getUniqueFile(workspace.directories.tempDir, "Urlaubsblatt_neuer_Antrag.pdf");
    var acroForm = pdfCat.getAcroForm();
    if (acroForm == null) {
        return false;
    }

    // Felder bef�llen
    for (var i = 0; i < fieldnames.length; i++) {
        try {
            entries[i] = acroForm.getField(fieldnames[i]).getValueAsString();
        } catch (e) {
            if (verbose) { workspace.showInfoBox("", e); }
        }
    }


}

// Felder aus ausgecheckter PDF Datei lesen
    function readPDFFieldsCheckout(elopdf, fieldnames, entries, verbose) {
    if (verbose == undefined) { verbose = true; }

    // Pdf �ffnen
    var pdf = PDDocument.load(elopdf.getDocumentFile());
    var pdfCat = pdf.getDocumentCatalog();
    var temppath = utils.getUniqueFile(workspace.directories.tempDir, "Urlaubsblatt_neuer_Antrag.pdf");
    var acroForm = pdfCat.getAcroForm();
    if (acroForm == null) {
        return false;
    }

    // Felder bef�llen
    for (var i = 0; i < fieldnames.length; i++) {
        try {
            entries[i] = acroForm.getField(fieldnames[i]).getValueAsString();
        } catch (e) {
            if (verbose) { workspace.showInfoBox("", e); }
        }
    }


}

// Datum formatieren
function formatDateAsString2(day, month, year) {
    var date = "";

    if (year < 10) { date += "200"; }
    else if (year < 50) { date += "20"; }
    else if (year < 100) { date += "19"; }
    date += year;

    if (month < 10) { date += "0"; }
    date += month;

    if (day < 10) { date += "0"; }
    date += day; 


    return date;
}

//@include lib_buttons
//@include lib_const
//@include lib_utility
