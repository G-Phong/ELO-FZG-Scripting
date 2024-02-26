//////////////////////////////////////////////////////////////////////////
//  //////  //          //  //  ////////         /////////////////////////
//  //////  //////  //////  //  ////////  ////////////////////////////////
//  //////  //////  //////  //  ////////         /////////////////////////
//  //////  //////  //////  //  ///////////////  /////////////////////////
//          //////  //////  //        //         /////////////////////////
//////////////////////////////////////////////////////////////////////////

// Java Lib
importPackage(Packages.java.io);
importPackage(Packages.java.util);
importPackage(Packages.java.awt);
importPackage(Packages.javax.swing);
importPackage(Packages.java.util);
importPackage(Packages.java.nio.file);
importPackage(Packages.java.lang);
importPackage(Packages.java.net);
// ELO
importPackage(Packages.de.elo.ix.jscript);
importPackage(Packages.de.elo.client);
importPackage(Packages.de.elo.ix.client);

/**
 * 
 * @param {String[]} array
 *
 * Output: Kürzestes Element in 'array'
 */
function shortestString(array) {
    var shortest = array[0];
    for (var i = 1; i < array.length; i++) {
        if (array[i].length < shortest) { shortest = array[i]; }
    }
    return shortest;
}

/**
 * 
 * @param {any} array
 * @param {any} elem
 *
 * Output: True, falls 'elem' in 'array' vorkommt, sonst False
 */
function contains(array, elem) {
    for (var i = 0; i < array.length; i++) {
        if (array[i] == elem) { return true; }
    }
    return false;
}

/**
 * 
 * @param {Sord} sord : Sord Objekt, welches befüllt werden soll
 * @param {String} name : Name des ObjectKeys; diese Namen werden in der 'Adminconsole' unter 'Verschlagwortungs Masken bearbeiten' angelegt. 
 * @param {String} value : Zu setzender Wert
 *
 * Output : ToDo: return False, falls es keinen entsprechenden ObjectKey gibt
 */
function setIndexValueByName(sord, name, value) {
    var objKeys = sord.objKeys;
    //iteration über alle ObjectKeys
    for (var i = 0; i < objKeys.length; i++) {
        var key = objKeys[i];
        if (key.name == name) {
            key.data = [value];
        } else {
        }
    }
}

/**
 * 
 * @param {Sord} sord
 * @param {String} name
 *
 * Output : Liefert den Eintrag 'name' (Gruppe) aus Sord (Verschlagwortung)
 */
function getIndexValueByName(sord, name) {
    // objKeys := Einträge
    var objKeys = sord.objKeys;
    // Iteration über Einträge 
    for (var i = 0; i < objKeys.length; i++) {
        var key = objKeys[i];
        // Bei match mit 'name' wird der Eintrag zurückgegeben
        if (key.name == name) {
            if (key.data.length > 0) { return String(key.data[0]); }
            // Leerer Eintrag wird nicht als "" übergeben
            else { return ""; }
        } //else { workspace.showInfoBox("get", "["+key.name+"]"); }
    }
    return "";
}

/**
 * 
 * @param {any} msg : Auszugebender Fehler
 * @param {any} reference : optional, betreffende Datei
 *
 * Gibt Error als showInfoBox() aus.
 */
function __err(msg, reference) {
    workspace.showInfoBox("Skripting ERROR", ((reference == null) ? "" : "Betrifft: " + reference + "<br>") + msg);
}

/**
 * 
 * @param {String []} type : BA/SA/MA/DA - Spezifische Konstanten
 * @param {String} name
 * @param {String} publicationNumber
 *
 * Objekt zur Informations-Zwischenspeicherung von für eine neue Publikation
 */
function LitTypeTuple(type, name, publicationNumber) {
    this.name = name;
    this.number = publicationNumber;
    this.short = type['short'];
    this.full = type['full'];
    this.folder = type['folder'];
    this.regFormPath = type['regFormPath'];
    this.isBA = false;
    if (this.short == LIT_TYPE_BA.short) {
        this.isBA = true;
    }
}

/**
 * 
 * @param {String} name
 * @param {String} split
 *
 * Output : Alternativer Konstruktor für LitTypeTuple; 'name' enthählt Typ und Publikationsnummer - getrennt durch 'split'
 */
function LitTypeTupleByName(name, split) {
    // Typ und Nummer holen
    var nameAA = name.split(split);
    var type = nameAA[0];
    var num = nameAA[1];

    // Auswahl der Typspezifischen konstanten
    switch (true) {
        case type == LIT_TYPE_BA['short'] || type == LIT_TYPE_BA['full']:
            return new LitTypeTuple(LIT_TYPE_BA, name, num);
        case type == LIT_TYPE_MA['short'] || type == LIT_TYPE_MA['full']:
            return new LitTypeTuple(LIT_TYPE_MA, name, num);
        case type == LIT_TYPE_SA['short'] || type == LIT_TYPE_SA['full']:
            return new LitTypeTuple(LIT_TYPE_SA, name, num);
        case type == LIT_TYPE_DA['short'] || type == LIT_TYPE_DA['full']:
            return new LitTypeTuple(LIT_TYPE_DA, name, num);
        case type == LIT_TYPE_DISS['short'] || type == LIT_TYPE_DISS['full']:
            return new LitTypeTuple(LIT_TYPE_DISS, name, num);
		case type == LIT_TYPE_DIV_STUD['short'] || type == LIT_TYPE_DIV_STUD['full']:
            return new LitTypeTuple(LIT_TYPE_DIV_STUD, name, num);
    }
    return null;
}

/**
 * 
 * @param {CONST} format
 *
 * Output : nach 'format' formatierter Nutzername;
 */
function getNameFormatted(name,format) {
    // Falls kein Name übergeben wird, dann Username auslesen	
    if(name=="") {
    var name = workspace.getUserName();
    }
    // Wenn Komma vorhanden dann Namensformatierung aus LDAP: Nachname, Vorname
    	if(name.contains(",")) {
    var un = name.split(", ");
    // [Vorname Nachname]
    if (format == FST_NAME_FST)
        return un[1] + " " + un[0];
    // [Vorname]
    if (format == FST_NAME_ONLY)
        return un[1];
    // [Nachname]
    if (format == LST_NAME_ONLY)
        return un[0];
    // [Nachname_Vorname]
    if (format == LST_NAME_FST_NAME)
        return un[0] + "_" + un[1];
    //[Nachname Vorname]
    return un[0] + " " + un[1];
    	}
    // Ansonsten andere Namensformatierung z.B.: Vorname Nachname
	else{
    // Sonderfall abfangen        
    if(name=="Administrator") {
    var vorname = name;
    var nachname = name;
    }   
    else{  
    var pos1 = name.lastIndexOf(" ");
    var pos2 = name.length();
    var vorname = name.substring(0,pos1);
    var nachname = name.substring(pos1+1,pos2);    
    // [Vorname Nachname]
    if (format == FST_NAME_FST)
        return vorname + " " + nachname;
    // [Vorname]
    if (format == FST_NAME_ONLY)
        return vorname;
    // [Nachname]
    if (format == LST_NAME_ONLY)
        return nachname;
    // [Nachname_Vorname]
    if (format == LST_NAME_FST_NAME)
        return nachname + "_" + vorname;
    //[Nachname Vorname]
    return nachname + " " + vorname;
    }
    	}
}

/**
 * 
 * @param {Sord} sord
 * @param {String} path
 *
 * Output : Kopiert alle Elemente im Ordner unter 'path' in den Ordner'sord''
 */
function copyFolder(sord, path) {
    // Source Ordner holen
    var source = archive.getElementByArcpath(path);

    // Zu kopierende Elemente
    var children = source.children;
    while (children.hasMoreElements()) {
        var toCopy = children.nextElement();
        // kopieren
        if (toCopy.parentId == source.id) {
            // Element 
            sord.addCopy(toCopy, true, true, true, false, null);
        } else {
            // Referenz
            ixc.refSord(null, sord.id, toCopy.id, -1);
        }
    }
}

/**
 *
 * @param {ArchiveElement} source
 *
 * Output : Erzeugt eine Kopie von 'source', um das anhängen an eine Mail per COM zu ermöglichen
 */
function attachFile(source) {
    // Datei holen
    var file = source.file;
    var ext = utils.getFileExtension(file);
    // Menschenleserlichen Namen hinzufügen
    var docName = File.separator + source.name + "." + ext;
    // kopieren
    var dest = utils.getUniqueFile(workspace.directories.tempDir, docName)
    utils.copyFile(file, dest);
    // COM formatierter Pfad
    return [dest.path];
}

/**
 * 
 * @param {String} path
 *
 * Output : Analog zu attachFile
 */
function attachFolder(path) {
    // Ordner holen
    var source = archive.getElementByArcpath(path);
    var list = new Array();

    // Iteration über Dateien
    var children = source.children;
    while (children.hasMoreElements()) {
        // Datei holen
        var child = children.nextElement();
        var file = child.file;
        var ext = utils.getFileExtension(file);
        // kopieren mit Menschenleserlichem Namen
        var docName = File.separator + child.name + "." + ext;
        var dest = utils.getUniqueFile(workspace.directories.tempDir, docName)
        utils.copyFile(file, dest);

        // COM formatierten Pfad hinzufügen
        list.push(dest.path);
    }
    return list;
}

/**
 * 
 * @param {String} path
 *
 * Output : Analog zu attachFolder
 */
function attachFolderAndFile(path, filepath) {
    // Ordner holen
    var source = archive.getElementByArcpath(path);
    var list = new Array();

    // Iteration über Dateien im Ordner
    var children = source.children;
    while (children.hasMoreElements()) {
        // Datei holen
        var child = children.nextElement();
        var file = child.file;
        var ext = utils.getFileExtension(file);
        // kopieren mit Menschenleserlichem Namen
        var docName = File.separator + child.name + "." + ext;
        var dest = utils.getUniqueFile(workspace.directories.tempDir, docName)
        utils.copyFile(file, dest);

        // COM formatierten Pfad hinzufügen
        list.push(dest.path);
    }
	
	//Datei hinzufuegen
	// Datei holen
	var source = archive.getElementByArcpath(filepath);
    var file = source.file;
    var ext = utils.getFileExtension(file);
    // Menschenleserlichen Namen hinzufügen
    var docName = File.separator + source.name + "." + ext;
    // kopieren
    var dest = utils.getUniqueFile(workspace.directories.tempDir, docName)
    utils.copyFile(file, dest);
    // COM formatierter Pfad
    list.push(dest.path);
	
    return list;
}

/**
 * 
 * @param {String} to
 * @param {String} subj
 * @param {String} text
 * @param {String []} attachPaths
 *
 * Output : Erzeugt per COM  eine E-Mail mit Betreff 'subj' und Text 'text' an 'to', und hängt 'attachPaths' an. 
 */
function mail(to, subj, text, attachPaths, noDisplay) {
    var sendNoDisplay = noDisplay;
    if (noDisplay == undefined)
        sendNoDisplay = false;
    if (attachPaths == undefined)
        attachPaths = [];
    ComThread.InitSTA();
    try {
        // Outlook öffnen
        var app = new ActiveXComponent("Outlook.Application");
        // Email erzeugen
        var mail = Dispatch.call(app, "CreateItem", 0).toDispatch();
        // Signatur zwischenspeichern
        Dispatch.call(mail, "Display");
        var signature = Dispatch.get(mail, "HTMLBody");
        // Bausteine einfügen
        Dispatch.put(mail, "Subject", subj);
        Dispatch.put(mail, "HTMLBody", text + signature);
        Dispatch.put(mail, "To", to);
        // Anhänge
        var dispAtach = Dispatch.get(mail, "Attachments").toDispatch();
        for (var i = 0; i < attachPaths.length; i++) {
            Dispatch.call(dispAtach, "Add", attachPaths[i]);
        }
        // Fenster anzeigen
        if (sendNoDisplay)
            Dispatch.call(mail, "Send");
        else
            Dispatch.call(mail, "Display");
    } catch (e) {
        workspace.showInfoBox("ERROR", e);
    } finally {
        ComThread.Release();
    }
}

/**
 * 
 * @param {String} day
 * @param {Sting} mmonth
 * @param {String} year
 * @param {boolean} iso
 *
 * Output : Formatiert ein Datum als String (dd[.]mm[.]yyyy); iso = true => keine Trennpunkte => Datum kann in Verschlagwortung eingefügt werden.
 */
function formatDateAsString(day, month, year, iso) {
    var date = "";
    if (day < 10) { date += "0"; }
    if (!iso) { date += day + "."; }

    if (month < 10) { date += "0"; }
    if (!iso) { date += month + "."; }

    if (year < 10) { date += "200"; }
    else if (year < 50) { date += "20"; }
    else if (year < 100) { date += "19"; }
    date += year;

    if (iso) { date += "000000"; }
    return date;
}

/**
 * 
 * @param {Gruppe} entry
 *
 * Output : Gibt den Wert des Feldes 'entry' aus dem Benutzereintrag zurück
 */
function getUserEntry(entry) {
    // Persönlichen Ordner holen
    var uname = workspace.getUserName();
    var item = archive.getElementByArcpathRelative(USERS_FOLDER_ID, "¶" + uname);
    // Index auslesen
    var sord = item.loadSord();
    return getIndexValueByName(sord, entry);
}

/**
 * Output : Gruppen (Feldnamen) der Maske Benutzereintrag
 */
function getUserFieldnames() {
    // Persönlichen Ordner holen
    var uname = workspace.getUserName();
    var item = archive.getElementByArcpathRelative(USERS_FOLDER_ID, "¶" + uname);
    var sord = item.loadSord();

    var ret = [];

    // Feldnamen auslesen
    var objKeys = sord.objKeys;
    for (var i = 0; i < objKeys.length; i++) {
        ret[i] = objKeys[i].name;
    }
    return ret;
}

/**
 * Prüft zugehörigkeit zu den Gruppen, welche Adminrechte im Bezug auf Skriptfunktionen erhalten. 
 */
function isAdmin(aclIn) {
    if (DEBUG_ADMIN_MODE) { return false; }
    var aclAA = [];
    if (aclIn != undefined && aclIn != null)
        aclAA = (aclIn);

    var dbadmin = new AclItem(ACL_FULLACC, 0, GROUP_DBADMIN, AclItemC.TYPE_GROUP);
    aclAA.push(dbadmin);

    var aai = AclAccessInfo();
    aai.setAclItems(aclAA);
    var aar = ixc.getAclAccess(aai);

    return aar.access & ACL_FULLACC != 0;
}



function getHiddenOption(item, option) {
    var info = item.getHiddenText();
    var regex = new RegExp(option + "=[\\S\\s]*\\" + HIDDEN_TEXT_SEPERATOR, "ig");
    var match = info.match(regex);
    if (match == null) return '';
    return match[0].replace(option + "=", "").replace(HIDDEN_TEXT_SEPERATOR, "");
}

function setHiddenOption(item, option, value) {
    clearHiddenOption(item, option);
    var info = item.getHiddenText();
    info += option + "=" + value + HIDDEN_TEXT_SEPERATOR;
    item.setHiddenText(info);
    item.saveSord();
}

function checkInOutClearHiddenText(item) {
    var info = item.getHiddenText();
    var regex = new RegExp(HIDDEN_TEXT_CHECKOUT_DESKTOP + "=[\\S\\s]*\\" + HIDDEN_TEXT_SEPERATOR, "ig");
    var match = info.match(regex);
    if (match == null) return;
    var todelete = match[0];
    info = info.replace(todelete, "");
    item.setHiddenText(info);
    item.saveSord();
}

function clearHiddenOption(item, option) {
    var info = item.getHiddenText();
    var regex = new RegExp(option + "=[\\S\\s]*\\" + HIDDEN_TEXT_SEPERATOR, "ig");
    var match = info.match(regex);
    if (match == null) return;
    var todelete = match[0];
    info = info.replace(todelete, "");
    item.setHiddenText(info);
    item.saveSord();
}

function getUserGroupIDsByName(name) {
    var uZ = CheckoutUsersC.BY_IDS;
    var user = ixc.checkoutUsers([new String(name)], uZ, LockC.NO);
    return user[0].groupList;
}

function getPermissionsBySord(sord) {
    out = [];
    for (var i = 0; i < sord.aclItems.length; i++) {
        var a = sord.aclItems[i];
        out.push([a.type == AclItemC.TYPE_GROUP, a.id, a.access]);
    }
    return out;
}

function checkPermissions(elem, user) {
    var groups = getUserGroupIDsByName(user);
    var permissions = getPermissionsBySord(elem.loadSord());
    var ret = 0;

    for (var i = 0; i < permissions.length; i++) {
        if (!permissions[i][0] && user == permissions[i][1]) {
            ret = ret | permissions[i][2];
            continue;
        }
        for (var j = 0; j < groups.length; j++)
            if (groups[j] == permissions[i][1])
                ret = ret | permissions[i][2];
    }
    return ret;
}

function canReadWrite(elem, userid) {
    var perm = checkPermissions(elem, userid);
    return perm & AccessC.LUR_READ != 0 && perm & AccessC.LUR_WRITE != 0;
}

function folderDepth(item) {
    var id;
    try {
        id = item.getId();
    } catch (e) {
        return 1;
    }
    if (id == 1)
        return 0;
    var parent;
    try{
        parent = item.getParent();
    } catch (e) {
        parent = null;
    }
    if (parent == null)
        return 0;
    return 1 + folderDepth(item.getParent());
}

function dateDiff(d1, d2) {
    var t1 = d1.getTime();
    var t2 = 0;
    if (d2 == undefined)
        t2 = Packages.java.lang.System.currentTimeMillis();
    else
        t2 = d2.getTime();
    return parseInt((t1 - t2) / 60000, 10);
}

function prnt(w) {
    workspace.showInfoBox("", w);
}

function getDescOption(elem, id) {
    var desc = elem.loadSord().desc;
    var pattern = OPTION_START + id + OPTION_TUPEL_SEP + ".*" + OPTION_END;
    var escapedpattern = pattern.replace(new RegExp("\\$", "gi"), "\\$");
    var tuple = new String(desc.match(escapedpattern));
    var opt = tuple.split(OPTION_TUPEL_SEP)[1].replace(OPTION_END, "");
    return opt;
}


function setDescOption(elem, id, val) {
    var sord = elem.loadSord();
    var desc = sord.desc;
    if ((new String(desc)).split(OPTION_START).length == 1)
        desc += "\n\n------------------------------------------------------------";
    sord.desc = desc + "\n" + OPTION_START + id + OPTION_TUPEL_SEP + val + OPTION_END;
    elem.sord = sord;
    elem.saveSord();
    return;
}


function reloadIXScript(id) {
    try {
        checkout.selectId(id);
        checkout.getFirstSelected().checkIn("0", "scripted ix reload");
    } catch (e) { prnt(e); }
    ixc.reload();
    archive.getElement(id).checkOut();
}

function testFunctions() {
    return isAdmin() && BETA_TESTING;
}

function readADBCredentialsFromFile(id) {
    var fis = new FileInputStream(archive.getElement(id).getFile());
    var reader = new BufferedReader(new InputStreamReader(fis));
    ADB_USERNAME = reader.readLine();
    ADB_PASSWORD = reader.readLine();
}


function deleteWindowsFolder(file) {
    if (file.isDirectory()) {
        var windowsChild = file.listFiles();
        for (var i = 0; i < windowsChild.length; i++) {
            deleteWindowsFolder(windowsChild[i]);
        }
    }
    file.delete();
}

function getWindowsExtension(windowsName) {
    var matches = windowsName.match(new RegExp("\\.[\\S\\s]*", "ig"));
    var ext = "";
    if (matches != null)
        ext = matches[0];
    while (ext.includes("."))
        var ext = ext.substring(1, ext.length);
    return ext;
}

function getEloNameFromWindowsFile(file) {
    var windowsName = file.getName();
    var ext = getWindowsExtension(windowsName);
    return windowsName.replace("." + ext, "");
}

function getEloExtension(item) {
    if (item.isStructure()) return "";
    return "." + getWindowsExtension(item.getFile().getName());
}

function copyIndexValueByName(sourcesord, destsord, name){
    setIndexValueByName(destsord, name, getIndexValueByName(sourcesord, name));
}

var SHEET = null, DOC = null, WORKBOOK = null;
function initCOM(filepath) {
    if (DOC != null || SHEET != null || WORKBOOK != null) {
        killCOM();
        workspace.showInfoBox("initCOM", "com already initialized for some file!");
    }
    ComThread.InitSTA();
    try {
        var EXL = new ActiveXComponent("Excel.Application");
        WORKBOOK = Dispatch.get(EXL, "Workbooks").toDispatch();
        DOC = Dispatch.call(WORKBOOK, "Open", filepath).toDispatch();
        SHEET = Dispatch.call(DOC, "Sheets", 1).toDispatch();
    } catch (e) {
        workspace.showInfoBox("Error", e);
        killCOM();
    }
}

function killCOM(nosave) {
    if (!nosave)
        Dispatch.call(DOC, "Save");
    Dispatch.call(DOC, "Close");
    ComThread.Release();
    DOC = null;
    SHEET = null;
    WORKBOOK = null;
}

function colorCell(sheet, row, col, color) {
    var cell = Dispatch.call(sheet, "Cells", row, col).toDispatch();
    var interior = Dispatch.get(cell, "Interior").toDispatch();
    Dispatch.put(interior, "ColorIndex", String(color));
}

function styleCell(sheet, row, col, stylename) {
    var cell = Dispatch.call(sheet, "Cells", row, col).toDispatch();
    var style = Dispatch.call(DOC, "Styles", stylename).toDispatch();
    Dispatch.put(cell, "Style", style);
}

function readCell(sheet, row, col) {
    if (col == null) return "";
    var cell = Dispatch.call(sheet, "Cells", row, col).toDispatch();
    return Dispatch.get(cell, "Text");
}

function setCell(row, col, value) {
    try {
        var cell = Dispatch.call(SHEET, "Cells", row, col).toDispatch();
        Dispatch.put(cell, "Value", String(value));
    } catch (e) {
        workspace.showInfoBox("setCell", e);
        killCOM();
    }
}

function newButton(button, func) {
    ribbon.getButton(button).setCallback(func, this);
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
        bezeichnung.toString() + "_pdf" + 
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
          bezeichnung.toString() + "_pdf" + 
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
 * Erstellt ein ELO-Lesezeichen (ECD-Link) für ein Element und legt es im temp-Ordner ab.
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