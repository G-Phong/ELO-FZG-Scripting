/*
 * Dieses Skript zeigt, wie Sie Outlook über die COM-Schnittstelle aus einem Java Client steuern können. S.129 Thiele-Skriptprogrammierung
 */

// Jacob Bibliothek = Java-COM Schnittstelle (Component Object Model)
// Ansteuerung von Outlook über COM-Schnittstelle

/*
Der Aufrufvon COM-Befehlen erfolgt über die statische Methode "Dispatch.call()".
 Als erster Parameter wird das Objekt übergeben, aus dem die Methode aufgerufen werden soll.
  Im Beispiel finden Sie das app-Objekt, das die Outlook-Applikation enthält.
  Der zweite Parameter enthält den Namen der Methode als Text.
  Danach folgen die Parameter der Methode, hier 0.
  Da der "Dispatch.caii"-Aufruf als Ergebnis eine Variante zurück liefert,
   muss diese bei Bedarfwieder in ein Dispatch-Objekt umgewandelt werden.
   Das passiert durch das abschließende "toDispatch()".
*/

//Dieses Skript sollte durch einen Button aufgerufen werden. Davor sollte die zu versendene Datei im ELO Archiv markiert sein.

importPackage(Packages.de.elo.client);
importPackage(Packages.com.jacob.com);
importPackage(Packages.com.jacob.activeX);
importPackage(Packages.com.ms.activeX);
importPackage(Packages.com.ms.com);

// TODO: Button ins FZG Ribbon setzen!!
function getScriptButton1Name() {
  return "Hello";
}

function getScriptButtonPositions() {
  return "1,home,view";
}

function getScriptButton1Start() {
  exportToOutlook(workspace.activeView.firstSelected);
}

function exportToOutlook(item) {
  ComThread.InitSTA();
  try {
    let app = new ActiveXComponent("Outlook.Application");
    let mail = Dispatch.call(app, "CreateItem", 0).toDispatch();
    Dispatch.put(mail, "Subject", " Dokument: " + item.name);
    let sord = item.loadSord();
    Dispatch.put(mail, "Body", sord.desc);
    let attachments = Dispatch.get(mail, "Attachments").toDispatch();
    Dispatch.call(attachments, "Add", item.file.path);
  } catch (e) {
    log.info("Error: " + e);
  } finally {
    ComThread.Release();
  }
}
