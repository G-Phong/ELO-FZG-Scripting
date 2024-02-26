/*
Word Dokument öffnen und im Skript bearbeiten
*/

function processWord(item) {
  let file = item.file;
  ComThread.InitSTA();
  try {
    //Word Instanz öffnen
    let word = new ActiveXComponent("Word.Application");

    //Dokument öffnen
    let documents = Dispatch.get(word, "Documents").toDispatch();
    let doc = Dispatch.call(documents, "Open", file.path).toDispatch();

    //Das dokument kann jetzt mit "doc" genutzt werden

    //Dokument schließen
    let aw = Dispatch.get(doc, "ActiveWindow").toDispatch();
    Dispatch.call(aw, "Close");
  } catch (e) {
    log.info("Error reading word document: " + e);
  }
}
