/*
Word Dokument �ffnen und im Skript bearbeiten
*/

function processWord(item) {
  let file = item.file;
  ComThread.InitSTA();
  try {
    //Word Instanz �ffnen
    let word = new ActiveXComponent("Word.Application");

    //Dokument �ffnen
    let documents = Dispatch.get(word, "Documents").toDispatch();
    let doc = Dispatch.call(documents, "Open", file.path).toDispatch();

    //Das dokument kann jetzt mit "doc" genutzt werden

    //Dokument schlie�en
    let aw = Dispatch.get(doc, "ActiveWindow").toDispatch();
    Dispatch.call(aw, "Close");
  } catch (e) {
    log.info("Error reading word document: " + e);
  }
}
