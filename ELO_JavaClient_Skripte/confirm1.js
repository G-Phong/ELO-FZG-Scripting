function eloFlowConfirmDialogStart() {
    var dialog = dialogs.flowConfirmDialog;
    var node = dialog.currentNode;
  
    if (node.name == "Eingabe der Basisdaten des Forschungsvorhabens") {
        prnt("confirm1");
    }

    return;
}

//@include lib_utility