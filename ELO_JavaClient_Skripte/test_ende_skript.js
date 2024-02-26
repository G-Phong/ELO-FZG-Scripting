//Dieses Skript testet die Funktion "Ende-Skript" in einem Workflow

function onExitNode(ci, userId, workflow, nodeId) {
  var editInfo = ix.checkoutSord(
    ci,
    workflow.objId,
    CONST.EDIT_INFO.mbSord,
    CONST.LOCK.NO
  );
  var sord = editInfo.sord;
  var keys = sord.objKeys;

  workspace.showInfoBox("Ende-Skript", "Ende-Skript");
}
