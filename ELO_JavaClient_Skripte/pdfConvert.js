/** 
 * Indexserver calls this function if a workflov node is activated
 * @param ci Clientlnfo object witb language , country and ticket
 * @param uaerid The calling users ID (Integer)
 * @param workflow WFDiagram object
 * @param nodeid The activated node ID (Integer)
 * @return Cycle nodes : true , to repeat the cycle
 */

function onEnterNode(ci, userId, workflow, nodeId) {
    workspace.showInfoBox(
        "Info",
        "Dieses Skript wird jetzt VOR Eintritt in den n�chsten Workflow-Knoten ausgef�hrt! Das geht mit onEnterNode()."
      );
}

/**
 * Indexserver calls this function if a workflow noda is " deactivated
 * @param ci Cliantinfo object with lanquage, country and ticket
 * @param uaarid The calling users ID (Integer)
 * @param workflow WFDiagram object
 * @param nodeid The activated node ID (Integer)
 */
function onExitNode(ci, userId, workflow, nodeId) {
    workspace.showInfoBox(
        "Info",
        "Dieses Skript wird jetzt NACH Eintritt des letzten Workflow-Knotens ausgef�hrt! Das geht mit onExitNode()."
      );
}
