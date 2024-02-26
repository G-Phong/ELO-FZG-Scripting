//javascript

/* fuer neuen shortcut:
 * 1. Block kopieren
 * 2. presel (Start bei 0), name und position anpassen
 */

function exportButtons(panel, chkbx) {
    /* 
    //<Outlook Export>
    var name = "Outlook Export"; 
    var position = 3; // x-Achse
    var button = new JButton(name);
    button.addActionListener(function () {
        var presel = [2,3]; // checkbox sind entsprechend der Ansicht durchnummeriert (beginnend bei 0)
        for (var i = 0; i < chkbx.length; i++)
            chkbx[i].setChecked(false);
        for (var i = 0; i < presel.length; i++)
            chkbx[presel[i]].setChecked(true);
    });
    panel.addComponent(position, 1, 1, 1, button);
    // </Outlook Export>
    */


    // Alles/Nichts auswaehlen
    var name = "Auswahl aufheben";
    var position = 1; // x-Achse
    var button = new JButton(name);
    button.addActionListener(function () {
        for (var i = 0; i < chkbx.length; i++)
            chkbx[i].setChecked(false);
    });
    panel.addComponent(position, 1, 1, 1, button);

    var name = "Alles auswählen";
    var position = 2; // x-Achse
    var button = new JButton(name);
    button.addActionListener(function () {
        for (var i = 0; i < chkbx.length; i++)
            chkbx[i].setChecked(true);
    });
    panel.addComponent(position, 1, 1, 1, button);

    var name = "Telefonliste";
    var position = 3; // x-Achse
    var button = new JButton(name);
    var presel3 = [4,5,12,13,15,23,26];
    button.addActionListener(function () {
        for (var i = 0; i < chkbx.length; i++)
            chkbx[i].setChecked(false);
        for (var i = 0; i < presel3.length; i++)
            chkbx[presel3[i]].setChecked(true);
    });
    panel.addComponent(position, 1, 1, 1, button);
	
    var name = "Outlook-Export";
    var position = 4; // x-Achse
    var button = new JButton(name);
    var presel4 = [3,4,5,18,19,20,22,23,25,26,27,33,34,36,38,39];
    button.addActionListener(function () {
        for (var i = 0; i < chkbx.length; i++)
            chkbx[i].setChecked(false);
        for (var i = 0; i < presel4.length; i++)
            chkbx[presel4[i]].setChecked(true);
    });
    panel.addComponent(position, 1, 1, 1, button);
	
	var name = "Briefumschlag";
    var position = 5; // x-Achse
    var button = new JButton(name);
    var presel5 = [1,2,3,4,5,18,19,20,33,34,35,36,37,38,39];
    button.addActionListener(function () {
        for (var i = 0; i < chkbx.length; i++)
            chkbx[i].setChecked(false);
        for (var i = 0; i < presel5.length; i++)
            chkbx[presel5[i]].setChecked(true);
    });
    panel.addComponent(position, 1, 1, 1, button);
}

