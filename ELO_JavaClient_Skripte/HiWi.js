function eloExpandRibbon() {
    createRibbon();
    newButton(BUTTON_NEW_HIWI, mailHiWi);
}

// //IndexDialog anzeigen
//     if (!indexDialog.editSord(sord, true, LITERATURE_SORD_TITLE)) {
//         return;
//     }
//     var type = indexDialog.getObjKeyValue(LIT_MASK_TYPE);
//     if (type == '') {
//         return;
//     }
// 
// // Start des Workflows 
//     var wf = ixc.checkoutWorkFlow("SAM", WFTypeC.TEMPLATE, WFDiagramC.mbAll, LockC.NO);
//     wf.type = WFTypeC.ACTIVE;
//     //workspace.showInfoBox("SAM", BASAMA_BENOT_INFO1);

function mailHiWi() {

    // Dialogfenster erstellen
    var dialog = workspace.createGridDialog("HiWi-Einstellung", 2, 2);

    // Auswahldialog (Art der Einstellung) einfuegen
    var choice1 = Choice();
    choice1.add(HIWI_CHOICE1[0]);   // 0 (Einstellung deutsch)
    choice1.add(HIWI_CHOICE1[1]);   // 1 (Einstellung englisch)
    choice1.add(HIWI_CHOICE1[2]);   // 2 (Verlaengerung ohne Pause deutsch)
    choice1.add(HIWI_CHOICE1[3]);   // 3 (Verlaengerung ohne Pause englisch)
    choice1.add(HIWI_CHOICE1[4]);   // 4 (Verlaengerung mit Pause (< 6 Monate) deutsch)
    choice1.add(HIWI_CHOICE1[5]);   // 5 (Verlaengerung mit Pause (< 6 Monate) englisch)
    dialog.addComponent(2, 1, 1, 1, choice1);
	
	// Auswahldialog (Art des Abschlusses) einfuegen
	var choice2 = Choice();
	choice2.add(HIWI_CHOICE2[0]);   // 0 (kein Abschluss)
	choice2.add(HIWI_CHOICE2[1]);   // 1 (Bachelorabschluss)
	//choice2.add(HIWI_CHOICE2[2]);   // 2 (Masterabschluss)
	dialog.addComponent(2, 2, 1, 1, choice2);

    // Beschriftung einfuegen
    dialog.addComponent(1, 1, 1, 1, JLabel(HIWI_LABEL[0]));
	dialog.addComponent(1, 2, 1, 1, JLabel(HIWI_LABEL[1]));

    // Dialog anzeigen
    var ok = dialog.show();

    // Oeffnen der E-Mail, wenn Dialog mit ok beendet wird
    if (ok) {

        var auswahl1 = choice1.getSelectedIndex();
		var auswahl2 = choice2.getSelectedIndex();
		
		if (auswahl2 == 0) {
			DIR_TEMPLATES_MAIL_HIWI_ABRECH = DIR_TEMPLATES_MAIL_HIWI_ABRECH_STUD;
		} else if (auswahl2 == 1) {
			DIR_TEMPLATES_MAIL_HIWI_ABRECH = DIR_TEMPLATES_MAIL_HIWI_ABRECH_BA;
		} else if (auswahl2 == 2) {
			DIR_TEMPLATES_MAIL_HIWI_ABRECH = DIR_TEMPLATES_MAIL_HIWI_ABRECH_MA;
		} else {
			return null;
		}
		
        if (auswahl1 == 0) {
            mail("",
				SUBJ_MAIL_HIWI_DEU,
				FST_MAIL_HIWI_DEU.replace(REPLACE_HIWI_NAME_FST, ""),
				attachFolderAndFile(DIR_TEMPLATES_MAIL_HIWI_DEU,DIR_TEMPLATES_MAIL_HIWI_ABRECH));
        } else if (auswahl1 == 1) {
            mail("",
				SUBJ_MAIL_HIWI_ENG,
				FST_MAIL_HIWI_ENG.replace(REPLACE_HIWI_NAME_FST, ""),
				attachFolderAndFile(DIR_TEMPLATES_MAIL_HIWI_ENG,DIR_TEMPLATES_MAIL_HIWI_ABRECH));
        } else if (auswahl1 == 2) {
            mail("",
				SUBJ_MAIL_HIWI_VERL_DEU,
				FST_MAIL_HIWI_VERL_DEU.replace(REPLACE_HIWI_NAME_FST, ""),
				attachFolderAndFile(DIR_TEMPLATES_MAIL_HIWI_VERL_DEU,DIR_TEMPLATES_MAIL_HIWI_ABRECH));
        } else if (auswahl1 == 3) {
            mail("",
				SUBJ_MAIL_HIWI_VERL_ENG,
				FST_MAIL_HIWI_VERL_ENG.replace(REPLACE_HIWI_NAME_FST, ""),
				attachFolderAndFile(DIR_TEMPLATES_MAIL_HIWI_VERL_ENG,DIR_TEMPLATES_MAIL_HIWI_ABRECH));
        } else if (auswahl1 == 4) {
            mail("",
				SUBJ_MAIL_HIWI_PAUS_DEU,
				FST_MAIL_HIWI_PAUS_DEU.replace(REPLACE_HIWI_NAME_FST, ""),
				attachFolderAndFile(DIR_TEMPLATES_MAIL_HIWI_PAUS_DEU,DIR_TEMPLATES_MAIL_HIWI_ABRECH));
        } else if (auswahl1 == 5) {
            mail("",
				SUBJ_MAIL_HIWI_PAUS_ENG,
				FST_MAIL_HIWI_PAUS_ENG.replace(REPLACE_HIWI_NAME_FST, ""),
				attachFolderAndFile(DIR_TEMPLATES_MAIL_HIWI_PAUS_ENG,DIR_TEMPLATES_MAIL_HIWI_ABRECH));
        } else {
            return null;
        }
    }

    // Dialog wird abgebrochen
    return null;
}

//@include lib_utility
//@include lib_const
//@include lib_buttons

///////////////////////////////////////////////////////////////////////////////////////////////////////////////