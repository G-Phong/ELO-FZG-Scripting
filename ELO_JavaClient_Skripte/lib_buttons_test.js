//@include lib_utility
//@include lib_const

function createRibbon2() {
  // Verwaltungen Band
  // Forschungsvorhaben
  button = ribbon.addButton(
    tabFzg.getId(),
    bandVerwaltung.getId(),
    BUTTON_FORSCHUNGSVORHABEN
  );
  button.setOrdinal(45); //TODO: Checken, ob der Wert hier passt
  button.setTitle("Forschungsvorhaben anlegen");
  button.setIconName("ScriptButton27");
  //Callback funktion f�r die Sichtbarkeit des Buttons
  //button.setEnabledCallback(function () {return isLeitungskreis() || isAdmin() || isVerwaltung();}, this); //Zugriff nur f�r Admins, Leitungskreis oder Verwaltung
  //button.setEnabledCallback(function () {return isAdmin();}, this); //Zugriff nur f�r Admins
  button.setEnabledCallback(function () {
    return false;
  }, this); //Zugriff gesperrt (f�r debugging zwecke)
  button.setTooltip("Nur f�r Leitungskreis, Verwaltung oder Admins m�glich");
}
