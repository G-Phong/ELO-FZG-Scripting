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
  //Callback funktion für die Sichtbarkeit des Buttons
  //button.setEnabledCallback(function () {return isLeitungskreis() || isAdmin() || isVerwaltung();}, this); //Zugriff nur für Admins, Leitungskreis oder Verwaltung
  //button.setEnabledCallback(function () {return isAdmin();}, this); //Zugriff nur für Admins
  button.setEnabledCallback(function () {
    return false;
  }, this); //Zugriff gesperrt (für debugging zwecke)
  button.setTooltip("Nur für Leitungskreis, Verwaltung oder Admins möglich");
}
