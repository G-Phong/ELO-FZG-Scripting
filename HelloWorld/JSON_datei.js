// Beispiel JSON-Daten
const jsonData = '{"name": "John", "age": 30, "city": "New York"}';

// Funktion zum Verarbeiten der JSON-Daten
function processJsonData(data) {
  try {
    // JSON-Daten in ein JavaScript-Objekt parsen
    const obj = JSON.parse(data);

    // Zugriff auf die Eigenschaften des Objekts
    const name = obj.name;
    const age = obj.age;
    const city = obj.city;

    // Verarbeitung der Daten
    const greeting = `Hallo ${name}, du bist ${age} Jahre alt und lebst in ${city}.`;

    // Ausgabe des Ergebnisses
    console.log(greeting);
  } catch (error) {
    console.error('Fehler beim Verarbeiten der JSON-Daten:', error.message);
  }
}

// Aufruf der Funktion mit den Beispiel-Daten
processJsonData(jsonData);
