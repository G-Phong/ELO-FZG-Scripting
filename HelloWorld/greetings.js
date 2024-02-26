const readline = require('readline');

let zahl = 3;

// Erstellen Sie eine Schnittstelle für den Benutzerinput
const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

// Fragen Sie den Benutzer nach seinem Namen
rl.question('Wie ist dein Name? ', (name) => {
  // Geben Sie eine personalisierte Begrüßung aus
  console.log(`Hallo ${name}! Willkommen im Terminal.`);
  console.log(zahl);
  zahl = 4;
  console.log(zahl);
  // Schließen Sie die Benutzerschnittstelle
  rl.close();
});
