# Officium

Outlook-Add-in für das neue Outlook und Outlook im Web. Es fasst Kalendertermine
pro Kunde und Zeitraum zusammen. Officium ist noch in Arbeit: Die Taskpane zeigt
im Moment nur an, dass das Add-in geladen wurde. Plan: `tasks/plan.md`,
Aufgaben: `tasks/todo.md`.

## Was man zum Benutzen braucht

Auf dem Rechner der Anwender muss nichts installiert werden. Outlook lädt die
Oberfläche als Webseite über HTTPS. Die Taskpane ist allerdings noch nirgends
gehostet und kommt vom Entwicklungsserver unter `https://localhost:3000`. Bis
das Hosting steht, funktioniert Officium deshalb nur auf dem Entwicklerrechner,
und auch dort nur, solange der Server läuft (siehe Entwicklung). Andernfalls
bleibt die Taskpane leer.

## Installation

Installiert wird das App-Paket `officium.zip`. Wer es selbst bauen will, ruft
`bash scripts/package.sh` auf. Das Paket landet dann in `dist/officium.zip`.

Keiner der beiden folgenden Wege ist mit Officium bisher ausprobiert. Beide
entsprechen dem, was Microsoft für Add-ins mit JSON-Manifest (Unified Manifest)
beschreibt.

Über Teams:

1. In Teams **Apps** > **Apps verwalten** > **App hochladen** öffnen.
2. **Benutzerdefinierte App hochladen** wählen und `officium.zip` auswählen.
3. Ein paar Minuten warten, bis Outlook das Add-in anzeigt.

Oder direkt in Outlook: <https://aka.ms/olksideload> öffnen, dann
**Meine Add-Ins** > **Benutzerdefinierte Add-Ins** > **Benutzerdefiniertes
Add-In hinzufügen** > **Aus Datei hinzufügen** und `officium.zip` auswählen.
Dieser Dialog nimmt womöglich nur XML-Manifeste an. Lehnt er das Paket ab,
bleibt der Weg über Teams.

Fehlt die Upload-Option an beiden Stellen, sperrt der Microsoft-365-Tenant
eigene Apps. Dann muss ein Administrator sie freigeben.

## Benutzung

Im neuen Outlook können Add-ins nicht in der reinen Kalenderansicht erscheinen.
Deshalb einen Termin (oder eine Mail) öffnen und in der Menüleiste oder im Menü
**Apps** auf **Officium** klicken. Die Taskpane öffnet sich rechts.

## Deinstallation

In Teams unter **Apps** > **Apps verwalten** Officium suchen und entfernen. Wer
über Outlook installiert hat, entfernt es dort unter **Meine Add-Ins** >
**Benutzerdefinierte Add-Ins**.

## Entwicklung (WSL2 und Outlook unter Windows)

Der Entwicklungsserver läuft in WSL2. Outlook läuft unter Windows und erreicht
ihn über `https://localhost:3000`, weil WSL localhost weiterleitet.

1. Abhängigkeiten installieren: `npm install`
2. Entwicklungszertifikat erzeugen: `npx office-addin-dev-certs install`. Das
   fragt einmal nach `sudo` und trägt die CA in WSL als vertrauenswürdig ein.
3. Dieselbe CA unter Windows vertrauen: `bash scripts/wsl-trust-dev-ca.sh`.
   Sie landet im Speicher CurrentUser Root, Windows fragt zur Bestätigung nach.
4. Entwicklungsserver starten: `npm run dev-server`
5. Unter Windows in Edge `https://localhost:3000/taskpane.html` öffnen. Es darf
   keine Zertifikatswarnung kommen.
6. Sideload: `npm start` meldet sich bei Microsoft 365 an und installiert
   `manifest.json`. Alternativ das App-Paket hochladen wie unter Installation
   beschrieben. `npm stop` beendet und deinstalliert wieder.

## Checks

`npm test`, `npm run typecheck`, `npm run lint`, `npm run validate`, `npm run build`
