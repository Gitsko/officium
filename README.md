# Officium

Makro für das klassische Outlook unter Windows. Es fasst die Kalendertermine
eines Kunden über einen gewählten Zeitraum zusammen: Datum, Uhrzeiten, Dauer,
Betreff und Text, mit Summen pro Monat oder Quartal. Der Bericht geht als CSV
raus und als Text in die Zwischenablage, und wenn Excel installiert ist, öffnet
er sich auch dort. Vorbild ist der olReporter von VBOffice.

Officium braucht nichts außer Office. Es gibt keine Webseite, keinen Server und
keinen Installer. Im neuen Outlook läuft es nicht, weil das keine Makros
ausführen kann.

Officium ist noch in Arbeit. Plan: `tasks/plan.md`, Aufgaben: `tasks/todo.md`.
Die Module unter `vba/` entstehen Schritt für Schritt. Wie man das Makro
signiert und als Button in die Menüleiste legt, kommt mit Aufgabe 2 dazu.

## Ausprobieren

1. `vba/Officium.bas` herunterladen. Auf GitHub über **Download raw file**,
   damit Zeilenenden und Kodierung erhalten bleiben. Die Datei nicht in einem
   Editor öffnen und neu speichern.
2. Im klassischen Outlook mit **Alt+F11** den VBA-Editor öffnen.
3. **Datei** > **Datei importieren** und `Officium.bas` auswählen. Links unter
   **Module** erscheint `Officium`. Eine ältere Version vorher entfernen:
   Rechtsklick auf `Officium` > **Entfernen von Officium**, Export ablehnen.
4. Läuft kein Makro, steht die Makrosicherheit vermutlich auf „alle
   deaktivieren“. Zum Testen unter **Datei** > **Optionen** > **Trust Center** >
   **Einstellungen für das Trust Center** > **Makroeinstellungen** die Option
   „Benachrichtigungen für alle Makros“ wählen und Outlook neu starten.
5. Im VBA-Editor mit **Strg+G** das Direktfenster einblenden, den Cursor in
   `OfficiumListAppointments` setzen und **F5** drücken.

Officium fragt nach einem Zeitraum und schreibt alle Termine darin mit Beginn,
Ende, Dauer, Betreff und Kategorien ins Direktfenster.

Ein früherer Prototyp als Web-Add-in für das neue Outlook steckt in der
Git-Historie bis Commit `2160434`.
