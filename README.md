# Officium

Makro für das klassische Outlook unter Windows, das die Kalendertermine eines
Kunden über einen gewählten Zeitraum zusammenfasst. Vorbild ist der olReporter
von VBOffice.

Officium braucht nichts außer Office. Es gibt keine Webseite, keinen Server und
keinen Installer. Im neuen Outlook läuft es nicht, weil das keine Makros
ausführen kann.

## Stand

Officium ist noch in Arbeit (Plan: `tasks/plan.md`, Aufgaben: `tasks/todo.md`).
Im Moment gibt es ein Makro: `OfficiumListAppointments`. Es fragt nach einem
Zeitraum und schreibt alle Termine darin mit Beginn, Ende, Dauer, Betreff und
Kategorien ins Direktfenster des VBA-Editors. Serientermine erscheinen einzeln,
mit jedem Vorkommen.

Geplant sind außerdem:

- Filter auf einen Kunden, über die Outlook-Kategorie oder ein Stichwort im Betreff
- Summen pro Monat oder Quartal und insgesamt
- Export als CSV und als Text in die Zwischenablage
- Bericht direkt in Excel öffnen, sofern Excel installiert ist
- Start über einen eigenen Button in der Menüleiste

## Voraussetzungen

- Klassisches Outlook unter Windows (Microsoft 365 Apps oder Office 2016 und neuer)
- Makros dürfen in Outlook nicht per Richtlinie gesperrt sein. Lässt sich die
  Makroeinstellung im Trust Center nicht ändern, hat die IT sie gesperrt.
- Excel ist optional und wird erst für den Excel-Export gebraucht.

## Installieren

1. `vba/Officium.bas` herunterladen, auf GitHub über **Download raw file**.
   Die Datei nicht in einem Editor öffnen und neu speichern. Sonst gehen
   Zeilenenden oder Umlaute kaputt.
2. Im klassischen Outlook mit **Alt+F11** den VBA-Editor öffnen.
3. Eine ältere Version vorher entfernen: links unter **Module** Rechtsklick auf
   `Officium` > **Entfernen von Officium**, den Export ablehnen.
4. **Datei** > **Datei importieren** und `Officium.bas` auswählen. Links unter
   **Module** erscheint `Officium`.
5. Läuft das Makro nicht, steht die Makrosicherheit vermutlich auf „alle
   deaktivieren“. Unter **Datei** > **Optionen** > **Trust Center** >
   **Einstellungen für das Trust Center** > **Makroeinstellungen** die Option
   „Benachrichtigungen für alle Makros“ wählen und Outlook neu starten. Das
   ist eine Lösung für den Übergang. Mit Aufgabe 2 wird das Makro signiert,
   dann reicht „nur digital signierte Makros“.

## Benutzen

Im VBA-Editor mit **Strg+G** das Direktfenster einblenden, den Cursor in
`OfficiumListAppointments` setzen und **F5** drücken. Officium fragt nach dem
ersten und dem letzten Tag, vorbelegt ist die aktuelle Woche. Danach meldet es
die Anzahl der Termine, die Liste steht im Direktfenster.

## Entfernen

Im VBA-Editor Rechtsklick auf `Officium` > **Entfernen von Officium**. Wurde
die Makrosicherheit für den Test gelockert, sie wieder zurückstellen.

## Entwicklung

Die Module unter `vba/` sind in Windows-1252 kodiert und haben CRLF als
Zeilenende, weil der VBA-Editor sie so importiert. Git speichert sie byte-genau
(`-text` in `.gitattributes`). Wer sie unter Linux bearbeitet, schreibt UTF-8
und wandelt danach um, zum Beispiel mit `iconv -f utf-8 -t cp1252` und
`unix2dos`.

Ein früherer Prototyp als Web-Add-in für das neue Outlook steckt in der
Git-Historie bis Commit `2160434`.
