# Officium

Makro für das klassische Outlook unter Windows, das die Kalendertermine eines
Kunden über einen gewählten Zeitraum zusammenfasst. Vorbild ist der olReporter
von VBOffice.

Officium braucht nichts außer Office. Es gibt keine Webseite, keinen Server und
keinen Installer. Im neuen Outlook läuft es nicht, weil das keine Makros
ausführen kann.

## Stand

Officium ist noch in Arbeit (Plan: `tasks/plan.md`, Aufgaben: `tasks/todo.md`).
Im Moment gibt es zwei Makros. `OfficiumSelfTest` prüft die interne Logik.
`OfficiumListAppointments` fragt nach einem
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

- Klassisches Outlook unter Windows (Microsoft 365 Apps oder Office 2016 und
  neuer), als 32- oder 64-Bit-Version
- Makros dürfen in Outlook nicht per Richtlinie gesperrt sein. Lässt sich die
  Makroeinstellung im Trust Center nicht ändern, hat die IT sie gesperrt.
- Excel ist optional und wird erst für den Excel-Export gebraucht.

## Installieren

### 1. Module importieren

1. Die drei Dateien aus `vba/` herunterladen, auf GitHub jeweils über
   **Download raw file**: `Officium.bas`, `OfficiumCore.bas` und
   `OfficiumTests.bas`. Die Dateien nicht in einem Editor öffnen und neu
   speichern, sonst gehen Zeilenenden oder Umlaute kaputt.
2. Im klassischen Outlook mit **Alt+F11** den VBA-Editor öffnen.
3. Eine ältere Version vorher entfernen: links unter **Module** jeweils
   Rechtsklick auf das Modul > **Entfernen von …**, den Export ablehnen.
4. **Datei** > **Datei importieren**, nacheinander alle drei Dateien. Links
   unter **Module** stehen danach `Officium`, `OfficiumCore` und `OfficiumTests`.
5. **Datei** > **Speichern** (Strg+S).

### 2. Selbsttest

Im VBA-Editor mit **Strg+G** das Direktfenster einblenden, den Cursor in
`OfficiumSelfTest` (Modul `OfficiumTests`) setzen und **F5** drücken. Es
erscheint eine Meldung wie „11 bestanden, 0 fehlgeschlagen“. Schlägt etwas
fehl, steht im Direktfenster, welcher Test und mit welchem Wert.

Läuft gar kein Makro, steht die Makrosicherheit auf „alle deaktivieren“. Dann
erst Schritt 3 erledigen.

### 3. Makro signieren

Outlook führt Makros sicher nur aus, wenn sie signiert sind. Ein eigenes
Zertifikat erzeugt `SELFCERT.EXE`, das bei Office dabei ist.

1. `SELFCERT.EXE` starten. Es liegt im Office-Ordner, meist
   `C:\Program Files\Microsoft Office\root\Office16\`, bei 32-Bit-Office auf
   64-Bit-Windows unter `C:\Program Files (x86)\…`. Als Namen `Officium`
   eingeben.
2. Im VBA-Editor **Extras** > **Digitale Signatur** > **Wählen** und das
   Zertifikat `Officium` auswählen, dann speichern.
3. In Outlook **Datei** > **Optionen** > **Trust Center** > **Einstellungen
   für das Trust Center** > **Makroeinstellungen**: „Benachrichtigungen für
   digital signierte Makros, alle anderen Makros sind deaktiviert“ wählen.
4. Outlook neu starten. Beim ersten Makroaufruf fragt Outlook nach dem
   Herausgeber, dann **Allen Dokumenten von diesem Herausgeber vertrauen**
   wählen.

Ändert man später etwas an den Modulen, signiert Outlook beim Speichern neu,
solange das Zertifikat auf dem Rechner liegt.

Lässt sich die Makroeinstellung nicht ändern, hat die IT sie per Richtlinie
gesperrt. Dann läuft Officium auf diesem Rechner nicht.

### 4. Button in der Menüleiste

1. In Outlook **Datei** > **Optionen** > **Menüband anpassen**.
2. Rechts die Registerkarte **Start** markieren und **Neue Gruppe** anlegen,
   über **Umbenennen** zum Beispiel „Officium“ nennen und ein Symbol wählen.
3. Links unter **Befehle auswählen** auf **Makros** umstellen,
   `Project1.OfficiumListAppointments` auswählen und **Hinzufügen** klicken.
4. Den Befehl über **Umbenennen** zum Beispiel „Termine auflisten“ nennen.
   Mit **OK** schließen.

Alternativ geht es über die Symbolleiste für den Schnellzugriff, dort ebenfalls
unter **Makros**.

## Benutzen

Den Button anklicken, oder mit **Alt+F8** die Makroliste öffnen und
`OfficiumListAppointments` starten. Officium fragt nach dem ersten und dem
letzten Tag, vorbelegt ist die aktuelle Woche. Danach meldet es die Anzahl der
Termine. Die Liste steht im Direktfenster des VBA-Editors (Alt+F11, dann
Strg+G).

## Entfernen

Im VBA-Editor Rechtsklick auf die drei Module > **Entfernen von …**, den Button
unter **Menüband anpassen** wieder herausnehmen. Das Zertifikat lässt sich in
Windows unter **Benutzerzertifikate verwalten** > **Eigene Zertifikate**
löschen.

## Entwicklung

Die Module unter `vba/` sind in Windows-1252 kodiert und haben CRLF als
Zeilenende, weil der VBA-Editor sie so importiert. Git speichert sie byte-genau
(`-text` in `.gitattributes`). Wer sie unter Linux bearbeitet, schreibt UTF-8
und wandelt danach um, zum Beispiel mit `iconv -f utf-8 -t cp1252` und
`unix2dos`.

Ein früherer Prototyp als Web-Add-in für das neue Outlook steckt in der
Git-Historie bis Commit `2160434`.
