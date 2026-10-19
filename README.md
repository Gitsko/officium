# Officium

Makro für das klassische Outlook unter Windows, das die Kalendertermine eines
Kunden über einen gewählten Zeitraum zusammenfasst. Vorbild ist der olReporter
von VBOffice.

Officium braucht nichts außer Office. Es gibt keine Webseite, keinen Server und
keinen Installer. Im neuen Outlook läuft es nicht, weil das keine Makros
ausführen kann.

## Was Officium macht

Das Makro `OfficiumReport` fragt nach Zeitraum, Kunde und Gruppierung. Danach
sucht es in allen Kalendern, auch freigegebenen, die passenden Termine,
Serientermine mit jedem einzelnen Vorkommen, und rechnet die Dauer pro Monat
oder Quartal und insgesamt zusammen. Das Ergebnis landet an drei Stellen:

- als CSV-Datei im Dokumente-Ordner, mit allen Terminen und dem vollständigen
  Text jedes Termins,
- als Text in der Zwischenablage, zum Einfügen in eine Mail,
- in einer neuen Excel-Mappe mit einem druckfertigen Bericht für den Kunden,
  falls Excel installiert ist. Ohne Excel öffnet Officium stattdessen die
  CSV-Datei.

## Stand

Version 1.0.0 läuft beim ersten Kunden. Die Änderungen danach (Kunde nur im
Betreff, alle Kalender, Excel-Bericht für den Kunden) sind noch auf keinem
Rechner gelaufen: Der Code entsteht auf einem PC, auf dem nur das neue Outlook
installiert ist. Was beim ersten Durchlauf zu prüfen ist, steht unter
„Erster Test“ weiter unten. Plan und offene Prüfpunkte: `tasks/plan.md` und
`tasks/todo.md`.

## Voraussetzungen

- Klassisches Outlook unter Windows (Microsoft 365 Apps oder Office 2016 und
  neuer), als 32- oder 64-Bit-Version
- Makros dürfen in Outlook nicht per Richtlinie gesperrt sein. Lässt sich die
  Makroeinstellung im Trust Center nicht ändern, hat die IT sie gesperrt.
- Excel ist optional.

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
5. **Debuggen** > **Kompilieren von Project1**. Kommt keine Meldung, ist alles
   in Ordnung. Sonst markiert der Editor die fehlerhafte Zeile.
6. **Datei** > **Speichern** (Strg+S).

### 2. Selbsttest

Im VBA-Editor mit **Strg+G** das Direktfenster einblenden, den Cursor in
`OfficiumSelfTest` (Modul `OfficiumTests`) setzen und **F5** drücken. Erwartet
ist die Meldung „73 bestanden, 0 fehlgeschlagen“. Schlägt etwas fehl, steht im
Direktfenster, welcher Test und mit welchem Wert.

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
4. Windows vertraut einem selbst erzeugten Zertifikat erst, wenn es auch
   unter den Stammzertifizierungsstellen liegt. Dazu `certmgr.msc` starten,
   unter **Eigene Zertifikate** > **Zertifikate** das Zertifikat `Officium`
   mit gedrückter Strg-Taste nach **Vertrauenswürdige
   Stammzertifizierungsstellen** > **Zertifikate** ziehen und die
   Sicherheitswarnung bestätigen. Ohne diesen Schritt meldet Outlook einen
   Fehler beim Überprüfen der Signatur.
5. Im VBA-Editor die Signatur wie in Schritt 2 erneut wählen, speichern und
   Outlook ganz schließen und neu starten. Beim ersten Makroaufruf fragt
   Outlook nach dem Herausgeber, dann **Allen Dokumenten von diesem
   Herausgeber vertrauen** wählen.

Passiert beim Klick auf den Button gar nichts, blockiert Outlook das Makro
ohne Meldung. Dann die Schritte 3 bis 5 prüfen und Outlook neu starten.

Ändert man später etwas an den Modulen, signiert Outlook beim Speichern neu,
solange das Zertifikat auf dem Rechner liegt.

### 4. Button in der Menüleiste

1. In Outlook **Datei** > **Optionen** > **Menüband anpassen**.
2. Rechts die Registerkarte **Start** markieren und **Neue Gruppe** anlegen,
   über **Umbenennen** zum Beispiel „Officium“ nennen und ein Symbol wählen.
3. Links unter **Befehle auswählen** auf **Makros** umstellen,
   `Project1.OfficiumReport` auswählen und **Hinzufügen** klicken.
4. Den Befehl über **Umbenennen** zum Beispiel „Kundenbericht“ nennen. Mit
   **OK** schließen.

Alternativ geht es über die Symbolleiste für den Schnellzugriff, dort ebenfalls
unter **Makros**.

## Benutzen

Den Button anklicken, oder mit **Alt+F8** die Makroliste öffnen und
`OfficiumReport` starten. Officium stellt drei Fragen:

1. **Zeitraum**: 1 dieser Monat, 2 letzter Monat (vorgeschlagen), 3 dieses
   Quartal, 4 letztes Quartal, 5 selbst wählen. Bei 5 folgen der erste und der
   letzte Tag, beide gehören zum Zeitraum.
2. **Kunde**: ein Name oder Stichwort aus dem Betreff.
   Officium schlägt den zuletzt eingegebenen Kunden vor. Ein leeres Feld
   nimmt alle Termine des Zeitraums.
3. **Gruppierung**: 1 Monat, 2 Quartal, 3 keine.

Abbrechen bei einer der Fragen beendet das Makro, ohne etwas zu ändern.

Danach zeigt eine Meldung die Summen und wohin der Bericht gegangen ist. Nach
**OK** öffnet sich Excel, oder ohne Excel die CSV-Datei.

### Welche Termine zählen

Ein Termin gehört zum Kunden, wenn sein Betreff das Stichwort enthält. Groß-
und Kleinschreibung spielt keine Rolle. Bei „Acme“ zählt also der Betreff
„Call mit ACME GmbH“. Kategorien spielen keine Rolle.

Nicht mitgezählt werden ganztägige Termine, abgesagte Besprechungen und
Besprechungen, die man abgelehnt hat. Termine mit Status „frei“ oder
„mit Vorbehalt“ zählen mit. Ein Termin gehört zu dem Zeitraum, in dem er
beginnt. Die Dauer ist Ende minus Beginn.

Officium liest alle Kalender, die in der Kalenderansicht links aufgeführt
sind, auch freigegebene, dazu den Standardkalender jedes Postfachs im Profil.
Ein freigegebener Kalender, der dort nicht steht, fehlt im Bericht; im
Zweifel einmal in der Kalenderansicht anhaken und den Bericht neu erstellen.
Die Meldung am Ende nennt die Zahl der durchsuchten Kalender und die, die
sich nicht lesen ließen. Steht derselbe Termin (gleicher Beginn, gleiches Ende,
gleicher Betreff) in mehreren Kalendern, zählt er nur einmal, und zwar beim
zuerst gelesenen Kalender.

### Die Ausgaben

**CSV**: im Dokumente-Ordner, zum Beispiel
`Officium Acme 2026-10-01 bis 2026-10-31.csv`. Getrennt mit Semikolon, damit
das deutsche Excel die Spalten erkennt. Ein Lauf mit gleichem Kunden und
Zeitraum überschreibt die Datei. Ist sie gerade in Excel geöffnet, meldet
Officium, dass sie nicht gespeichert wurde.

**Zwischenablage**: Titel, Summen und eine Zeile pro Termin mit Datum, Uhrzeit,
Dauer, Betreff und der ersten Zeile des Termintexts.

**Excel**: Das erste Blatt „Bericht“ ist für den Kunden gedacht: Kopf mit
Kunde, Zeitraum und Erstellungsdatum, die Termine nach Datum mit dem
vollständigen Termintext, eine Zwischensumme pro Monat oder Quartal und die
Gesamtsumme. Es ist für A4 quer eingerichtet, die Kopfzeile wiederholt sich
auf jeder Seite. Danach folgt pro Kalender ein internes Blatt mit Autofilter
und einer Summe, die sich nach dem Filter richtet. Vor dem Versand die
internen Blätter löschen oder nur das Blatt „Bericht“ als PDF speichern. Die
Mappe ist nicht gespeichert.

Die CSV-Datei hat eine Spalte „Kalender“.

### Termine prüfen

Das Makro `OfficiumListAppointments` schreibt alle Termine eines Zeitraums
ohne Filter ins Direktfenster des VBA-Editors (Alt+F11, dann Strg+G),
getrennt nach Kalender. Damit lässt sich nachvollziehen, warum ein Termin im
Bericht fehlt oder welcher Kalender nicht lesbar ist.

## Erster Test

Weil der Code noch nie gelaufen ist, beim ersten Mal diese Punkte prüfen:

1. Kompilieren ohne Fehler, Selbsttest mit 73 bestanden.
2. `OfficiumListAppointments` für eine Woche mit einem Serientermin: Jedes
   Vorkommen erscheint einzeln, die Anzahl passt zur Kalenderansicht.
3. `OfficiumReport` für einen Kunden und den letzten Monat: Die Summe stimmt
   mit einer Handzählung überein.
4. Die CSV-Datei öffnet sich in Excel mit richtigen Spalten und Umlauten.
5. Der Text aus der Zwischenablage zeigt in einer Mail die Umlaute richtig.
6. In Excel steht die Dauer als Stunden:Minuten und lässt sich mit `=SUMME()`
   addieren.
7. Mit einem freigegebenen Kalender: Er erscheint in der Meldung, hat ein
   eigenes Blatt, und ein Meeting, das in beiden Kalendern steht, zählt im
   Blatt „Bericht“ nur einmal.
8. Die Druckvorschau des Blatts „Bericht“ zeigt A4 quer, alle Spalten auf der
   Seitenbreite und die Zwischensummen stimmen.

Fehlermeldungen samt der markierten Codezeile helfen bei der Korrektur.

## Entfernen

Im VBA-Editor Rechtsklick auf die drei Module > **Entfernen von …**, den Button
unter **Menüband anpassen** wieder herausnehmen. Das Zertifikat lässt sich in
Windows unter **Benutzerzertifikate verwalten** > **Eigene Zertifikate**
löschen. Den zuletzt eingegebenen Kunden speichert Officium in der Registry
unter `HKEY_CURRENT_USER\Software\VB and VBA Program Settings\Officium`, der
Schlüssel kann gelöscht werden. Die CSV-Dateien im Dokumente-Ordner bleiben
liegen, bis man sie löscht.

## Entwicklung

`OfficiumCore` enthält die Logik ohne Outlook-Objekte, `OfficiumTests` deckt
sie mit dem Selbsttest ab. `Officium` enthält die Makros, die Dialoge, den
Kalenderzugriff und die Ausgaben. Neue Tests müssen in `OfficiumSelfTest`
aufgerufen werden, die erwartete Anzahl im Abschnitt „Selbsttest“ dann
nachziehen.

Die Module unter `vba/` sind in Windows-1252 kodiert und haben CRLF als
Zeilenende, weil der VBA-Editor sie so importiert. Git speichert sie byte-genau
(`-text` in `.gitattributes`). Wer sie unter Linux bearbeitet, schreibt UTF-8
und wandelt danach um, zum Beispiel mit `iconv -f utf-8 -t cp1252` und
`unix2dos`.

Ein früherer Prototyp als Web-Add-in für das neue Outlook steckt in der
Git-Historie bis Commit `2160434`.
