# Officium – Anleitung

Officium fasst die Kalendertermine eines Kunden für einen Zeitraum zusammen:
Termine mit Datum, Uhrzeit, Dauer und Text, dazu die Summe pro Monat oder
Quartal. Den Bericht gibt es als CSV-Datei, als Text zum Einfügen in eine Mail
und, wenn Excel installiert ist, als druckfertige Excel-Mappe, die Sie an Ihren
Kunden weitergeben können.

Officium ist ein Makro für das klassische Outlook unter Windows (32 oder
64 Bit). Im neuen Outlook läuft es nicht. Weitere Software braucht es nicht.

## Einrichten

Sie haben zwei Dateien erhalten: `Officium.bas` und `OfficiumCore.bas`.
Öffnen Sie die Dateien nicht in einem Editor, sonst können Umlaute kaputtgehen.

**Module einfügen**

1. In Outlook **Alt+F11** drücken, der VBA-Editor öffnet sich.
2. **Datei** > **Datei importieren** und `Officium.bas` wählen, danach genauso
   `OfficiumCore.bas`.
3. **Datei** > **Speichern** und den VBA-Editor schließen.

**Makro freigeben**

Outlook führt Makros nur aus, wenn sie signiert sind. Dafür legen Sie einmal
ein eigenes Zertifikat an, signieren damit das Makro und sagen Windows, dass
es diesem Zertifikat vertrauen soll. Das Zertifikat gilt nur für Ihren
Windows-Benutzer auf diesem PC; jeder weitere PC braucht ein eigenes.

1. **Zertifikat anlegen.** Windows-Taste drücken und `Digitales Zertifikat`
   tippen. Erscheint **Digitales Zertifikat für VBA-Projekte**, starten.
   Sonst im Explorer das Programm `SELFCERT.EXE` öffnen, je nach Office in
   `C:\Program Files\Microsoft Office\root\Office16\` oder
   `C:\Program Files (x86)\Microsoft Office\root\Office16\`. Als Namen
   `Officium` eingeben und mit **OK** bestätigen. Fehlt das Programm, ist
   die Office-Komponente „Digitale Signatur für VBA-Projekte“ nicht
   installiert; dann hilft Ihre IT.
2. **Makro signieren.** In Outlook Alt+F11 drücken, im VBA-Editor
   **Extras** > **Digitale Signatur** > **Wählen**, `Officium` markieren,
   zweimal **OK**, dann **Datei** > **Speichern**.
3. **Makroeinstellung.** In Outlook **Datei** > **Optionen** > **Trust
   Center** > **Einstellungen für das Trust Center** > **Makroeinstellungen**
   die Option „Benachrichtigungen für digital signierte Makros, alle anderen
   Makros sind deaktiviert“ wählen und mit **OK** schließen.
4. **Zertifikat vertrauen.** Windows-Taste drücken, `certmgr.msc` eingeben
   und Enter drücken. Links **Eigene Zertifikate** > **Zertifikate** öffnen.
   Rechts mit Rechtsklick auf `Officium` **Kopieren** wählen. Links
   **Vertrauenswürdige Stammzertifizierungsstellen** aufklappen und mit
   Rechtsklick auf deren Ordner **Zertifikate** **Einfügen** wählen. Die
   Sicherheitswarnung, dass ein Zertifikat installiert wird, mit **Ja**
   bestätigen. Zur Kontrolle steht `Officium` danach in beiden Ordnern.
5. **Neu signieren.** Im VBA-Editor die Signatur wie in Schritt 2 noch
   einmal wählen und speichern. Erst jetzt gilt das Zertifikat als
   vertrauenswürdig.
6. **Neu starten.** Outlook ganz schließen und neu starten. Beim ersten
   Start des Makros fragt Outlook nach; **Allen Dokumenten von diesem
   Herausgeber vertrauen** wählen. Danach fragt Outlook nicht mehr.

Passiert beim Klick auf den Button nichts, blockiert Outlook das Makro ohne
Meldung. Dann die Schritte 3 bis 6 prüfen.

Lassen sich die Makroeinstellungen nicht ändern, hat Ihre IT Makros gesperrt.
Dann muss die IT Officium freigeben.

**Button anlegen**

1. **Datei** > **Optionen** > **Menüband anpassen**.
2. Rechts **Start** markieren, **Neue Gruppe** anlegen und zum Beispiel
   „Officium“ nennen.
3. Links unter **Befehle auswählen** auf **Makros** umstellen,
   `Project1.OfficiumReport` auswählen und **Hinzufügen** klicken.
4. Über **Umbenennen** einen Namen wie „Kundenbericht“ und ein Symbol wählen,
   mit **OK** schließen.

## Bericht erstellen

Auf den Button klicken. Officium fragt nacheinander:

1. **Zeitraum**: dieser Monat, letzter Monat, dieses Quartal, letztes Quartal
   oder selbst wählen (erster und letzter Tag zählen mit).
2. **Kunde**: ein Name oder Wort aus dem Betreff.
   Der zuletzt verwendete Kunde ist vorgeschlagen. Bleibt das Feld leer, kommen
   alle Termine in den Bericht.
3. **Gruppierung**: nach Monat, nach Quartal oder keine.

**Abbrechen** beendet Officium jederzeit.

Während Officium arbeitet, zeigt ein kleines Fenster in der Bildschirmmitte,
welcher Kalender gerade gelesen wird und woran Excel gerade baut. Danach zeigt
eine Meldung die Summen. Nach **OK** öffnet sich der Bericht in Excel, ohne
Excel die CSV-Datei.

## Welche Termine zählen

Ein Termin gehört zum Kunden, wenn das Wort im Betreff steht. Groß- und
Kleinschreibung ist egal, Kategorien spielen keine Rolle. Ganztägige Termine,
abgesagte und abgelehnte Besprechungen zählen nicht. Ein Termin zählt in dem
Monat, in dem er beginnt.

Gelesen werden alle Kalender, die links in der Kalenderansicht stehen, auch
freigegebene. Fehlt ein freigegebener Kalender, ihn dort einmal anhaken.
Jeder Kalender erscheint mit seinem Postfach, etwa „max@acme.de - Kalender“. Steht
derselbe Termin in mehreren Kalendern, zählt er nur einmal. Die Meldung am Ende
nennt, wie viele Kalender durchsucht wurden und welche sich nicht öffnen
oder lesen ließen.

## Wo der Bericht landet

- **CSV-Datei** im Ordner Dokumente, zum Beispiel
  `Officium Acme 2026-10-01 bis 2026-10-31.csv`. Sie enthält den vollständigen
  Text jedes Termins. Ein neuer Bericht mit gleichem Kunden und Zeitraum
  ersetzt die Datei. Ist sie noch in Excel geöffnet, wird sie nicht
  gespeichert, das steht dann in der Meldung.
- **Zwischenablage**: Mit Strg+V in eine Mail einfügen. Pro Termin steht dort
  die erste Zeile des Termintexts.
- **Excel**: Das Blatt „Bericht“ ist für Ihren Kunden: Kopf mit Kunde und
  Zeitraum, alle Termine mit Text, Zwischensummen und Gesamtsumme, fertig für
  den Druck auf A4 quer. Bei mehreren Kalendern steht bei jedem Termin, aus
  welchem Kalender er stammt. Danach folgt für Sie pro Kalender ein Blatt mit
  Filter. Vor dem Versand diese Blätter löschen oder nur „Bericht“ als PDF
  speichern (**Datei** > **Exportieren**). Die Mappe ist noch nicht
  gespeichert. Erscheint eine Meldung „Excel-Fehler bei …“, öffnet sich
  stattdessen die CSV-Datei; den Text der Meldung bitte an uns weitergeben.

## Neue Version einspielen

Im VBA-Editor (Alt+F11) die Module `Officium` und `OfficiumCore` mit
Rechtsklick **Entfernen** (Export ablehnen), dann die neuen Dateien wie unter
„Module einfügen“ importieren. Danach die Signatur wie unter „Makro freigeben“,
Schritt 2, neu wählen, speichern und Outlook neu starten.

## Entfernen

Im VBA-Editor (Alt+F11) mit Rechtsklick auf `Officium` und `OfficiumCore`
jeweils **Entfernen** wählen und den Button unter **Menüband anpassen** wieder
herausnehmen.
