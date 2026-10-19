# Officium – Anleitung

Officium fasst die Kalendertermine eines Kunden für einen Zeitraum zusammen:
Termine mit Datum, Uhrzeit, Dauer und Text, dazu die Summe pro Monat oder
Quartal. Den Bericht gibt es als CSV-Datei, als Text zum Einfügen in eine Mail
und, wenn Excel installiert ist, als Excel-Mappe.

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

Outlook führt Makros nur aus, wenn sie signiert sind.

1. Das Programm `SELFCERT.EXE` aus dem Office-Ordner starten, meist
   `C:\Program Files\Microsoft Office\root\Office16\`. Als Namen `Officium`
   eingeben.
2. Im VBA-Editor (Alt+F11) **Extras** > **Digitale Signatur** > **Wählen**,
   das Zertifikat `Officium` auswählen und speichern.
3. In Outlook **Datei** > **Optionen** > **Trust Center** > **Einstellungen
   für das Trust Center** > **Makroeinstellungen** die Option
   „Benachrichtigungen für digital signierte Makros, alle anderen Makros sind
   deaktiviert“ wählen.
4. Outlook neu starten. Beim ersten Start des Makros **Allen Dokumenten von
   diesem Herausgeber vertrauen** wählen.

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
2. **Kunde**: Name einer Outlook-Kategorie oder ein Wort aus dem Betreff.
   Der zuletzt verwendete Kunde ist vorgeschlagen. Bleibt das Feld leer, kommen
   alle Termine in den Bericht.
3. **Gruppierung**: nach Monat, nach Quartal oder keine.

**Abbrechen** beendet Officium jederzeit.

Danach zeigt eine Meldung die Summen. Nach **OK** öffnet sich der Bericht in
Excel, ohne Excel die CSV-Datei.

## Welche Termine zählen

Ein Termin gehört zum Kunden, wenn er dessen Kategorie trägt oder das Wort im
Betreff steht. Groß- und Kleinschreibung ist egal. Ganztägige Termine,
abgesagte und abgelehnte Besprechungen zählen nicht. Ein Termin zählt in dem
Monat, in dem er beginnt.

Gelesen wird nur Ihr eigener Hauptkalender.

## Wo der Bericht landet

- **CSV-Datei** im Ordner Dokumente, zum Beispiel
  `Officium Acme 2026-10-01 bis 2026-10-31.csv`. Sie enthält den vollständigen
  Text jedes Termins. Ein neuer Bericht mit gleichem Kunden und Zeitraum
  ersetzt die Datei. Ist sie noch in Excel geöffnet, wird sie nicht
  gespeichert, das steht dann in der Meldung.
- **Zwischenablage**: Mit Strg+V in eine Mail einfügen. Pro Termin steht dort
  die erste Zeile des Termintexts.
- **Excel**: Blatt „Termine“ mit Filter und Summe, Blatt „Summen“ mit den
  Monats- oder Quartalssummen. Die Mappe ist noch nicht gespeichert.

## Entfernen

Im VBA-Editor (Alt+F11) mit Rechtsklick auf `Officium` und `OfficiumCore`
jeweils **Entfernen** wählen und den Button unter **Menüband anpassen** wieder
herausnehmen.
