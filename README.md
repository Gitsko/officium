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
Die Module unter `vba/` entstehen Schritt für Schritt. Die Anleitung zum
Importieren, Signieren und zum Button in der Menüleiste folgt mit den ersten
Modulen.

Ein früherer Prototyp als Web-Add-in für das neue Outlook steckt in der
Git-Historie bis Commit `2160434`.
