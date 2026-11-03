# OpenRouter-Add-in für Excel

Excel-Funktion, die Prompts an OpenRouter schickt:

```
=OPENROUTER.PROMPT(Prompt; [Modell]; [Systemprompt])
```

- Läuft asynchron: Die Zelle zeigt `#BUSY!`, und Excel bleibt bedienbar.
- Begrenzt gleichzeitige Anfragen (Standard: 8) und wiederholt bei Rate-Limits (429) automatisch.
- Antworten werden lokal gecacht, auch über Neustarts hinweg. Fehler werden nicht gecacht.
- Über **Start → OpenRouter** öffnet sich ein Bereich für API-Key, Standardmodell, „Fehler neu abfragen“ und „Cache leeren“.
- Funktioniert in Excel für Windows (Microsoft 365), auf dem Mac und im Web.

## Dateien

| Datei | Zweck |
|---|---|
| `manifest.xml` | Beschreibt das Add-in für Excel. Nur diese Datei wird in Excel bzw. im Admin Center hochgeladen. |
| `taskpane.html` | Einstellungsbereich und Laufzeitumgebung der Funktion |
| `functions.js` | Die eigentliche Funktion |
| `functions.json` | Beschreibung der Funktion (Name, Argumente) |
| `assets/` | Icons |

## 1. Dateien hosten (HTTPS ist Pflicht)

Alle Dateien außer `manifest.xml` müssen per HTTPS erreichbar sein. Sie enthalten keine Geheimnisse, weil der API-Key nur lokal beim Benutzer liegt. Öffentliches Hosting ist also unproblematisch.

**Am einfachsten geht es mit GitHub Pages:**
1. Auf GitHub ein neues Repository `openrouter-addin` anlegen und alle Dateien hochladen.
2. Unter Settings → Pages → „Deploy from a branch“ den Branch `main` und den Ordner `/ (root)` wählen.
3. Nach ein bis zwei Minuten ist das Add-in unter `https://<github-name>.github.io/openrouter-addin/` erreichbar.

Ein eigener Webserver oder IIS mit HTTPS geht genauso.

## 2. Manifest anpassen

In `manifest.xml` alle Vorkommen von

```
https://DEIN-HOST/openrouter-addin
```

durch die echte Adresse ersetzen, ohne Schrägstrich am Ende, z. B. `https://<github-name>.github.io/openrouter-addin`.

## 3. In Excel installieren

**Für das ganze Team (empfohlen):** Microsoft 365 Admin Center → Einstellungen → Integrierte Apps → Benutzerdefinierte Apps hochladen → Typ „Office-Add-in“ → `manifest.xml` hochladen → Benutzer bzw. Gruppen zuweisen. Bis das Add-in bei allen auftaucht, kann es einige Stunden dauern.

**Nur für dich bzw. zum Testen (Netzwerkfreigabe):**
1. `manifest.xml` in einen freigegebenen Ordner legen, z. B. `\\server\addins`.
2. In Excel: Datei → Optionen → Trust Center → Einstellungen für das Trust Center → Vertrauenswürdige Add-In-Kataloge → `\\server\addins` als Katalog-URL eintragen → „Katalog hinzufügen“ → „Im Menü anzeigen“ anhaken.
3. Excel neu starten.
4. Einfügen → Add-Ins → Meine Add-Ins → Freigegebener Ordner → OpenRouter.

**Excel im Web:** Einfügen → Add-Ins → Meine Add-Ins → Mein Add-In hochladen → `manifest.xml`.

## 4. Einrichten

Start → **OpenRouter** → API-Key eintragen → Speichern. Danach z. B.:

```
=OPENROUTER.PROMPT(A2)
=OPENROUTER.PROMPT("Übersetze ins Englische: " & A2)
=OPENROUTER.PROMPT(A2; "anthropic/claude-sonnet-4.5"; "Antworte in max. 10 Wörtern.")
```

## Hinweise

- **Updates:** Änderungen an `functions.js` oder `taskpane.html` einfach auf dem Server austauschen. Werden Funktionen oder Argumente geändert (`functions.json`), in `manifest.xml` die `<Version>` hochzählen und das Manifest neu hochladen.
- **Weitergeben:** Wer die Datei ohne das Add-in öffnet, sieht `#NAME?`. Antworten deshalb vorher als Werte einfügen.
- **Fehler** erscheinen als Text `#FEHLER …` in der Zelle. „Fehler neu abfragen“ im Aufgabenbereich holt nur diese Zellen erneut, alles andere kommt aus dem Cache.
