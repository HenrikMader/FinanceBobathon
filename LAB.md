# Finance Bobathon — Budget Controlling mit IBM Bob

**Dauer:** ~90 Minuten  
**Level:** Einsteiger — keine Programmiererfahrung erforderlich  
**Szenario:** Ihr seid Controller:innen bei der fiktiven **Acme GmbH** und bereitet den Q1-Abschluss vor.

> **Wichtig:** Ihr müsst kein Python verstehen. Ihr beschreibt Bob was ihr wollt — Bob liest, erklärt und schreibt den Code.

---

## Übersicht

| Teil | Thema | Werkzeug | Dauer |
|------|-------|----------|-------|
| [Teil 1](#teil-1) | Excel-Budget verstehen & reparieren | `budget_tracker.xlsx` | ~30 min |
| [Teil 2](#teil-2) | Python-Script debuggen & ausführen *(optional, technisch)* | `process_budget.py` | ~35 min |
| [Teil 3](#teil-3) | Guidelines-Check & Executive Summary | `reporting_guidelines.docx` | ~25 min |

---

## Vorbereitung

### 1. Workspace in IBM Bob öffnen

1. Lade die Lab-Materialien herunter — es gibt zwei Möglichkeiten:

   **Option A — ZIP herunterladen** *(kein Git erforderlich)*
   Gehe auf [github.com/HenrikMader/FinanceBobathon](https://github.com/HenrikMader/FinanceBobathon), klicke auf **Code → Download ZIP** und entpacke das Archiv in einen Ordner auf deinem Rechner (z. B. `FinanceBobathon` auf dem Desktop).

   **Option B — Git Clone**
   Falls Git noch nicht installiert ist, lade es kostenlos herunter:
   - **Windows:** [git-scm.com/download/win](https://git-scm.com/download/win) → Installer ausführen, alle Standardeinstellungen übernehmen
   - **macOS:** Terminal öffnen und `git --version` eingeben — macOS bietet die Installation automatisch an, falls Git fehlt

   Öffne danach ein Terminal, navigiere in einen Ordner (z. B. Desktop) und führe aus:
   ```bash
   git clone https://github.com/HenrikMader/FinanceBobathon.git
   ```
   Der Ordner `FinanceBobathon` wird automatisch erstellt.

2. Öffne IBM Bob und klicke auf **Open Folder** — wähle den soeben erstellten Ordner aus.
3. Bob fragt, ob du dem Ordner vertrauen möchtest: Klicke oben auf **Manage** und dann auf **Trust** (ohne das funktionieren Bobs Werkzeuge nicht).

Stelle sicher, dass folgende Dateien im Bob-Workspace sichtbar sind:

```
data/
  budget_tracker.xlsx        ← Excel: Plan vs. Ist-Daten
  transactions_q1.csv        ← Rohdaten: alle Buchungen Q1 2024
scripts/
  process_budget.py          ← Python-Script (noch nicht fertig!)
docs/
  reporting_guidelines.docx  ← Acme-interne Reporting-Richtlinien
```

---

## Teil 1 — Excel: Budget-Datei verstehen & reparieren {#teil-1}

**Lernziel:** Bob kann Excel-Dateien lesen, Struktur erklären, Fehler finden und direkt korrigieren.

---

### Schritt 1.1 — Datei erklären lassen

> **Prompt 1:**
> ```
> Bitte öffne data/budget_tracker.xlsx und erkläre mir kurz, was diese Datei
> enthält — Struktur, Kostenstellen und was die Spalten bedeuten.
> ```

💡 Bob erkennt die 5 Kostenstellen, die Quarterly-Struktur und die Plan/Ist/Abweichungs-Spalten.

---

### Schritt 1.2 — Fehler finden & korrigieren

> **Prompt 2:**
> ```
> Gibt es in dieser Datei Berechnungen die falsch sein könnten?
> Prüfe besonders die Abweichungs-Prozentspalte — und korrigiere
> den Fehler direkt, wenn du einen findest.
> ```

💡 Die "Abw. %"-Spalte dividiert durch den Ist-Wert statt durch den Plan-Wert. Bob findet und fixt das.

---

### Schritt 1.3 — Auffälligkeiten & offene Punkte

> **Prompt 3:**
> ```
> Welche Kostenstelle hat die auffälligste Abweichung? Gibt es im
> Sheet "Notes & Assumptions" Hinweise dazu? Und was bedeuten
> die TODO-Einträge für unseren Abschluss?
> ```

💡 Bob verbindet die +35%-Überschreitung bei CC200 Marketing mit dem Kommentar im Notes-Sheet und erklärt die offenen Punkte (FX-Anpassung, fehlende IT-Rechnung).

---

### ✅ Checkpoint Teil 1

- Struktur und Inhalt der Excel-Datei verstanden
- Formelfehler gefunden und korrigiert
- Auffälligkeiten und offene TODOs erklärt

---

## Teil 2 — Python: Script debuggen & ausführen *(optional)* {#teil-2}

> 🛠️ **Dieser Teil richtet sich an Teilnehmer:innen mit etwas technischem Hintergrund.**
> Wer kein Terminal nutzen möchte, kann diesen Teil überspringen und direkt mit [Teil 3](#teil-3) weitermachen — der Lab-Abschluss funktioniert auch ohne ihn.

**Lernziel:** Bob kann Code erklären, Bugs finden/fixen und fehlende Logik ergänzen — ohne dass ihr Python kennen müsst.

---

### Schritt 2.0 — Python einrichten

Bevor das Script ausgeführt werden kann, muss Python einmalig eingerichtet werden.

**Python installieren (falls noch nicht vorhanden)**

Prüfe im Terminal:
***python --version***
oder
***python3 --version***

Lade Python 3.10 oder neuer von [python.org/downloads](https://www.python.org/downloads/) herunter und installiere es.

> ⚠️ **Windows:** Beim Installieren unbedingt **„Add Python to PATH"** aktivieren (Checkbox auf dem ersten Installer-Bildschirm). Ohne diese Option findet Windows den `python`-Befehl nicht.

**Abhängigkeiten installieren**

Öffne ein Terminal (macOS: Terminal-App, Windows: Eingabeaufforderung) und führe im Projektordner aus:

```bash
# Virtuelle Umgebung erstellen und aktivieren
python3 -m venv .venv
source .venv/bin/activate      # Windows: .venv\Scripts\activate

# Abhängigkeiten installieren
pip install -r requirements.txt
```

💡 Ihr müsst das nur einmal machen. Danach reicht es, die virtuelle Umgebung zu aktivieren.

---

### Schritt 2.1 — Script erklären lassen

> **Prompt 4:**
> ```
> Bitte öffne scripts/process_budget.py und erkläre mir in einfachen
> Worten was dieses Script tun soll — als wärst du ein Lehrer der
> einem Nicht-Programmierer erklärt was der Code macht.
> ```

---

### Schritt 2.2 — Bugs finden und fixen

> **Prompt 5:**
> ```
> Gibt es Bugs in diesem Script? Liste alle Probleme auf, erkläre was
> jeweils schief gehen würde — und korrigiere dann alle Bugs auf einmal.
> ```

💡 Bob findet drei Bugs: (1) Datum wird nie in ein echtes Datum umgewandelt → Crash, (2) Approved-Filter ist case-sensitiv → filtert alles heraus, (3) CC500 fehlt in den Monatszielen.

---

### Schritt 2.3 — Fehlende Funktionen ergänzen

> **Prompt 6:**
> ```
> Im Script gibt es zwei leere TODO-Funktionen: find_duplicates() und
> find_over_budget(). Bitte implementiere beide — und erkläre mir danach,
> wie ich das Script ausführen kann.
> ```

---

### Schritt 2.4 — Script ausführen & Report lesen

Aktiviere zuerst die virtuelle Umgebung (falls noch nicht aktiv) und führe dann das Script aus:

```bash
# Virtuelle Umgebung aktivieren
source .venv/bin/activate      # Windows: .venv\Scripts\activate

python scripts/process_budget.py
```

> ⚠️ **Windows-Hinweis:** Falls der Befehl `python` nicht gefunden wird, versuche stattdessen `py scripts/process_budget.py`. Sollte das auch nicht funktionieren, war beim Installieren „Add Python to PATH" nicht aktiviert — Python muss in diesem Fall neu installiert werden (siehe Schritt 2.0).

> **Prompt 7:**
> ```
> Das Script hat output/q1_budget_report.xlsx erstellt.
> Bitte öffne die Datei und erkläre mir: Wurden Duplikate oder
> Budget-Überschreitungen gefunden? Was sind die wichtigsten Befunde?
> ```

---

### ✅ Checkpoint Teil 2

- Script-Logik verstanden ohne Python zu können
- Alle Bugs gefunden, erklärt und gefixt
- Report erfolgreich generiert und gelesen

---

## Teil 3 — Guidelines-Check & Executive Summary {#teil-3}

**Lernziel:** Bob kann Daten gegen interne Richtlinien prüfen und fertige Management-Texte schreiben.

---

### Schritt 3.1 — Richtlinien lesen & auf unsere Daten anwenden

> **Prompt 8:**
> ```
> Bitte lies docs/reporting_guidelines.docx und vergleiche die
> Anforderungen mit unseren Q1-Daten aus budget_tracker.xlsx und
> dem generierten Report. Welche Anforderungen erfüllen wir?
> Welche nicht? Bitte als strukturierte Übersicht.
> ```

💡 Bob sollte mindestens finden:
- ✅ Duplikat-Prüfung durchgeführt
- ⚠️ CC200 >10% Abweichung → CFO-Freigabe nötig (Abschnitt 2.1 der Richtlinien)
- ⚠️ FX-Anpassung für CC100 fehlt noch
- ⚠️ Fehlende IT-Rechnung nicht als Accrual ausgewiesen

---

### Schritt 3.2 — Executive Summary für den CFO

> **Prompt 9:**
> ```
> Bitte schreibe eine Executive Summary (max. 150 Wörter, Deutsch) für
> den CFO mit: (1) Gesamtergebnis Q1, (2) den 2 wichtigsten Abweichungen,
> (3) was noch offen ist vor dem finalen Report.
> Ton: professionell und direkt.
> ```

---

### Schritt 3.3 — Freies Erkunden (falls noch Zeit)

Probiert eigene Prompts aus:

- *"Welche Buchungskategorie hat in Q1 am meisten Geld verbraucht?"* (aus `transactions_q1.csv`)
- *"Schreibe eine kurze E-Mail an den Marketing-Leiter: seine Ausgaben lagen 35% über Plan, er braucht eine CFO-Freigabe."*
- *"Was wäre das finanzielle Risiko wenn Marketing in Q2 wieder 35% überzieht?"*

---

## ✅ Lab-Abschluss

In 90 Minuten habt ihr mit IBM Bob Folgendes gemacht — alles per natürlicher Sprache:

| Was | |
|-----|-|
| Excel-Datei strukturell analysiert | ✅ |
| Formelfehler gefunden und gefixt | ✅ |
| Python-Code ohne Programmierkenntnisse verstanden | ✅ |
| 3 Bugs gefunden, erklärt und behoben | ✅ |
| Automatisierten Daten-Report generiert | ✅ |
| Interne Richtlinien gegen echte Daten geprüft | ✅ |
| Executive Summary für den CFO geschrieben | ✅ |

---

*Acme GmbH und alle Daten in diesem Lab sind fiktiv. Erstellt für den Finance Bobathon.*
