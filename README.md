# Finance Bobathon — Budget Controlling mit IBM Bob

Ein ~90-minütiges Hands-on Lab für Finance-Teams: Plan-vs.-Ist-Analyse, Python-Automatisierung und Management-Reporting — alles gesteuert per natürlicher Sprache mit **IBM Bob**.

---

## Szenario

Ihr seid Controller:innen bei der fiktiven **Acme GmbH** und bereitet den Q1 2024-Abschluss vor. Ihr arbeitet mit echten Excel-Dateien, einem Python-Script mit eingebauten Bugs und internen Reporting-Richtlinien — und löst alles gemeinsam mit IBM Bob.

**Keine Programmiererfahrung erforderlich.**

---

## Lab-Struktur

| Teil | Inhalt | Dauer |
|------|--------|-------|
| Teil 1 | Excel-Budget verstehen & Formelfehler reparieren | ~30 min |
| Teil 2 | Python-Script debuggen, erweitern & ausführen | ~35 min |
| Teil 3 | Richtlinien-Check & Executive Summary schreiben | ~25 min |

👉 **[Zur Lab-Anleitung →](LAB.md)**

---

## Dateien

```
data/
  budget_tracker.xlsx        Plan vs. Ist, 5 Kostenstellen, 4 Quarters
                             (enthält einen eingebauten Formelfehler)
  transactions_q1.csv        181 GL-Buchungen Q1 2024
                             (enthält ein Duplikat und ~25% unapproved)
  fx_rates.csv               EUR/USD + EUR/GBP Tageskurse Q1 2024

scripts/
  process_budget.py          Halbfertiges Python-Script für Teil 2
                             (3 Bugs, 2 leere TODO-Funktionen)
  _create_lab_artifacts.py   Instructor-Tool: generiert alle Datendateien neu

docs/
  reporting_guidelines.docx  Fiktive Acme-Reporting-Policy (5 Abschnitte)
```

---

## Setup

**Voraussetzung:** Python 3.10 oder neuer muss installiert sein.
Download: [python.org/downloads](https://www.python.org/downloads/) — beim Installieren auf Windows "Add Python to PATH" aktivieren.

```bash
# 1. Repo klonen
git clone <repo-url>
cd FinanceBobathon

# 2. Virtuelle Umgebung erstellen und aktivieren
python3 -m venv .venv
source .venv/bin/activate      # Windows: .venv\Scripts\activate

# 3. Abhängigkeiten installieren
pip install -r requirements.txt
```

---

## Voraussetzungen

- [IBM Bob](https://www.ibm.com/products/ibm-bob) mit Zugriff auf diesen Workspace
- Python 3.10+ — [python.org/downloads](https://www.python.org/downloads/)
- Keine Programmierkenntnisse erforderlich

---

*Alle Daten und das Unternehmen "Acme GmbH" sind fiktiv. Erstellt für den Finance Bobathon.*
