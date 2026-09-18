# Egzamin magisterski — pedagogika przedszkolna i wczesnoszkolna

Materiały do egzaminu **ustnego**: losowane są **2 pytania z 30 tez** (10 ogólnych + 20 kierunkowych).
**Termin egzaminu: środa 23 września.** Plan nauki rozpisany na **18–22 września**.

## Od czego zacząć

| Plik | Co to jest |
|---|---|
| **[plan-nauki-5-dni.md](plan-nauki-5-dni.md)** | **Zacznij tutaj.** Plan na 5 dni z konkretnymi datami, szkielet odpowiedzi ustnej, liczby do wykucia, lista kontrolna |
| **[tezy-maks-skrot.md](tezy-maks-skrot.md)** | **Maks-skrót 30 tez** — tylko najważniejsze informacje: definicja, nazwiska, rdzeń, zdanie-klucz. Do nauki i powtórek |
| [tezy-ogolne-skrot.md](tezy-ogolne-skrot.md) | Pełne wersje 10 tez ogólnych (O1–O10) — do zaglądania, gdy czegoś nie rozumiesz |
| [tezy-kierunkowe-skrot.md](tezy-kierunkowe-skrot.md) | Pełne wersje 20 tez kierunkowych (K1–K20) |

> Pliki `.md` **wyświetlają się bezpośrednio w przeglądarce** — wystarczy kliknąć nazwę powyżej. Nie trzeba nic pobierać.

## Wersje Word (Times New Roman 12, do druku)

Pliki `.docx` są binarne — GitHub ich nie wyświetli, trzeba je **pobrać** (otwórz plik → przycisk *Download*).

| Plik | Zawartość |
|---|---|
| **DO-NAUKI-maks-skrot-i-plan.docx** | **Główny plik do nauki:** plan 5 dni + maks-skrót 30 tez + lista kontrolna |
| Tezy-MAKS-SKROT.docx | Same skrócone tezy (30 stron, teza = strona) |
| Plan-nauki-5-dni.docx | Sam plan z datami |
| KOMPLET-tezy-i-plan.docx | Wszystko razem: plan + maks-skrót + pełne wersje tez |
| Tezy-OGOLNE-skrot.docx | Pełne 10 tez ogólnych |
| Tezy-KIERUNKOWE-skrot.docx | Pełne 20 tez kierunkowych |

Formatowanie: Times New Roman 12, interlinia 1,15, A4, każda teza na osobnej stronie, numeracja stron, wyróżnione zdania-klucze i pytania komisji.

## Plan nauki w skrócie

| Dzień | Data | Czas | Materiał | Tez |
|---|---|---|---|---|
| 1 | piątek 18.09 | 4 h | O1–O7 — podstawy, metodologia, cele i wartości | 7 |
| 2 | sobota 19.09 | 4 h | O8–O10 + K1–K3 | 6 |
| 3 | **niedziela 20.09** | **8 h** | K4–K11 + powtórka wszystkich tez ogólnych | 8 |
| 4 | poniedziałek 21.09 | 4 h | K12–K17 | 6 |
| 5 | wtorek 22.09 | 4 h | K18–K20 + pełna powtórka + symulacja losowania | 3 |

Razem **24 godziny na 30 tez**. Nowy materiał kończy się we wtorek przed południem — resztę dnia zajmuje powtórka.

## Jak regenerować pliki Word

Po edycji plików `.md` wystarczy:

```bash
pip install python-docx
python3 build_docx.py
```

Skrypt [`build_docx.py`](build_docx.py) czyta pliki `.md` z własnego katalogu i nadpisuje wszystkie sześć dokumentów `.docx`.

## Ważne zastrzeżenie

Treść tez powstała na podstawie **standardowego materiału dla kierunku**, nie z notatek z konkretnych wykładów. **Trzeba ją zestawić z własnymi notatkami**, szczególnie:

- **O6** — koncepcje wartości (Chałas, Łobacz) — odsyła do konkretnej książki,
- **O8** — koncepcje wychowania chrześcijańskiego (10 postaci) — wykładowcy różnie dobierają postacie,
- **K20** — kompetencje nauczyciela — bywa własne ujęcie wykładowcy.
