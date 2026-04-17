# QR Ticket Email Sender

Automatický systém na rozosielanie vstupeniek e-mailom. Načíta zoznam študentov z Excelu, vygeneruje personalizovaný lístok zo SVG šablóny s QR kódom, priloží pozvánku do kalendára a odošle všetko e-mailom cez [Resend](https://resend.com).

---

## Obsah

- [Požiadavky](#požiadavky)
- [Inštalácia](#inštalácia)
- [Konfigurácia .env](#konfigurácia-env)
- [Príprava SVG šablóny](#príprava-svg-šablóny)
- [Štruktúra Excelu](#štruktúra-excelu)
- [Použitie](#použitie)
- [Nastavenie PDF offsetu textu](#nastavenie-pdf-offsetu-textu)
- [Súbory projektu](#súbory-projektu)

---

## Požiadavky

- Python 3.9+
- Účet na [Resend](https://resend.com) (bezplatný plán stačí na testovanie)
- Overená doména v Resende (pre reálne odosielanie)

---

## Inštalácia

```bash
pip install resend qrcode[pil] openpyxl pillow python-dotenv playwright
playwright install chromium
```

> `playwright install chromium` stiahne ~150 MB Chromium prehliadač. Stačí urobiť raz.

---

## Konfigurácia .env

Vytvor súbor `.env` v rovnakom priečinku ako skripty:

```env
RESEND_API_KEY=re_xxxxxxxxxxxxxxxxxxxx
SENDER_EMAIL=listky@tvojadomena.sk
SENDER_NAME=Organizátor
EVENT_NAME=Rozlúčka 2026
EVENT_DATE=15.6.2026
EVENT_TIME=18:00
EVENT_LOCATION=Aula, Gymnázium Košice
EVENT_DURATION_MINUTES=120
```

**Formáty dátumu** akceptované pre kalendárovú pozvánku: `DD.MM.YYYY`, `YYYY-MM-DD`, `DD/MM/YYYY`.

`EVENT_DURATION_MINUTES` je voliteľné (predvolené: 120). Určuje dĺžku udalosti v kalendári.

---

## Príprava SVG šablóny

Nástroj `prepare_svg.py` upraví tvoj Figma export tak, aby fungoval so systémom.

```bash
python prepare_svg.py tvoj_listok.svg
```

Skript sa ťa spýta:
1. Ktorý textový element je **meno študenta** → nahradí ho zástupným textom `{NAME_PLACEHOLDER}`
2. Ktorý je **trieda** → nahradí ho `{CLASS_PLACEHOLDER}`
3. Či chceš pridať **QR sidebar** vpravo (odporúčané)

Po každom zástupnom texte sa môžeš opýtať na **X offset** (posun vpravo/vľavo v pixeloch), ak text nesedí presne.

Výsledok sa uloží ako `ticket_template.svg`.

> **Dôležité:** Export z Figmy musí obsahovať skutočné `<text>` elementy. Ak si v Figme použil „Outline text", text bude skonvertovaný na krivky a skript ho nenájde. Exportuj znova bez tejto možnosti.

---

## Štruktúra Excelu

Súbor `students_database.xlsx` musí mať tieto stĺpce:

| A — Meno | B — Trieda | C — ID | D — Email | E — Odoslané |
|----------|------------|--------|-----------|--------------|
| Ján Novák | 4.A | 001 | jan@skola.sk | 0 |

- Stĺpec **E** skript spravuje sám — po úspešnom odoslaní nastaví hodnotu `1`
- Riadky s `1` v stĺpci E sa pri ďalšom spustení preskočia

---

## Použitie

### 1. Preview — skontroluj výsledok bez odosielania

V `main_svg.py` nechaj `MODE = "preview"` (predvolené) a spusti:

```bash
python main_svg.py
```

Vygeneruje sa priečinok `ticket_preview/` so súbormi pre každého študenta:
- `ticket_Meno.svg` — vygenerovaná vstupenka
- `ticket_Meno.pdf` — výsledný PDF (ten sa bude posielať)
- `ticket_Meno.ics` — kalendárová pozvánka

### 2. Nastavenie PDF offsetu textu

Ak je text v PDF posunutý oproti SVG, skalibruj offset:

```bash
python main_svg.py adjustpdf
```

Otvorí sa `ticket_preview/_adjust_test.pdf` (len prvý študent). Podľa výsledku uprav v `main_svg.py`:

```python
TEXT_Y_OFFSET = -10  # záporné = hore, kladné = dole
```

Opakuj, kým text nesedí.

### 3. Reálne odosielanie

Keď si spokojný s výsledkom, zmeň v `main_svg.py`:

```python
MODE = "real"
```

A spusti:

```bash
python main_svg.py
```

Systém sa opýta na potvrdenie, potom odošle e-maily všetkým, ktorí ešte nemajú `1` v stĺpci E.

### 4. Opätovné odoslanie konkrétnemu študentovi

```bash
python main_svg.py resend 001
python main_svg.py resend "Ján Novák"
python main_svg.py resend jan@skola.sk
```

Funguje podľa ID, mena alebo e-mailu (bez rozlišovania veľkých/malých písmen). Ak je `MODE = "real"`, opýta sa na potvrdenie.

---

## Čo dostane každý študent

Každý e-mail obsahuje:
- **HTML telo** s informáciami o podujatí
- **PDF prílohu** — personalizovaná vstupenka s QR kódom
- **ICS prílohu** — pozvánka do kalendára (Google Calendar, Apple Calendar, Outlook)

QR kód obsahuje: `Meno | Trieda | ID`

---

## Nastavenie PDF offsetu textu

Playwright renderuje SVG cez skutočný Chromium prehliadač, čo zaručuje vernú reprodukciu Figma dizajnu. Niekedy však tlačový renderer posúva text mierne inak ako prehliadač. Parameter `TEXT_Y_OFFSET` v `main_svg.py` to koriguje priamo na úrovni SVG súradníc — bez CSS transformácií, ktoré by mohli kolidovať s Figma layoutom.

```python
TEXT_Y_OFFSET = 0    # žiadna korekcia
TEXT_Y_OFFSET = -8   # posun textu o 8px nahor
```

---

## Súbory projektu

```
├── main_svg.py              # hlavný skript — generovanie a odosielanie
├── prepare_svg.py           # príprava SVG šablóny z Figma exportu
├── ticket_template.svg      # pripravená šablóna (generuje prepare_svg.py)
├── students_database.xlsx   # zoznam študentov
├── .env                     # API kľúče a info o podujatí (neverziovať!)
└── ticket_preview/          # výstupy preview módu
    ├── ticket_Meno.svg
    ├── ticket_Meno.pdf
    ├── ticket_Meno.ics
    └── _adjust_test.pdf     # test pre kalibráciu TEXT_Y_OFFSET
```

> **Nikdy neverziovaj `.env`** — obsahuje API kľúče. Pridaj ho do `.gitignore`.
