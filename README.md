# Kalendarz zadań – wtyczka OneNote (przeglądarka)

Panel boczny w OneNote dla przeglądarki: kalendarz miesięczny, zadania na każdy dzień,
przypomnienie o wybranej godzinie. Zadania zapisują się w **Microsoft To Do** (lista
„Kalendarz OneNote”), więc widać je też w Outlooku, To Do i Teams (aplikacja Planner).
Maile i wiadomości w Teams wysyła **Power Automate**, niezależnie od tego, czy OneNote jest otwarty.

```
OneNote (przeglądarka) ──► panel wtyczki (GitHub Pages) ──Graph──► Microsoft To Do
                                                                       │ co 15 min
                                                     Power Automate ◄──┘
                                                       ├─► mail (Outlook)
                                                       └─► wiadomość w Teams (Flow bot)
```

## Pliki

| Plik | Rola |
|---|---|
| `manifest.xml` | Manifest wtyczki – wgrywasz go do OneNote |
| `docs/taskpane.*` | Panel z kalendarzem |
| `docs/config.js` | **clientId aplikacji** + ustawienia |
| `docs/redirect.html`, `docs/dialog.html` | Logowanie (popup MSAL / zapasowo okno Office) |
| `docs/lib/msal-browser.min.js` | Biblioteka logowania Microsoft (v3.30) |

Podgląd bez logowania: `docs/taskpane.html?demo=1` (dane tylko w przeglądarce).

## Krok 1. Hosting (GitHub Pages)

Wtyczka musi być pod adresem HTTPS. Repozytorium `onenote-kalendarz` na koncie `Indy97`,
Pages z gałęzi `main`, folder `/docs` → `https://indy97.github.io/onenote-kalendarz/`.
W kodzie nie ma żadnych sekretów (clientId i tenantId nie są tajne).

## Krok 2. Rejestracja aplikacji w Entra ID (5 minut, bez admina)

> **Zrobione 2026-10-02** przez Azure CLI: aplikacja `Kalendarz OneNote`,
> Client ID `80b6218c-8514-436e-bce7-3f3affb99808` (już wpisany w `config.js`).
> Poniższe kroki zostają na wypadek ponownej rejestracji.

1. https://entra.microsoft.com → **Aplikacje → Rejestracje aplikacji → Nowa rejestracja**
2. Nazwa: `Kalendarz OneNote`, typ kont: **Tylko to katalog organizacyjny**
3. Identyfikator URI przekierowania: platforma **Aplikacja jednostronicowa (SPA)**,
   adres `https://indy97.github.io/onenote-kalendarz/redirect.html` → Zarejestruj
4. **Uwierzytelnianie → Dodaj URI**: `https://indy97.github.io/onenote-kalendarz/dialog.html` → Zapisz
5. **Uprawnienia interfejsu API**: powinno już być `User.Read`. Dodaj
   **Microsoft Graph → Uprawnienia delegowane → `Tasks.ReadWrite`**.
   Oba nie wymagają zgody administratora.
6. Skopiuj **Identyfikator aplikacji (klienta)** i wpisz go w `docs/config.js` jako `clientId`.

Jeśli przy pierwszym logowaniu pojawi się „Wymagane zatwierdzenie administratora”, znaczy to,
że w tenancie wyłączono zgody użytkowników – wtedy GARy musi kliknąć zgodę dla tej jednej
aplikacji (tylko `User.Read` + `Tasks.ReadWrite`, uprawnienia delegowane – bardzo małe ryzyko).

## Krok 3. Wgranie wtyczki do OneNote (sideload)

1. Otwórz notes w **OneNote dla przeglądarki** (onenote.com / OneDrive / SharePoint).
2. **Wstawianie → Dodatki pakietu Office** (Office Add-ins) → **Przekaż mój dodatek**
   (Upload My Add-in) → wybierz `manifest.xml`.
3. Na karcie **Narzędzia główne** pojawi się przycisk **Kalendarz zadań**.
4. Kliknij **Zaloguj się** (przeglądarka musi zezwolić na wyskakujące okienko).

## Krok 4. Przypomnienia – przepływ Power Automate

https://make.powerautomate.com → **Utwórz → Zaplanowany przepływ w chmurze**.
Nazwa `Przypomnienia – Kalendarz OneNote`, powtarzaj co **15 minut**.

1. **Recurrence**: Interwał 15, Częstotliwość Minuta.
2. **Microsoft To-Do (Business) → List to-do's by folder (V2)**:
   Folder = `Kalendarz OneNote`.
3. **Filter array** (Filtruj tablicę) – From: `value` z poprzedniego kroku,
   tryb zaawansowany, wyrażenie:
   ```
   @and(
     equals(item()?['isReminderOn'], true),
     not(equals(item()?['status'], 'completed')),
     lessOrEquals(ticks(convertToUtc(item()?['reminderDateTime']?['dateTime'], item()?['reminderDateTime']?['timeZone'])), ticks(utcNow())),
     greater(ticks(convertToUtc(item()?['reminderDateTime']?['dateTime'], item()?['reminderDateTime']?['timeZone'])), ticks(addMinutes(utcNow(), -15)))
   )
   ```
   (zadanie trafia do przypomnienia dokładnie raz – w 15-minutowym oknie, w którym wypada jego godzina)
4. **Apply to each** na `Body` z Filter array, w środku:
   - **Office 365 Outlook → Send an email (V2)**: Do `ac@techcomserwis.pl`,
     Temat `🔔 @{items('Apply_to_each')?['title']}`,
     Treść `@{items('Apply_to_each')?['body']?['content']}`
   - **Microsoft Teams → Post message in a chat or channel**: Post as **Flow bot**,
     Post in **Chat with Flow bot**, Recipient `ac@techcomserwis.pl`,
     Message `🔔 Przypomnienie: @{items('Apply_to_each')?['title']}`
5. Zapisz i kliknij **Testuj** – dodaj w wtyczce zadanie z przypomnieniem za 5 minut.

Wszystkie connectory są standardowe (w licencji M365) i nie wymagają admina.

**Uwaga:** przepływ nie był jeszcze testowany na żywo. Jeśli filtr nic nie przepuszcza,
otwórz historię uruchomienia, podejrzyj wynik kroku 2 i sprawdź nazwy pól
(`isReminderOn`, `reminderDateTime.dateTime`, `reminderDateTime.timeZone`, `status`) –
w razie różnic dopasuj wyrażenie.

## Ograniczenia

- Działa tylko w **OneNote dla przeglądarki** (desktopowy OneNote nie obsługuje tych wtyczek).
- Zadania tylko z datą pojawiają się w kalendarzu; dodane w To Do bez terminu są pomijane.
- Ta sama lista w To Do daje dodatkowo natywne powiadomienia To Do / Outlook.
