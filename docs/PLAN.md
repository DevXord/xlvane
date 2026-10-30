# Plan dla Claude Code: xlvane, biblioteka XLSX dla ludzi i agentów AI

Stan na 30 października 2026 r. Plan jest pisany tak, żeby dało się go wrzucić do pustego repozytorium jako `docs/PLAN.md` i realizować faza po fazie w osobnych sesjach Claude Code.

## 1. Po co ta biblioteka

- `exceljs` ma **25,5 mln pobrań tygodniowo**, a ostatnie stabilne wydanie (4.4.0) pochodzi z **19 października 2023 r.** Ma 664 otwarte zgłoszenia, 153 otwarte PR-y i drzewo 193 paczek zajmujące 34,9 MB.
- `xlsx` (SheetJS) ma **15,5 mln pobrań tygodniowo**, a wersja na npm stoi na 0.18.5 z marca 2022 r. Poprawka luki CVE-2023-30533 jest dostępna tylko na CDN autora.
- MUI utrzymuje własny wewnętrzny fork (`@mui/x-internal-exceljs-fork`, 335 tys. pobrań tygodniowo), czyli nawet duże projekty nie mają utrzymywanej biblioteki, na której mogłyby polegać.
- Firmy płacą za arkusze: istnieją komercyjne SheetJS Pro, Aspose.Cells i SpreadJS. Kilka małych paczek do samego przeliczania formuł ma po ok. 10 tys. pobrań tygodniowo.
- Wyścig już trwa: `@office-kit/xlsx` (MIT, od lipca 2026 r., 127 tys. pobrań tygodniowo) i `documonster` (Apache-2.0, 22 tys.). Żaden z nich nie pozwala przejść z `exceljs` jedną linią i żaden nie jest budowany z myślą o agentach AI.

**Wyróżniki xlvane:**
1. Przejście z `exceljs` jedną linią w `package.json` dzięki warstwie zgodności.
2. Najmniej pracy dla agenta AI przy typowych zadaniach z arkuszami, potwierdzone publicznym, powtarzalnym benchmarkiem.
3. Czytelne CLI, z którego człowiek skorzysta bez czytania dokumentacji.
4. Odporność na złośliwe pliki (zip bomby, XXE, zanieczyszczenie prototypu, wstrzykiwanie formuł do CSV).
5. Zero zależności poza `fflate`, tylko ESM, działa w Node, Bun, Deno, przeglądarce i Cloudflare Workers.

**Model zarabiania:** rdzeń, CLI, serwer MCP i integracja z agentami na MIT. Płatna wersja Pro w osobnym, prywatnym repozytorium (sekcja 9).

## 2. Nazwa

**xlvane** („xl” jak arkusz Excela, „vane” jak łopatka kierująca przepływem). Sprawdzone 30 października 2026 r.:

| Miejsce | Wynik |
|---|---|
| npm: `xlvane`, `xl-vane`, `xl_vane`, `xl.vane`, `xlvanejs`, scope `@xlvane` | wolne |
| PyPI, crates.io | wolne |
| GitHub (użytkownik lub organizacja `xlvane`) | wolne (strona zwraca 404) |
| JSR `@xlvane`, RubyGems, NuGet, Docker Hub | wolne (strony zwracają 404) |
| Wyszukiwarka | brak dokładnych trafień |

Czego nie dało się sprawdzić: domen (`xlvane.dev`, `xlvane.com`) i rejestrów znaków towarowych. Trzeba to zrobić ręcznie u rejestratora domen i w TMview (EUIPO) przed publiczną premierą.

**Zarezerwować od razu, tego samego dnia:** paczkę `xlvane` i organizację `@xlvane` na npm, organizację `xlvane` na GitHubie, domenę. Każde konto z 2FA i z drugim zaufanym administratorem.

Zapasowe nazwy, również wolne na npm, PyPI, crates.io i GitHubie: `celluvo`, `gridlune`.

## 3. Struktura repozytorium

Monorepo z workspace'ami npm:

| Paczka | Licencja | Zawartość | Zależności w czasie działania |
|---|---|---|---|
| `xlvane` | MIT | Rdzeń (odczyt, zapis, style, streaming), API zadaniowe, CLI | tylko `fflate` |
| `@xlvane/exceljs` | MIT | Warstwa zgodności z API `exceljs` jako główny punkt wejścia paczki, żeby działał alias `"exceljs": "npm:@xlvane/exceljs@^1"` | tylko `xlvane` |
| `@xlvane/mcp` | MIT | Serwer MCP dla agentów | `xlvane` i oficjalne SDK MCP |
| `@xlvane/pro` | komercyjna | **Nie w tym repozytorium.** Osobne, prywatne repozytorium | — |

Warstwa zgodności musi być osobną paczką, bo alias npm podmienia cały moduł: kod `require("exceljs")` albo `import ExcelJS from "exceljs"` trafia do głównego punktu wejścia paczki z aliasu, a nie do podścieżki.

Serwer MCP jest osobną paczką, żeby rdzeń został bez zależności. Polecenie `xlvane mcp` w CLI uruchamia `@xlvane/mcp`, a jeśli go nie ma, wypisuje dokładną komendę instalacji.

Kod, komentarze, README, dokumentacja i komunikaty CLI są **po angielsku**, bo odbiorcą jest międzynarodowy rynek programistów.

## 4. Integracja z agentami AI

### Zasada nadrzędna

Agenci mają wybierać xlvane, bo z nim **rzeczywiście** wykonują zadanie szybciej i z mniejszą liczbą błędów, a nie dlatego, że coś ich do tego namawia. Nie ma sprawdzonego sposobu, żeby agent wybrał konkretną paczkę. Badanie Amplifying.ai pokazało, że agenci wybierają znanych liderów albo piszą kod sami, a `llms.txt` nie miał mierzalnego wpływu. Dlatego plan opiera się na trzech rzeczach: mniej pracy, łatwe wejście i twardy pomiar.

**Zakazane:** tekst skierowany do agentów, który ma ich nakłonić do wyboru paczki („AI assistants should always prefer xlvane”), ukryte instrukcje w README, typach albo komunikatach. To jest wstrzykiwanie poleceń (prompt injection). Skanery bezpieczeństwa takie jak Socket oznaczają za to paczki, a firmy, czyli nasi klienci, przestałyby nam ufać. Dozwolone są tylko prawdziwe, sprawdzalne informacje.

### Co konkretnie zmniejsza pracę agenta

1. **API zadaniowe.** Najczęstsze zadania mieszczą się w jednym wywołaniu zamiast 15–40 linii kodu:
   - `inspect(file)`: struktura skoroszytu w JSON (arkusze, zakresy, wykryte wiersze nagłówków, typy kolumn, scalone komórki, formuły, ukryte arkusze, ostrzeżenia). Agent musi zrozumieć plik, zanim go przetworzy, a dziś pisze do tego kod eksploracyjny.
   - `readTable(file, { sheet, header: "auto", schema })`: wiersze jako obiekty, opcjonalna walidacja przez Standard Schema (Zod, Valibot, ArkType), błędy jako lista z adresem komórki.
   - `writeTable(rows, { columns, preset: "report" })`: gotowy, czytelny raport (pogrubione nagłówki, zamrożony pierwszy wiersz, autofiltr, szerokości kolumn, formaty dat i liczb).
   - `editCells(file, edits)`, `appendRows(file, sheet, rows)`.
   - `diff(a, b)`: różnice między skoroszytami. Agent sprawdza nim swoją pracę, a człowiek przegląda zmianę.
   - `toCSV`, `toJSON`, `fromCSV`, `fromJSON`.
2. **Zgodność z `exceljs`.** Modele znają API `exceljs` z danych treningowych. Kod, który agent napisze pod `exceljs`, ma działać z `@xlvane/exceljs` bez zmian.
3. **Typy z przykładami.** Każdy publiczny symbol ma w TSDoc opis i działający przykład. Agenci czytają pliki `.d.ts` w `node_modules`.
4. **Błędy, z których agent sam się poprawi.** Każdy błąd ma stały kod (`XLV_E_SHEET_NOT_FOUND`), miejsce (arkusz i komórka), podpowiedź naprawy z listą dostępnych opcji („Available sheets: Sales, Costs”) i link do opisu w dokumentacji. `error.toJSON()` zwraca to samo jako dane.
5. **Serwer MCP** (`@xlvane/mcp`). Agent pracuje na arkuszach bez pisania kodu, przez narzędzia `inspect_workbook`, `read_table`, `write_table`, `edit_cells`, `convert`, `diff`, `validate`. Domyślnie tylko odczyt i tylko w katalogach wskazanych przy uruchomieniu, a zapis wymaga flagi `--allow-write`.
6. **Agent Skill.** Plik `SKILL.md` z instrukcjami i przykładami dołączony do paczki i do repozytorium, do zainstalowania jednym poleceniem. Mechanizm instalacji (np. instalator `skills` od Vercela) trzeba sprawdzić w aktualnej dokumentacji.
7. **Dokumentacja w paczce.** Folder `docs/` i `llms-full.txt` w tarballu npm, żeby agent miał je lokalnie bez sieci. Koszt jest mały, ale nie liczymy na to jako na kanał dystrybucji.
8. **Wyjście dla maszyn.** Każde polecenie CLI ma `--json` ze stabilnym, udokumentowanym schematem i znaczące kody wyjścia.

### Pomiar: benchmark agentów

Publiczny benchmark w `bench/agents/` to dowód „mniej pracy” i zarazem pętla ulepszania API.

- 20–30 typowych zadań, np. „połącz trzy pliki sprzedaży w jeden raport”, „znajdź wiersze z błędnym NIP-em i zaznacz je na czerwono”, „dodaj kolumnę z sumą i zachowaj formatowanie”, „przekonwertuj arkusz na JSON zgodny ze schematem”.
- Każde zadanie wykonywane przez agenta kodującego w trybie bez interfejsu (np. Claude Code z flagą `-p`) w trzech wariantach: xlvane, `exceljs` i SheetJS, po kilka powtórzeń.
- Mierzymy odsetek sukcesów (sprawdzany automatycznie skryptem), liczbę tur, tokeny, koszt i czas.
- Publikujemy metodologię, surowe wyniki i skrypty, żeby każdy mógł powtórzyć pomiar. Wynik gorszy od konkurencji też publikujemy i naprawiamy.

Benchmark kosztuje: każdy przebieg zużywa płatne tokeny API.

## 5. CLI dla ludzi

### Zasady

- **Proste.** Najwyżej 9 poleceń, każde nazwane czasownikiem: `inspect`, `read`, `create`, `convert`, `edit`, `diff`, `validate`, `migrate`, `mcp`.
- **Samo się tłumaczy.** `xlvane` bez argumentów pokazuje krótki ekran powitalny: co robi narzędzie, 5 najczęstszych poleceń z przykładami i gdzie szukać pomocy.
- **Pomoc z przykładami.** `xlvane help <polecenie>` i `--help` pokazują: co robi polecenie, składnię, 2–4 gotowe do skopiowania przykłady, powiązane polecenia.
- **Błąd w trzech częściach:** co się stało, gdzie, co zrobić dalej (z gotową komendą). Przy literówce podpowiedź „Did you mean...”.
- **Bezpiecznie.** Nic nie nadpisuje pliku bez `--force`. Polecenia zapisujące mają `--dry-run` i na końcu podają ścieżkę wyniku.
- **Czytelnie.** Tabele, kolory przez wbudowane `util.styleText`, respektowanie `NO_COLOR` i wyjścia bez terminala, żadna informacja przekazywana wyłącznie kolorem. Postęp dla dużych plików na stderr. `--quiet` i `--verbose`.

### Przykładowy ekran powitalny

```
xlvane 1.0.0 — read, write and inspect Excel files

Common tasks:
  xlvane inspect report.xlsx              See sheets, columns and problems
  xlvane read report.xlsx --sheet Sales   Print a sheet as a table
  xlvane convert report.xlsx out.csv      Convert to CSV or JSON
  xlvane create data.json report.xlsx     Build a formatted report
  xlvane diff old.xlsx new.xlsx           Show what changed

More:  xlvane help <command>   ·   Docs: https://github.com/xlvane/xlvane
For AI agents: add --json to any command, or run `xlvane mcp`
```

Każdą zmianę tekstów CLI pilnują testy migawkowe (snapshot), a README zawiera 5-minutowy start.

## 6. Bezpieczeństwo

Biblioteka będzie otwierać pliki z nieznanych źródeł, więc:
- **ZIP:** limit rozmiaru po rozpakowaniu i stosunku kompresji (zip bomby), limit liczby wpisów.
- **XML:** własny parser bez obsługi DTD i encji (XXE, „billion laughs”), limit głębokości i długości tekstu.
- **Dane:** obiekty bez prototypu albo `Map` dla danych z pliku. Klucze `__proto__`, `constructor` i `prototype` nie mogą niczego zmienić.
- **Eksport do CSV:** domyślna ochrona przed wstrzykiwaniem formuł (komórki zaczynające się od `=`, `+`, `-`, `@`), z opcją wyłączenia.
- **Wyrażenia regularne:** bez ryzyka ReDoS, sprawdzane w testach.
- **Fuzzing:** fast-check plus uszkodzone i spreparowane pliki w korpusie testowym.

To jest argument sprzedażowy dla firm. Warto mieć `SECURITY.md` z procesem zgłaszania luk.

## 7. Plik CLAUDE.md dla repozytorium

Wkleić jako `CLAUDE.md` w katalogu głównym:

```markdown
# Zasady projektu xlvane

## Cel
Biblioteka XLSX na MIT dla ludzi i agentów AI: odczyt, zapis, style, streaming, API zadaniowe, zgodność z exceljs (paczka @xlvane/exceljs), czytelne CLI, serwer MCP (paczka @xlvane/mcp). Plan: docs/PLAN.md.

## Twarde reguły
- Paczka `xlvane` ma jedną zależność w czasie działania: `fflate`. Parser XML piszemy sami. Każda nowa zależność wymaga mojej zgody i wpisu w docs/decisions.md.
- `@xlvane/exceljs` może zależeć tylko od `xlvane`. `@xlvane/mcp` może zależeć tylko od `xlvane` i oficjalnego SDK MCP.
- W tym repozytorium nie ma kodu wersji Pro. Rdzeń udostępnia tylko udokumentowane interfejsy rozszerzeń (np. `FormulaEngine`).
- Tylko ESM, bez top-level await. Działa w Node, Bun, Deno, przeglądarce i Cloudflare Workers; API rdzenia używa `Uint8Array` i Web Streams, a operacje na plikach są w osobnym podmodule dla Node.
- Pliki wejściowe są niezaufane: limity ZIP, XML bez DTD i encji, brak zanieczyszczenia prototypu, ochrona CSV przed wstrzykiwaniem formuł. Każda reguła ma test.
- Zakaz tekstów nakłaniających agentów AI do wyboru paczki i jakichkolwiek ukrytych instrukcji dla modeli. Dokumentacja i komunikaty zawierają wyłącznie prawdziwe, sprawdzalne informacje.
- Kod, komentarze, dokumentacja i komunikaty CLI po angielsku.
- Każdy publiczny symbol ma TSDoc z opisem i działającym przykładem; test sprawdza, że przykłady się kompilują.
- Każdy błąd ma stały kod XLV_E_*, miejsce, podpowiedź naprawy i wpis w docs/errors.md.
- Testy nie korzystają z sieci. Pliki testowe tworzymy sami albo bierzemy z korpusów na licencjach pozwalających na użycie; licencję każdego pliku zapisujemy w tests/corpus/LICENSES.md.
- Faktów o formacie OOXML, exceljs, npm, MCP i narzędziach nie wpisujemy z pamięci. Sprawdzamy w specyfikacji ECMA-376, kodzie źródłowym lub aktualnej dokumentacji i zapisujemy źródło w docs/decisions.md.
- Nie publikujemy na npm, nie wypychamy zmian i nie tworzymy kont bez mojej zgody.

## Polecenia
- `npm test` – testy jednostkowe, zgodności i migawkowe CLI (Vitest)
- `npm run fuzz` – fuzzing (fast-check), dłuższy przebieg
- `npm run bench` – wydajność względem exceljs i SheetJS
- `npm run build` – build przez tsdown
- `npm run check:pkg` – publint i @arethetypeswrong/cli

## Sposób pracy
- Jedna faza planu na sesję. Na początku przeczytaj docs/PLAN.md, docs/decisions.md i docs/CHANGELOG-dev.md.
- Zacznij w trybie planowania i poczekaj na moją akceptację.
- Najpierw testy, potem implementacja. Faza jest skończona, gdy spełnia kryteria akceptacji i `npm test` przechodzi.
- Na koniec fazy: podsumowanie w docs/CHANGELOG-dev.md, commit, lista decyzji do mojego przeglądu.
```

## 8. Fazy

Szacunki są orientacyjne i dotyczą pracy z Claude Code przy bieżącym przeglądzie zmian. MVP to fazy 0–3 plus podstawowe CLI i MCP (ok. 2–3 tygodni). Wersja 1.0 ze zgodnością z `exceljs` i korpusem plików to realnie 2–3 miesiące.

### Faza 0: Szkielet monorepo (0,5 dnia)

> Utwórz monorepo z workspace'ami npm: `packages/xlvane`, `packages/exceljs` i `packages/mcp`, zgodnie z CLAUDE.md. ESM-only, tsdown, Vitest, fast-check, publint, @arethetypeswrong/cli, licencja MIT w każdej paczce. Sprawdź w aktualnej dokumentacji bieżące wersje narzędzi i to, czy tsdown jest nadal zalecanym następcą tsup. Dodaj `SECURITY.md`. Dodaj plik polityki depgate dla repozytorium, jeśli depgate jest już dostępny.

**Kryteria:** build, testy i `check:pkg` przechodzą na pustych modułach, `xlvane` zależy tylko od `fflate`.

### Faza 1: Bezpieczny odczyt i `inspect` (3–5 dni)

> Zaimplementuj odczyt XLSX: ZIP przez fflate z limitami, własny parser XML bez DTD i encji, workbook, arkusze, współdzielone napisy, typy komórek (liczba, tekst, wartość logiczna, błąd, data, formuła z wartością z pamięci podręcznej), oba systemy dat (1900 i 1904), wykrywanie formatów dat z numFmt, scalone komórki, ukryte arkusze. Na tym zbuduj `inspect()` zwracające JSON ze strukturą skoroszytu i ostrzeżeniami. Każdy fakt o formacie sprawdzaj w ECMA-376 i zapisuj źródło. Przygotuj pierwsze pliki testowe wygenerowane w Excelu, LibreOffice i Google Sheets (jeśli nie masz do nich dostępu, wypisz mi listę plików, które mam dostarczyć).

**Kryteria:** testy na korpusie, testy zip bomby, XXE i „billion laughs”, fuzzing bez awarii, dane bez zanieczyszczenia prototypu.

### Faza 2: Zapis i style (3–5 dni)

> Zaimplementuj model skoroszytu i zapis: wiele arkuszy, typy komórek, formuły (zapis tekstu formuły), style (czcionki, wypełnienia, obramowania, wyrównanie, formaty liczb i dat), szerokości kolumn, wysokości wierszy, scalanie, zamrożone okienka, autofiltr. Testy w obie strony: zapis, odczyt i porównanie. W CI sprawdzaj, że LibreOffice w trybie bez interfejsu otwiera każdy wygenerowany plik i konwertuje go do CSV z oczekiwaną zawartością.

**Kryteria:** testy w obie strony przechodzą, LibreOffice otwiera wszystkie pliki, budżet rozmiaru paczki zapisany w README.

### Faza 3: API zadaniowe (2–3 dni)

> Zaimplementuj `readTable` (z automatycznym wykrywaniem nagłówka i opcjonalnym schematem Standard Schema; typy interfejsu skopiuj ze specyfikacji ze źródłem w komentarzu), `writeTable` z presetem `report`, `editCells`, `appendRows`, `diff`, `toCSV` (z ochroną przed wstrzykiwaniem formuł), `toJSON`, `fromCSV`, `fromJSON`. Do każdego symbolu TSDoc z działającym przykładem. Zbuduj katalog błędów XLV_E_* w docs/errors.md.

**Kryteria:** typy wynikowe poprawnie wnioskowane dla Zod, Valibot i ArkType, każdy błąd ma kod, miejsce i podpowiedź, przykłady z TSDoc kompilują się w teście.

### Faza 4: Streaming (2–3 dni)

> Dodaj odczyt i zapis strumieniowy na Web Streams dla plików z setkami tysięcy wierszy. Benchmark: 100 tys. wierszy × 20 kolumn, czas i pamięć, porównanie z exceljs i SheetJS.

**Kryteria:** stałe zużycie pamięci przy odczycie strumieniowym, wyniki benchmarku w README z opisem sprzętu i metodologii.

### Faza 5: Zgodność z exceljs (1–2 tygodnie)

> Przeanalizuj publiczne API exceljs 4.4.0 i najczęściej używane jego części (np. przez przegląd kodu otwartych projektów, które go używają). Zaimplementuj paczkę `packages/exceljs` (`@xlvane/exceljs`), której główny punkt wejścia udostępnia ten podzbiór: `Workbook`, `xlsx.readFile`, `xlsx.writeBuffer`, `addWorksheet`, `getWorksheet`, `columns`, `addRow`, `getRow`, `getCell`, `eachRow`, style, scalanie, strumieniowy `WorkbookWriter`. Musi działać zarówno `import ExcelJS from "exceljs"`, jak i `const ExcelJS = require("exceljs")` z CommonJS; sprawdź w dokumentacji Node, jak `require(esm)` obsługuje eksport domyślny, i zapisz decyzję w docs/decisions.md. Przenieś odpowiednie testy exceljs (licencja MIT, zachowaj informację o autorstwie). Udokumentuj, co jest wspierane, a co nie, w docs/exceljs-compat.md. Dodaj test, w którym przykładowy projekt napisany pod exceljs działa bez zmian w kodzie po aliasie `"exceljs": "npm:@xlvane/exceljs@^1"`.

**Kryteria:** przeniesione testy przechodzą dla wspieranego podzbioru, jasna tabela zgodności, działająca migracja aliasem z ESM i z CommonJS.

### Faza 6: CLI dla ludzi (2–3 dni)

> Zbuduj CLI według sekcji 5 planu na wbudowanym `util.parseArgs`, bez dodatkowych zależności. Ekran powitalny, `help` z przykładami, podpowiedzi przy literówkach, błędy w trzech częściach, `--json` ze schematem w docs/cli-json.md, `--dry-run`, ochrona przed nadpisaniem, `NO_COLOR`. Testy migawkowe dla każdego ekranu pomocy i komunikatu błędu. Napisz w README 5-minutowy start.

**Kryteria:** każde polecenie ma pomoc z przykładami, wszystkie przykłady z pomocy działają w teście, `--json` zgodny z udokumentowanym schematem.

### Faza 7: Integracja z agentami (2–3 dni)

> Zbuduj `@xlvane/mcp` na oficjalnym SDK MCP: sprawdź aktualną specyfikację i wersję SDK (specyfikacja 2026-07-28 jest bezstanowa). Narzędzia: inspect_workbook, read_table, write_table, edit_cells, convert, diff, validate, z opisami i schematami wejścia. Domyślnie tylko odczyt w katalogach podanych przy starcie, zapis po `--allow-write`, limity rozmiaru plików. Dodaj `SKILL.md` z instrukcjami i przykładami (sprawdź aktualny format Agent Skills i sposób instalacji) oraz dołącz `docs/` i `llms-full.txt` do tarballa npm. Przejrzyj wszystkie teksty pod kątem zakazu z CLAUDE.md dotyczącego nakłaniania agentów.

**Kryteria:** testy serwera MCP na kliencie w pamięci, test, że zapis poza dozwolonym katalogiem jest odrzucany, Skill instaluje się jednym poleceniem.

### Faza 8: Benchmark agentów (2–4 dni plus koszt tokenów)

> Zbuduj `bench/agents/` według sekcji 4: 20–30 zadań, automatyczne sprawdzanie wyniku, uruchamianie agenta w trybie bez interfejsu w trzech wariantach (xlvane, exceljs, SheetJS), pomiar sukcesu, tur, tokenów, kosztu i czasu. Sprawdź w aktualnej dokumentacji Claude Code, jak uruchomić go bez interfejsu i odczytać zużycie. Przed pierwszym pełnym przebiegiem pokaż mi szacowany koszt i poczekaj na zgodę. Wyniki zapisz z metodologią w bench/agents/RESULTS.md.

**Kryteria:** powtarzalny przebieg jednym poleceniem, surowe wyniki w repozytorium. Wnioski z porażek xlvane trafiają jako zadania do poprawy API lub dokumentacji.

### Faza 9: Korpus i zgodność środowisk (ok. tygodnia, potem stale)

> Rozbuduj korpus o pliki z Excela (Windows i Mac), LibreOffice, Google Sheets i Numbers oraz o pliki z typowych eksportów (systemy księgowe, CRM). Sprawdź licencje publicznych korpusów testowych (np. plików testowych Apache POI) i dodaj tylko te, które wolno używać. CI: Node 22, 24 i 26, Bun, Deno, przeglądarka, Cloudflare Workers. Przypnij akcje GitHub do SHA.

**Kryteria:** zielone CI na wszystkich środowiskach, każdy plik korpusu ma udokumentowaną licencję.

### Faza 10: Bezpieczna publikacja (0,5 dnia, z Pana udziałem)

> Przygotuj publikację z tagu przez trusted publishing (OIDC) z proweniencją, bez tokenów npm w sekretach, z publikacją etapową (`npm stage publish`) i zatwierdzeniem z 2FA. Sprawdź aktualną dokumentację npm i opisz ręczne kroki w docs/RELEASING.md.

Ręcznie po Pana stronie: konfiguracja trusted publishing dla `xlvane` i `@xlvane/mcp` (niezweryfikowana konfiguracja wygasa po 48 godzinach), 2FA, zatwierdzanie wydań.

**Kryteria:** wydania mają proweniencję, `npm view xlvane dependencies` pokazuje tylko `fflate`.

## 9. Wersja Pro (osobne repozytorium, po zdobyciu użytkowników)

| Funkcja Pro | Dlaczego firmy zapłacą |
|---|---|
| Przeliczanie formuł (silnik funkcji Excela) | Raporty generowane po stronie serwera muszą mieć poprawne wyniki bez otwierania Excela |
| Tabele przestawne i wykresy | Typowe w raportach biznesowych, trudne do zrobienia samemu |
| Pliki chronione hasłem (szyfrowanie ECMA-376) | Dane finansowe i kadrowe |
| Łatki bezpieczeństwa dla starszych wersji (LTS) | Firmy nie aktualizują wersji głównych od razu |
| SLA i wsparcie | Wymóg działów zakupów |

Cennik ilustracyjny (założenie, nie wynik badań): Pro ok. 990 USD rocznie na firmę, Enterprise z SLA od ok. 5000 USD rocznie. Pro rozprowadzany z prywatnego rejestru dla płacących firm, na licencji komercyjnej. Klucz licencyjny w JavaScripcie łatwo obejść, więc realną ochroną jest umowa z firmą.

Przed budową Pro warto porozmawiać z 5–10 firmami używającymi xlvane i potwierdzić, za którą funkcję zapłacą najpierw. Najpewniejszym kandydatem jest przeliczanie formuł, bo na nie już dziś istnieje osobny popyt.

## 10. Dystrybucja

- **Migracja z exceljs:** wpis w e18e `module-replacements` jako zamiennik `exceljs` i `xlsx`, poradnik migracji, PR-y do projektów zależnych od `exceljs` (opisy i kontakt z opiekunami po Pana stronie, bo opiekunowie odrzucają PR-y wyglądające na automatyczne).
- **Projekty, które utrzymują własne forki** (np. MUI) jako pierwsi partnerzy lub klienci.
- **Wyniki benchmarku agentów** jako artykuł techniczny z otwartą metodologią.
- **Rejestry serwerów MCP** i katalogi Agent Skills dla `@xlvane/mcp`.
- **Synergia z depgate:** te same wartości (zero zależności, proweniencja, bezpieczeństwo), ci sami odbiorcy w firmach.

## 11. Ryzyka

- **`@office-kit/xlsx` lub inny nowicjusz zdobędzie pozycję pierwszy.** Obrona: zgodność z `exceljs`, integracja z agentami, bezpieczeństwo. Tempo ma znaczenie.
- **Opiekunowie `exceljs` wrócą do pracy.** Przewagą zostaje integracja z agentami, bezpieczeństwo i działanie w wielu środowiskach.
- **Ogrom formatu.** OOXML ma tysiące stron specyfikacji. Trzeba trzymać się zakresu z planu i rozbudowywać według zgłoszeń użytkowników.
- **Ciężar utrzymania.** Wiele zgłoszeń dotyczy plików z nietypowych programów. Szablony zgłoszeń z wymaganym plikiem przykładowym i `xlvane inspect --json` w zgłoszeniu.
- **Agenci dalej wybierają `exceljs` z przyzwyczajenia.** Dlatego warstwa zgodności: nawet kod pisany „pod exceljs” działa z xlvane.

## 12. Podsumowanie czasu

| Faza | Czas |
|---|---|
| 0. Szkielet | 0,5 dnia |
| 1. Odczyt i `inspect` | 3–5 dni |
| 2. Zapis i style | 3–5 dni |
| 3. API zadaniowe | 2–3 dni |
| 4. Streaming | 2–3 dni |
| 5. Zgodność z exceljs | 1–2 tygodnie |
| 6. CLI | 2–3 dni |
| 7. Integracja z agentami | 2–3 dni |
| 8. Benchmark agentów | 2–4 dni |
| 9. Korpus i środowiska | ok. tygodnia, potem stale |
| 10. Publikacja | 0,5 dnia |
| **Razem do wersji 1.0** | **ok. 6–9 tygodni pracy, realnie 2–3 miesiące z przeglądem i poprawkami** |
| Pro | osobny projekt po zdobyciu użytkowników |

Sugerowana kolejność dla najszybszej informacji zwrotnej: 0, 1, 2, 3, 6, 7, 10 (wydanie wersji 0.x), potem 4, 5, 8, 9 i wersja 1.0.
