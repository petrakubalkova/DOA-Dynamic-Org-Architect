## Dynamic Org Architect

Dynamic Org Architect je nástroj pro **analýzu vytížení manažerů (Span of Control)** a **vizualizaci komplexních organizačních struktur**. Umožňuje rychle odhalit přetížené, podvyužité i „multicompany“ manažery (řídící více firem/entit) a prezentovat výsledky jak staticky (PNG), tak interaktivně (HTML).

### Hlavní přínosy

- **Analýza Span of Control (SPOC)**: počítá, kolik přímých podřízených (zejména specialistů/pracovníků) má každý manažer (profese 1–3).
- **Detekce Multicompany manažerů**: identifikuje manažery, kteří vedou týmy složené z více společností, a zvýrazní je ve vizualizaci.
- **Čitelné vizualizace**: generuje přehledné organizační schéma ve formátu PNG i interaktivní síť v HTML pro hlubší analýzu vztahů.

### Klíčové vlastnosti

- **Detekce „Multicompany“ manažerů**:  
  Manažeři, jejichž přímí podřízení pocházejí z více různých společností, jsou označeni příznakem *Multicompany*, mají zvýrazněný rámeček a v tooltipu i seznam spravovaných společností.

- **Barevné kódování podle vytížení**:  
  - **Zelená**: optimální nebo vysoký SPOC (SPOC ≥ 5, včetně CEO).  
  - **Červená**: nízký SPOC (SPOC < 3).  
  - **Modrá**: střední vytížení (SPOC 3–4).  

- **Export do statického PNG i interaktivního HTML**:  
  - Statický export do `Organizacni_Schema.png` je vhodný pro **prezentace, reporty a tisk (A4)**.  
  - Interaktivní export do `interaktivni_schema.html` umožňuje **zoom, posun, proklikávání a čtení detailních tooltipů** v prohlížeči.

- **Inteligentní hierarchické řazení**:  
  - Do vizualizací jsou zahrnuti pouze **manažeři** (profese 1 – CEO, 2 – Manažer, 3 – Vedoucí týmu).  
  - Specialisté a pracovníci (profese 4, 5) jsou započteni do SPOC, ale **nezahlcují schéma** (nevykreslují se jako uzly).  
  - Hierarchie je renderována **shora dolů (UD)**, s řízenými rozestupy mezi patry a uzly, aby zůstala čitelná i u rozsáhlejších struktur.
  
- **Živé demo**: [Prohlédnout interaktivní schéma online](https://petrakubalkova.github.io/DOA-Dynamic-Org-Architect/interaktivni_schema.html)

### Požadované knihovny a instalace

Projekt je postaven na Pythonu 3 a využívá následující balíčky:

- `pandas` – práce s tabulkovými daty (Excel, DataFrame)
- `matplotlib` – kreslení statického PNG organizačního schématu
- `pyvis` – generování interaktivní síťové vizualizace v HTML
- `openpyxl` – engine pro čtení/zápis Excel souborů (XLSX)

Instalaci doporučených balíčků můžete provést například takto:

```bash
pip install pandas matplotlib pyvis openpyxl
```

### Struktura projektu (hlavní skript)

Hlavní logika je v souboru `analyza_spoc.py` a typický průběh běhu vypadá takto:

1. **Načtení nebo generování dat**  
   - Pokud existuje `Org_struktura_vypis.xlsx`, skript jej načte jako zdrojovou organizační strukturu.  
   - Pokud soubor neexistuje, vygeneruje **testovací data** (včetně vícero firem) a uloží je do `Testovaci_Struktura.xlsx` na plochu.

2. **Výpočet SPOC a Multicompany**  
   - Pro všechny manažery (profese 1–3) se spočítá:
     - **SPOC** – počet přímých podřízených (specialistů a pracovníků).  
     - **Multicompany příznak** – zda jejich přímí podřízení pocházejí z více firem.  
     - **Počet_firem** – kolik unikátních společností manažer v týmu má.  
   - Výsledky se ukládají do tabulky `spoc_df` a vypíší do konzole.

3. **Vytvoření statického PNG schématu** (`vizualizuj_org_schema`)  
   - Generuje se top‑down strom manažerské hierarchie (CEO → Manažeři → Vedoucí týmů).  
   - Každý box obsahuje: **jméno, profesi, organizační jednotku, velikost týmu (SPOC) a jméno nadřízeného**.  
   - Multicompany manažeři mají **silnější/odlišný rámeček** a text `[MULTICOMPANY]`.  
   - Výsledný obrázek `Organizacni_Schema.png` je uložen na plochu.

4. **Vytvoření interaktivního HTML schématu** (`vytvor_interaktivni_sit`)  
   - Vytvoří se interaktivní síť pouze z manažerů (profese 1–3) v hierarchickém layoutu `UD`.  
   - Uzly jsou vykresleny jako **obdélníkové boxy** s popisem (jméno, OJ, profese) a barevným kódováním vytížení.  
   - Tooltipy zobrazují **SPOC** a u Multicompany manažerů také **seznam společností v týmu**.  
   - Výsledný HTML soubor `interaktivni_schema.html` je uložen na plochu a lze jej otevřít v libovolném prohlížeči.

### Příprava vstupního Excelu

Pokud chcete pracovat s vlastními reálnými daty, připravte soubor `Org_struktura_vypis.xlsx` se sloupci:

- `ID_zaměstnance` – jedinečný identifikátor osoby  
- `Jméno_zaměstnance` – jméno a příjmení zaměstnance  
- `ID_organizační_jednotka` – kód organizační jednotky  
- `Název_organizační_jednotka` – název organizační jednotky  
- `ID_profese` – číselník profese (1 = CEO, 2 = Manažer, 3 = Vedoucí týmu, 4 = Specialista, 5 = Pracovník)  
- `Název_profese` – textový název profese  
- `úvazek` – velikost úvazku (např. 1.0, 0.8, 0.5)  
- `Nadřízený_ID` – ID nadřízeného zaměstnance (podle `ID_zaměstnance`)  
- `Jméno_nadřízený` – jméno nadřízeného (volitelné, používá se pro popisky)  
- `Společnost` – název nebo kód společnosti (pro detekci Multicompany)  

Skript se snaží být robustní vůči drobným rozdílům v názvech sloupců (přemapování podle pořadí a názvů).

### Jak skript spustit

1. Ujistěte se, že máte nainstalované požadované knihovny (viz sekce Instalace).  
2. Umístěte `analyza_spoc.py` a případný `Org_struktura_vypis.xlsx` do stejné složky.  
3. V příkazové řádce (PowerShell, CMD, Terminal) přejděte do adresáře projektu:

```bash
cd "cesta/k/vasi/slozce/Dynamic-Org-Architect"
```
(zde si doplňte vlastní cestu na úložiště)

4. Spusťte skript:

```bash
python analyza_spoc.py
```

5. Po doběhu skriptu najdete na ploše:
   - `Organizacni_Schema.png` – statické organizační schéma.  
   - `interaktivni_schema.html` – interaktivní síťové schéma (otevřete v prohlížeči).

### Interpretace výsledků

- **Tabulka SPOC (konzole)** – přehled manažerů s jejich SPOC, Multicompany příznakem a počtem firem.  
- **PNG schéma** – reprezentuje manažerskou hierarchii v podobě stromu, vhodné k tištěné prezentaci a rychlému přehledu.

  ![Statické schéma](Organizacni_Schema.png)
  
- **HTML schéma (Pyvis)** – dovoluje přiblížení, posun, proklikávání komplexních větví a snadné čtení detailů u konkrétních manažerů.

Dynamic Org Architect je určen pro HR, top management, architekty organizace a konzultanty, kteří potřebují **rychle pochopit, jak je management vytížen a jak jsou týmy strukturovány napříč více společnostmi**.

