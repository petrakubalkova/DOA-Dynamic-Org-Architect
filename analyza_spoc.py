import pandas as pd
import os

def vytvor_testovaci_zamestnance(n=100, vystup='Testovaci_Struktura.xlsx'):
    import numpy as np
    import random

    profese_map = {
        1: 'CEO',
        2: 'Manažer',
        3: 'Vedoucí týmu',
        4: 'Specialista',
        5: 'Pracovník'
    }
    org_jedn_map = {
        1: 'Vývoj', 2: 'Obchod', 3: 'Marketing', 4: 'Výroba', 5: 'IT', 6: 'Logistika',
        7: 'HR', 8: 'Finance', 9: 'Support', 10: 'Expanze'
    }
    firmy = ['AlphaCorp', 'BetaInc', 'DeltaTech']

    rows = []
    next_id = 1

    ceo_list = []
    for spol in firmy:
        jmeno = f"{spol}_CEO"
        rows.append({
            'ID_zaměstnance': next_id, 'Jméno_zaměstnance': jmeno,
            'ID_organizační_jednotka': 0, 'Název_organizační_jednotka': 'Hlavní vedení',
            'ID_profese': 1, 'Název_profese': 'CEO',
            'úvazek': 1.0, 'Nadřízený_ID': None, 'Jméno_nadřízený': '',
            'Společnost': spol
        })
        ceo_list.append((next_id, spol, jmeno))
        next_id += 1

    ceo_ids = [x[0] for x in ceo_list]

    manazeri_count = 6
    manazeri = []
    for _ in range(manazeri_count):
        id_ceo, spol, ceo_jm = random.choice(ceo_list)
        id_oj = random.randint(1, 10)
        jmeno = f"Man_{next_id}"
        rows.append({
            'ID_zaměstnance': next_id, 'Jméno_zaměstnance': jmeno,
            'ID_organizační_jednotka': id_oj,
            'Název_organizační_jednotka': org_jedn_map[id_oj],
            'ID_profese': 2, 'Název_profese': 'Manažer',
            'úvazek': 1.0,
            'Nadřízený_ID': id_ceo,
            'Jméno_nadřízený': ceo_jm,
            'Společnost': spol
        })
        manazeri.append((next_id, spol, jmeno, id_ceo))
        next_id += 1

    vedouci_count = 18
    vedouci_list = []
    for _ in range(vedouci_count):
        manazer = random.choice(manazeri)
        man_id, spol, man_jm, id_ceo = manazer
        id_oj = random.randint(1, 10)
        jmeno = f"Lead_{next_id}"
        rows.append({
            'ID_zaměstnance': next_id, 'Jméno_zaměstnance': jmeno,
            'ID_organizační_jednotka': id_oj,
            'Název_organizační_jednotka': org_jedn_map[id_oj],
            'ID_profese': 3, 'Název_profese': 'Vedoucí týmu',
            'úvazek': 1.0,
            'Nadřízený_ID': man_id,
            'Jméno_nadřízený': man_jm,
            'Společnost': spol
        })
        vedouci_list.append((next_id, spol, jmeno, man_id))
        next_id += 1

    zbyva = n - len(rows)
    for _ in range(zbyva):
        leader_id, leader_spol, leader_jm, man_id = random.choice(vedouci_list)
        id_oj = random.randint(1, 10)
        typ = random.choices([4, 5], weights=(0.4, 0.6))[0]
        jmeno = (
            f"Spec_{next_id}" if typ == 4 else f"Prac_{next_id}"
        )
        rows.append({
            'ID_zaměstnance': next_id, 'Jméno_zaměstnance': jmeno,
            'ID_organizační_jednotka': id_oj,
            'Název_organizační_jednotka': org_jedn_map[id_oj],
            'ID_profese': typ, 'Název_profese': profese_map[typ],
            'úvazek': round(random.choices([1.0, 0.75, 0.5, 0.25], weights=(0.75,0.06,0.13,0.06))[0], 2),
            'Nadřízený_ID': leader_id,
            'Jméno_nadřízený': leader_jm,
            'Společnost': leader_spol
        })
        next_id += 1

    df = pd.DataFrame(rows)
    df.to_excel(vystup, index=False)

def spoc_vypocet(df):
    """
    Spočítá Span of Control (SPOC) a příznak Multicompany pro všechny manažery.

    Vstup:
        df (pandas.DataFrame): Organizační data s minimálně sloupci
            'ID_zaměstnance', 'ID_profese', 'Nadřízený_ID', 'Společnost'.
            Manažerské profese jsou 1 (CEO), 2 (Manažer), 3 (Vedoucí týmu).

    Logika:
        - Pro každého manažera (profese 1–3) spočítá počet jeho přímých podřízených
          (SPOC – bez ohledu na profesi).
        - Z přímých podřízených zjistí množinu společností a určí, zda manažer
          řídí tým složený z více firem (Multicompany).

    Výstup:
        spoc_dict (dict[int, int]): mapování ID manažera -> SPOC (počet přímých podřízených).
        multicompany_dict (dict[int, bool]): ID manažera -> True/False, zda má tým z více firem.
        manager_companies_dict (dict[int, list[str]]): ID manažera -> seznam společností v týmu.
    """
    spoc_dict = {}
    multicompany_dict = {}
    manager_companies_dict = {}
    for _, row in df[df['ID_profese'].isin([1, 2, 3])].iterrows():
        emp_id = row['ID_zaměstnance']
        podrizeni = df[df['Nadřízený_ID'] == emp_id]
        spoc_dict[emp_id] = len(podrizeni)
        spol_list = podrizeni['Společnost'].dropna().unique()
        multicompany_dict[emp_id] = len(spol_list) > 1
        manager_companies_dict[emp_id] = list(spol_list)
    return spoc_dict, multicompany_dict, manager_companies_dict

def spoc_vypis_table(df, spoc_dict, multicompany_dict):
    """
    Vytvoří souhrnnou tabulku SPOC a Multicompany pro manažery.

    Vstup:
        df (pandas.DataFrame): Organizační data.
        spoc_dict (dict[int, int]): SPOC hodnoty pro manažery.
        multicompany_dict (dict[int, bool]): příznak Multicompany pro manažery.

    Logika:
        - Vyfiltruje pouze manažerské profese (1–3).
        - Pro každého manažera sestaví řádek se základními údaji, jeho SPOC
          a textovým označením Multicompany (Ano/Ne).

    Výstup:
        pandas.DataFrame: tabulka se sloupci
            ['ID_zaměstnance', 'Jméno', 'Profese', 'Společnost', 'SPOC', 'Multicompany'].
    """
    rows = []
    for idx, row in df[df['ID_profese'].isin([1,2,3])].iterrows():
        emp_id = row['ID_zaměstnance']
        rows.append({
            'ID_zaměstnance': emp_id,
            'Jméno': row['Jméno_zaměstnance'],
            'Profese': row['Název_profese'],
            'Společnost': row['Společnost'],
            'SPOC': spoc_dict.get(emp_id, 0),
            'Multicompany': 'Ano' if multicompany_dict.get(emp_id, False) else 'Ne'
        })
    spoc_df = pd.DataFrame(rows)
    return spoc_df

def nacti_nebo_generuj_df():
    vstup = 'Org_struktura_vypis.xlsx'
    if os.path.exists(vstup):
        df = pd.read_excel(vstup)
        cols = [
            'ID_zaměstnance','Jméno_zaměstnance','ID_organizační_jednotka',
            'Název_organizační_jednotka','ID_profese','Název_profese','úvazek',
            'Nadřízený_ID','Jméno_nadřízený','Společnost'
        ]
        df = df.rename(columns={c: cols[i] for i,c in enumerate(df.columns) if i < len(cols)})
    else:
        vytvor_testovaci_zamestnance(100, 'Testovaci_Struktura.xlsx')
        df = pd.read_excel('Testovaci_Struktura.xlsx')
    return df

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches

def vizualizuj_org_schema(df, spoc_dict, multicompany_dict, output_img_path):
    """
    Vykreslení manažerského top-down stromu pouze pro profese 1,2,3.
    - Boxy: jméno, profese, OJ, tým: X (všech přímých podřízených vč. specialista, pracovník), jméno nadřízeného, a případně MULTICOMPANY.
    - Barva: CEO a manažeři s SPOC >= 5 = zelená; SPOC < 3 = červená; ostatní modrá.
    - Multicompany: silné/dvojité orámování.
    - Spojnice pravoúhlé.
    - Velikost grafu se přizpůsobí nejširší vrstvě.
    - CEO je nahoře (osm Y osy).
    - Boxy vyšší a text po řádcích pro čitelnost.
    """
    # Zůstanou pouze manažeři
    managers = df[df['ID_profese'].isin([1,2,3])].copy()
    # CEO = kdokoli s profese==1, ber první
    ceo_rows = managers[managers['ID_profese']==1]
    if ceo_rows.empty:
        raise Exception("CEO (profese ID 1) nebyl nalezen")
    ceo_id = ceo_rows.iloc[0]['ID_zaměstnance']

    # BFS pro vrstvy
    level_map = {}
    parent_map = {}
    children_map = {}
    level_map[ceo_id] = 0
    parent_map[ceo_id] = None
    children_map[ceo_id] = []
    queue = [(ceo_id, 0)]
    while queue:
        current_id, lvl = queue.pop(0)
        pod = managers[managers['Nadřízený_ID'] == current_id]['ID_zaměstnance'].tolist()
        for child_id in pod:
            level_map[child_id] = lvl+1
            parent_map[child_id] = current_id
            if current_id not in children_map:
                children_map[current_id] = []
            children_map[current_id].append(child_id)
            children_map[child_id] = []
            queue.append((child_id, lvl+1))

    managers_ids = set(level_map.keys())
    managers = managers[managers['ID_zaměstnance'].isin(managers_ids)].copy()

    # Patra podle levelu
    layers = {}
    for id_, lvl in level_map.items():
        if lvl not in layers:
            layers[lvl] = []
        layers[lvl].append(id_)
    n_layers = max(layers.keys())+1
    max_width = max(len(ids) for ids in layers.values())

    y_gap = 2.5
    x_gap = 3.6
    positions = {}
    for lvl in range(n_layers):
        count = len(layers.get(lvl, []))
        start_x = -x_gap * (count-1)/2 if count > 1 else 0
        for idx, id_ in enumerate(layers.get(lvl, [])):
            x = start_x + idx * x_gap
            y = lvl * y_gap
            positions[id_] = (x, y)

    node_width = 3.1
    node_height = 1.6

    green = "#a6f5b6"    # SPOC >= 5
    red = "#d24545"      # SPOC < 3
    blue = "#b4c4ee"     # Ostatní
    border_normal = '#333'
    border_multicomp = '#390'
    borderwidth_normal = 2
    borderwidth_multi = 4

    def get_box_color(id_, spoc_val):
        if id_ in spoc_dict:
            if spoc_val >= 5:
                return green
            elif spoc_val < 3:
                return red
            else:
                return blue
        else:
            return blue

    def draw_orthogonal(ax, start, end, node_height):
        x0, y0 = start
        x1, y1 = end
        y_bottom = y0 + node_height/2
        y_top = y1 - node_height/2
        mid_y = (y_bottom + y_top) / 2
        ax.plot([x0, x0], [y_bottom, mid_y], color="#385380", linewidth=1.8, zorder=1, solid_capstyle='round')
        ax.plot([x0, x1], [mid_y, mid_y], color="#385380", linewidth=1.8, zorder=1, solid_capstyle='round')
        ax.plot([x1, x1], [mid_y, y_top], color="#385380", linewidth=1.8, zorder=1, solid_capstyle='round')

    plt.figure(figsize=(max(15, max_width * 3.4), n_layers * 2.6 + 1))
    ax = plt.gca()
    ax.axis('off')

    # Do boxu u každého manažera počítáme počet všech přímých podřízených (včetně specialistů a pracovníků!)
    # Pro zobrazení jména nadřízeného
    nadrizene_jmeno_map = {}
    for midx, mrow in managers.iterrows():
        id_ = mrow['ID_zaměstnance']
        nadrizeny_id = mrow['Nadřízený_ID']
        if pd.isna(nadrizeny_id):
            nadrizene_jmeno_map[id_] = ""
        else:
            nrow = df[df['ID_zaměstnance'] == nadrizeny_id]
            if len(nrow) > 0:
                nadrizene_jmeno_map[id_] = nrow.iloc[0]['Jméno_zaměstnance']
            else:
                nadrizene_jmeno_map[id_] = ""

    for id_, (x, y) in positions.items():
        mrow = managers[managers['ID_zaměstnance']==id_].iloc[0]
        spoc = spoc_dict.get(id_, 0)
        jmeno = mrow['Jméno_zaměstnance']
        profese = mrow['Název_profese']
        org_jedn = mrow['Název_organizační_jednotka']
        profese_id = mrow['ID_profese']
        nadrizeny_jmeno = nadrizene_jmeno_map.get(id_, "")
        multicomp = multicompany_dict.get(id_, False)

        box_color = get_box_color(id_, spoc)
        edgecol = border_multicomp if multicomp else border_normal
        linewidth = borderwidth_multi if multicomp else borderwidth_normal

        multistr = "\n[MULTICOMPANY]" if multicomp else ""
        box_label = (
            f"{jmeno}{multistr}\n"
            f"{profese}\n"
            f"{org_jedn}\n"
            f"Tým: {spoc}\n"
            f"Nadřízený: {nadrizeny_jmeno}"
        )

        rect = mpatches.FancyBboxPatch(
            (x-node_width/2, y-node_height/2),
            node_width, node_height,
            boxstyle="round,pad=0.11",
            linewidth=linewidth,
            edgecolor=edgecol,
            facecolor=box_color,
            zorder=2,
            mutation_scale=0.115
        )
        ax.add_patch(rect)
        ax.text(
            x, y, box_label, ha='center', va='center',
            fontsize=9.8, family="monospace", linespacing=1.27, zorder=3, fontweight='regular'
        )

    for child_id, parent_id in parent_map.items():
        if parent_id is not None and parent_id in positions and child_id in positions:
            px, py = positions[parent_id]
            cx, cy = positions[child_id]
            draw_orthogonal(ax, (px, py), (cx, cy), node_height)

    all_x = [x for x, y in positions.values()]
    all_y = [y for x, y in positions.values()]
    plt.ylim(min(all_y)-node_height*1.7, max(all_y)+node_height*1.6)
    plt.xlim(min(all_x)-node_width*1.18, max(all_x)+node_width*1.18)

    legend_handles = [
        mpatches.Patch(color=green, label="SPOC >= 5 (& CEO)"),
        mpatches.Patch(color=red, label="SPOC < 3 (málo)"),
        mpatches.Patch(color=blue, label="SPOC 3-4"),
        mpatches.Patch(edgecolor=border_multicomp, facecolor="none", linewidth=borderwidth_multi, label="Multicompany manažer")
    ]
    plt.legend(handles=legend_handles, loc="upper left", fontsize=9)
    plt.title("Organizační schéma – manažerská hierarchie (jen manažeři)", fontsize=14)
    plt.tight_layout()
    plt.savefig(output_img_path, dpi=220)
    plt.close()
    print(f"Organizační schéma uloženo na: {output_img_path}")

# PYVIS – interaktivní síť s hierarchickým layoutem (Up-Down) a custom tooltip/pro lazout (followup):
def vytvor_interaktivni_sit(df, spoc_dict, manager_companies_dict, output_html_path):
    from pyvis.network import Network

    net = Network(
        height="900px",
        width="95%",
        bgcolor="#ffffff",
        font_color="#000000",
        directed=True,
        notebook=False
    )

    # Hierarchical, up-down
    net.set_options('''
    var options = {
      "layout": {
        "hierarchical": {
          "enabled": true,
          "direction": "UD",
          "sortMethod": "hubsize",
          "levelSeparation": 180,
          "treeSpacing": 225,
          "nodeSpacing": 245,
          "blockShifting": true,
          "edgeMinimization": true,
          "parentCentralization": true
        }
      },
      "physics": {
        "enabled": false
      }
    }''')

    df_indexed = df.set_index("ID_zaměstnance")

    for emp_id, row in df_indexed.iterrows():
        jmeno = row["Jméno_zaměstnance"]
        profese = row["Název_profese"]
        profese_id = int(row["ID_profese"])

        spoc_val = int(spoc_dict.get(emp_id, 0))
        companies = manager_companies_dict.get(emp_id, [])
        is_multicomp = len(companies) > 1
        companies_str = ", ".join(companies) if companies else "-"

        # Barva a tvar; decentní, zvýraznit multicomp
        if profese_id in [1, 2, 3]:
            size = 20 + spoc_val * 3
            color = "#34D768" if spoc_val >= 5 else "#E74C3C" if spoc_val < 3 else "#4b7bec"
            shape = "diamond" if is_multicomp else "box"
            border_width = 6 if is_multicomp else 2
        else:
            size = 10
            color = "#a5b1c2"
            shape = "dot"
            border_width = 1

        multicomp_label = " [MULTICOMPANY]" if is_multicomp and profese_id in [1,2,3] else ""
        # Tooltip: jmeno, profese, SPOC, společnosti
        if profese_id in [1, 2, 3]:
            title = (
                f"<b>{jmeno}</b>{multicomp_label}<br>"
                f"Profese: {profese}<br>"
                f"Počet lidí v týmu (SPOC): {spoc_val}<br>"
                f"Společnosti v týmu: {companies_str}"
            )
        else:
            title = (
                f"<b>{jmeno}</b><br>Profese: {profese}"
            )

        # Add only management to hierarchy
        if profese_id in [1, 2, 3]:
            net.add_node(
                emp_id,
                label=jmeno + multicomp_label,
                title=title,
                size=size,
                color=color,
                shape=shape,
                borderWidth=border_width
            )
        # Ostatní (s profese 4,5) se nekreslí (do hierarchie dle zadání)

    # Vztahy manažerské
    for _, row in df[df['ID_profese'].isin([1,2,3])].iterrows():
        emp_id = row["ID_zaměstnance"]
        nadrizeny_id = row["Nadřízený_ID"]
        if pd.notna(nadrizeny_id):
            try:
                nadrizeny_id_int = int(nadrizeny_id)
            except ValueError:
                continue
            if nadrizeny_id_int in df_indexed.index:
                if int(df_indexed.loc[nadrizeny_id_int]['ID_profese']) in [1,2,3]:
                    net.add_edge(nadrizeny_id_int, emp_id)

    # Uložit výsledek jako HTML
    net.write_html(output_html_path)
    print(f"Interaktivní síťové schéma uloženo na: {output_html_path}")

import random
import os
from pyvis.network import Network

# Cesty a názvy souborů
LOCAL_STRUCT = "Org_struktura_vypis.xlsx"

# Zjistíme, zda existuje struktura z reálných dat
if os.path.exists(LOCAL_STRUCT):
    print("--- Detekován existující soubor s organizační strukturou. Načítám data z něj... ---")
    df = pd.read_excel(LOCAL_STRUCT)
    # Přehodíme názvy sloupců, pokud je nutné
    col_map = {
        'ID_zaměstnance': 'ID_zaměstnance',
        'Jméno_zaměstnance': 'Jméno_zaměstnance',
        'ID_organizační_jednotka': 'ID_organizační_jednotka',
        'Název_organizační_jednotka': 'Název_organizační_jednotka',
        'ID_profese': 'ID_profese',
        'Název_profese': 'Název_profese',
        'úvazek': 'úvazek',
        'Nadřízený_ID': 'Nadřízený_ID',
        'Jméno_nadřízený': 'Jméno_nadřízený',
        'Společnost': 'Společnost'
    }
    # Přejmenujeme pouze, pokud někdo z lidí použil jiné - robustní mapování
    columns_lower = [c.lower() for c in df.columns]
    true_map = {}
    for stdcol in col_map:
        # Najít v ignoraci case/unicode
        found = None
        for c in df.columns:
            if c.strip().lower() == stdcol.strip().lower():
                found = c
                break
        if found:
            true_map[found] = stdcol
    df = df.rename(columns=true_map)
else:
    print("--- Nenašel jsem Org_struktura_vypis.xlsx, generuji testovací data ---")
    # --- Data Generation ---
    jmena = [
        "Jan", "Petr", "Lucie", "Eva", "Martin", "Tereza", "Jakub", "Barbora", "Michal", "Katerina",
        "Ondrej", "Alena", "Roman", "Hana", "Viktor", "Jana", "Pavel", "Lenka", "Daniel", "Marketa"
    ]
    prijmeni = [
        "Novak", "Svoboda", "Dvorak", "Cerny", "Prochazka", "Kucera", "Vesely", "Horak", "Nemec", "Marek",
        "Pospisil", "Hajek", "Kratochvil", "Jelinek", "Ruzicka", "Fiala", "Sedlak", "Urban", "Blaha", "Kolar"
    ]
    organizacni_jednotky = [
        ("100", "IT"), ("200", "HR"), ("300", "Finance"),
        ("400", "Sales"), ("500", "Marketing")
    ]
    profese = [("1", "CEO"), ("2", "Manažer"), ("3", "Vedoucí týmu"), ("4", "Specialista"), ("5", "Pracovník")]
    spolecnosti = ["AlphaCorp", "BetaCorp", "GammaLtd"]

    zamestnanci = []

    # CEO - Level 0
    idzam = 1
    ceo_jmeno = f"{random.choice(jmena)} {random.choice(prijmeni)}"
    ceo_spolecnost = random.choice(spolecnosti)
    zamestnanci.append({
        "ID_zaměstnance": idzam,
        "Jméno_zaměstnance": ceo_jmeno,
        "ID_organizační_jednotka": "0",
        "Název_organizační_jednotka": "Centrála",
        "ID_profese": 1,
        "Název_profese": "CEO",
        "úvazek": 1,
        "Nadřízený_ID": None,
        "Jméno_nadřízený": "",
        "Společnost": ceo_spolecnost
    })
    idzam += 1

    # Manažeři - Level 1
    manazeri = []
    for i in range(5):
        jmeno = f"{random.choice(jmena)} {random.choice(prijmeni)}"
        spol = random.choice(spolecnosti)
        org_id, org_nazev = random.choice(organizacni_jednotky)
        zamestnanci.append({
            "ID_zaměstnance": idzam,
            "Jméno_zaměstnance": jmeno,
            "ID_organizační_jednotka": org_id,
            "Název_organizační_jednotka": org_nazev,
            "ID_profese": 2,
            "Název_profese": "Manažer",
            "úvazek": 1,
            "Nadřízený_ID": 1,  # CEO id
            "Jméno_nadřízený": ceo_jmeno,
            "Společnost": spol
        })
        manazeri.append({
            "id": idzam,
            "jmeno": jmeno,
            "spolecnost": spol
        })
        idzam += 1

    # Vedoucí týmů - Level 2
    vedouci = []
    for m in manazeri:
        ved_tymu_count = random.randint(2, 4)
        for _ in range(ved_tymu_count):
            jmeno = f"{random.choice(jmena)} {random.choice(prijmeni)}"
            spol = m["spolecnost"]  # Udržet stejnou společnost
            org_id, org_nazev = random.choice(organizacni_jednotky)
            zamestnanci.append({
                "ID_zaměstnance": idzam,
                "Jméno_zaměstnance": jmeno,
                "ID_organizační_jednotka": org_id,
                "Název_organizační_jednotka": org_nazev,
                "ID_profese": 3,
                "Název_profese": "Vedoucí týmu",
                "úvazek": 1,
                "Nadřízený_ID": m["id"],
                "Jméno_nadřízený": m["jmeno"],
                "Společnost": spol
            })
            vedouci.append({
                "id": idzam,
                "jmeno": jmeno,
                "spolecnost": spol,
                "nad_id": m["id"],
                "jmeno_nadrizeny": m["jmeno"]
            })
            idzam += 1

    # Specialisté a pracovníci - Level 3
    for v in vedouci:
        pocet_podrizenych = random.randint(3, 6)
        for _ in range(pocet_podrizenych):
            typ = random.choices([("4", "Specialista"), ("5", "Pracovník")], weights=[0.5, 0.5])[0]
            jmeno = f"{random.choice(jmena)} {random.choice(prijmeni)}"
            org_id, org_nazev = random.choice(organizacni_jednotky)
            uvazek = random.choice([1.0, 0.8, 0.5])
            zamestnanci.append({
                "ID_zaměstnance": idzam,
                "Jméno_zaměstnance": jmeno,
                "ID_organizační_jednotka": org_id,
                "Název_organizační_jednotka": org_nazev,
                "ID_profese": int(typ[0]),
                "Název_profese": typ[1],
                "úvazek": uvazek,
                "Nadřízený_ID": v["id"],
                "Jméno_nadřízený": v["jmeno"],
                "Společnost": v["spolecnost"]
            })
            idzam += 1

    # Omez na přesně 100 lidí (testovací případ)
    if len(zamestnanci) > 100:
        zamestnanci = zamestnanci[:100]

    df = pd.DataFrame(zamestnanci)
    # Uložit pro další běh
    desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    output_excel = os.path.join(desktop_path, "Testovaci_Struktura.xlsx")
    df.to_excel(output_excel, index=False)
    print(f"--- HOTOVO! Testovací soubor najdeš na ploše pod názvem: {output_excel} ---")

# Pokud dataframe nebyl uložen na ploše, doplníme, kde bude vizualizace ukládána
desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
output_img = os.path.join(desktop_path, "Organizacni_Schema.png")

# ----------------------------------
# SPOČÍTÁNÍ SPOC pro manažery + Multicompany
# ----------------------------------
# Manažer = profese 1,2,3; SPOC = kolik má podřízených specialista/pracovník (TÝM)
id_manazeru = df[df['ID_profese'].isin([1,2,3])]['ID_zaměstnance'].tolist()
spoc_dict = {}
multicompany_dict = {}       # True/False – zda má tým z více firem
company_count_dict = {}      # Počet unikátních společností v týmu
manager_companies_dict = {}  # Mapování ID manažera -> seřazený seznam společností

for idm in id_manazeru:
    # Přímí podřízení specialisté a pracovníci (SPOC)
    spoc = len(df[(df['Nadřízený_ID'] == idm)
                  & (df['ID_profese'].isin([4, 5]))])
    spoc_dict[idm] = spoc

    # Multicompany analytika: kolik různých společností mají jeho PŘÍMÍ podřízení (všech profesí)
    podrizeni = df[df['Nadřízený_ID'] == idm]
    spol_set = sorted(set(podrizeni['Společnost'].dropna()))
    pocet_firem = len(spol_set)
    company_count_dict[idm] = pocet_firem
    manager_companies_dict[idm] = spol_set
    multicomp = pocet_firem > 1
    multicompany_dict[idm] = multicomp

spoc_df = pd.DataFrame([{
    "ID_manažera": idm,
    "Jméno_manažera": df[df['ID_zaměstnance']==idm]['Jméno_zaměstnance'].values[0],
    "Společnost": df[df['ID_zaměstnance']==idm]['Společnost'].values[0],
    "SPOC": spoc_dict[idm],
    "Multicompany": "Ano" if multicompany_dict[idm] else "Ne",
    "Počet_firem": company_count_dict[idm]
} for idm in id_manazeru])
print("SPOC pro manažery (+ Multicompany):")
print(spoc_df)

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches

def vizualizuj_org_schema(df, spoc_dict, output_img_path, multicompany_dict=None):
    """
    Vykreslí statické organizační schéma managementu do PNG souboru.

    Vstup:
        df (pandas.DataFrame): kompletní organizační data.
        spoc_dict (dict[int, int]): SPOC – počet přímých podřízených pro manažery.
        output_img_path (str): cílová cesta k výstupnímu PNG souboru.
        multicompany_dict (dict[int, bool] | None): volitelné mapování
            ID manažera -> True/False, zda je Multicompany (ovlivňuje rámeček a popisek).

    Logika:
        - Filtrování pouze na profese 1–3 (CEO, Manažer, Vedoucí týmu).
        - Sestavení stromu managementu (BFS) a rozložení uzlů do pater shora dolů.
        - Pro každý box se zobrazí: jméno, profese, OJ, velikost týmu (všichni přímí podřízení)
          a jméno nadřízeného, případně tag [MULTICOMPANY].
        - Barvy boxů odpovídají vytížení (SPOC) a CEO je vždy zelený.
        - Multicompany manažeři mají silnější rámeček odlišné barvy.

    Výstup:
        PNG obrázek uložený na `output_img_path` s top‑down stromem manažerské hierarchie.
    """
    if multicompany_dict is None:
        multicompany_dict = {}

    # Filtrovat pouze management (profese 1,2,3)
    managers = df[df['ID_profese'].isin([1,2,3])].copy()

    # CEO (profese==1), vezmem první výskyt
    ceo_rows = managers[managers['ID_profese']==1]
    if ceo_rows.empty:
        raise Exception("CEO (profese ID 1) nebyl nalezen")
    ceo_id = ceo_rows.iloc[0]['ID_zaměstnance']

    # Získat vrstvy stromu (BFS nad grafem managementu)
    level_map = {}
    parent_map = {}
    children_map = {}
    level_map[ceo_id] = 0
    parent_map[ceo_id] = None
    children_map[ceo_id] = []
    queue = [(ceo_id, 0)]
    while queue:
        current_id, lvl = queue.pop(0)
        pod = managers[managers['Nadřízený_ID'] == current_id]['ID_zaměstnance'].tolist()
        for child_id in pod:
            level_map[child_id] = lvl+1
            parent_map[child_id] = current_id
            if current_id not in children_map:
                children_map[current_id] = []
            children_map[current_id].append(child_id)
            children_map[child_id] = []
            queue.append((child_id, lvl+1))

    # Omezit managers jen na uzly, které jsou v management hierarchii (napočítané BFS)
    managers_ids = set(level_map.keys())
    managers = managers[managers['ID_zaměstnance'].isin(managers_ids)].copy()

    # Vrstvy pro zarovnání boxů (úroveň = level)
    layers = {}
    for id_, lvl in level_map.items():
        if lvl not in layers:
            layers[lvl] = []
        layers[lvl].append(id_)

    n_layers = max(layers.keys())+1
    max_width = max(len(ids) for ids in layers.values())

    # Pozice ve fígu – CEO naprosto nahoře, úroveň = osa Y (Y = 0 nejvýš)
    y_gap = 2.95
    x_gap = 4.35
    positions = {}
    for lvl in range(n_layers):
        count = len(layers.get(lvl, []))
        start_x = -x_gap * (count-1)/2 if count > 1 else 0
        for idx, id_ in enumerate(layers.get(lvl, [])):
            x = start_x + idx*x_gap
            y = -lvl * y_gap  # Záporně (CEO y=0, ostatní níž/dolů!)
            positions[id_] = (x, y)

    # Rozměry boxů (vyšší kvůli více textu, tučně - multicompany)
    node_width = 3.6
    node_height = 2.6

    # Barvy
    green = "#a6f5b6"
    red = "#d24545"
    blue = "#b4c4ee"

    def get_box_color(id_, spoc_val, profese_id):
        if int(profese_id) == 1:  # CEO vždy zelený
            return green
        if id_ in spoc_dict:
            if spoc_val >= 5:
                return green
            elif spoc_val < 3:
                return red
            else:
                return blue
        else:
            return blue

    def get_box_linewidth(id_):
        # Pokud je MULTICOMPANY, okraj bude tučnější/dvojitý
        return 4 if multicompany_dict.get(id_, False) else 2

    def draw_orthogonal(ax, start, end, node_height):
        # Pravoúhlá spojka: dolů, pak vodorovně, pak dolů
        x0, y0 = start
        x1, y1 = end
        y_bottom = y0 - node_height/2
        y_top = y1 + node_height/2
        mid_y = (y_bottom + y_top) / 2
        ax.plot([x0, x0], [y_bottom, mid_y], color="#385380", linewidth=1.7, zorder=1, solid_capstyle='round')
        ax.plot([x0, x1], [mid_y, mid_y], color="#385380", linewidth=1.7, zorder=1, solid_capstyle='round')
        ax.plot([x1, x1], [mid_y, y_top], color="#385380", linewidth=1.7, zorder=1, solid_capstyle='round')

    plt.figure(figsize=(max(13, max_width * 3.20), n_layers * 3.2 + 2.2))
    ax = plt.gca()
    ax.axis('off')

    # Boxy a text (vždy víceřádkové, čitelné)
    for id_, (x, y) in positions.items():
        row = managers[managers['ID_zaměstnance']==id_].iloc[0]
        profese = row['Název_profese']
        org_jedn = row['Název_organizační_jednotka']
        jmeno = row['Jméno_zaměstnance']
        profese_id = row['ID_profese']
        nadrizeny_jmeno = row['Jméno_nadřízený']
        # Počet všech přímých podřízených (všech profesí!)
        team_count = len(df[(df['Nadřízený_ID']==id_)])
        # Multicompany?
        is_mc = multicompany_dict.get(id_, False)
        barva = get_box_color(id_, spoc_dict.get(id_, 0), profese_id)
        linewidth_box = get_box_linewidth(id_)

        # Box s tučnějším okrajem pro multicompany
        rect = mpatches.FancyBboxPatch(
            (x-node_width/2, y-node_height/2),
            node_width, node_height,
            boxstyle="round,pad=0.13",
            linewidth=linewidth_box,
            edgecolor='#f58500' if is_mc else '#333',
            facecolor=barva,
            zorder=2,
            mutation_scale=0.120
        )
        ax.add_patch(rect)
        # Texty
        mc_str = "\n[MULTICOMPANY]" if is_mc else ""
        box_label = (
            f"{jmeno}{mc_str}\n"
            f"{profese}\n"
            f"{org_jedn}\n"
            f"Tým: {team_count}\n"
            f"Nadřízený: {nadrizeny_jmeno if nadrizeny_jmeno else '-'}"
        )
        ax.text(
            x, y, box_label, ha='center', va='center',
            fontsize=11.3, family="monospace", linespacing=1.28, zorder=3, fontweight="bold" if is_mc else "normal"
        )

    # Spojnice (pravouhlé) pouze mezi manažery
    for child_id, parent_id in parent_map.items():
        if parent_id is not None and parent_id in positions and child_id in positions:
            px, py = positions[parent_id]
            cx, cy = positions[child_id]
            draw_orthogonal(ax, (px, py), (cx, cy), node_height)

    # Osa: CEO nahoře
    all_x = [x for x, y in positions.values()]
    all_y = [y for x, y in positions.values()]
    plt.ylim(min(all_y)-node_height*1.3, max(all_y)+node_height*1.3)
    plt.xlim(min(all_x)-node_width*1.13, max(all_x)+node_width*1.13)

    # Legenda
    legend_handles = [
        mpatches.Patch(color=green, label="SPOC >= 5 (včetně CEO)"),
        mpatches.Patch(color=red, label="SPOC < 3"),
        mpatches.Patch(color=blue, label="SPOC 3-4"),
        mpatches.Patch(edgecolor='#f58500', facecolor="w", label="Manažer MULTICOMPANY", linewidth=3)
    ]
    plt.legend(handles=legend_handles, loc="upper left", fontsize=9)
    plt.title("Organizační schéma – manažerská hierarchie (jen manažeři)", fontsize=14)
    plt.tight_layout()
    plt.savefig(output_img_path, dpi=220)
    plt.close()
    print(f"Organizační schéma uloženo na: {output_img_path}")


def vytvor_interaktivni_sit(df, spoc_dict, manager_companies_dict, output_html_path):
    """
    Vytvoří interaktivní manažerské schéma v HTML pomocí knihovny Pyvis.

    Vstup:
        df (pandas.DataFrame): kompletní organizační data.
        spoc_dict (dict[int, int]): SPOC hodnoty pro manažery (použito pro barvení a velikost).
        manager_companies_dict (dict[int, list[str]]): ID manažera -> seznam společností v týmu,
            slouží k detekci Multicompany manažerů.
        output_html_path (str): cílová cesta k výstupnímu HTML souboru.

    Logika:
        - Do sítě zahrnuje pouze manažerské profese (1–3) a skládá je do hierarchického layoutu
          shora dolů (`UD`), s vypnutou fyzikou pro stabilní rozložení.
        - Každý uzel je obdélníkový box (`shape="box"`) s textem: jméno / OJ / profese.
        - Velikost uzlu a barva odpovídají velikosti týmu (SPOC); barvy kopírují schéma PNG.
        - Multicompany manažeři mají výrazně silnější rámeček a v tooltipu seznam společností.
        - Spojení (edges) se vykreslují pouze mezi manažery podle vztahu nadřízený–podřízený.

    Výstup:
        HTML soubor s interaktivní sítí uložený na `output_html_path`, vhodný pro otevření v prohlížeči.
    """
    # Inicializace sítě – fyzikální rozložení, lepší čitelnost
    net = Network(
        height="800px",
        width="100%",
        bgcolor="#ffffff",
        font_color="#000000",
        directed=True,
        notebook=False
    )

    # Hierarchické uspořádání shora dolů, fixní patra (vypnutá fyzika)
    net.set_options("""
    var options = {
      "layout": {
        "hierarchical": {
          "enabled": true,
          "direction": "UD",
          "sortMethod": "hubsize",
          "nodeSpacing": 150,
          "levelSeparation": 400,
          "edgeMinimization": false,
          "blockShifting": true
        }
      },
      "physics": {
        "enabled": false
      }
    }
    """)

    # Pracujeme pouze s manažery (profese 1,2,3)
    managers_df = df[df["ID_profese"].isin([1, 2, 3])].copy()

    # Vypočítáme úrovně (level) v manažerském stromu s "pod-levly"
    # tak, aby manažeři s mnoha podřízenými měli děti rozhozené do více vertikálních pater.
    level_map = {}
    children_map = {}
    parent_map = {}

    # Sestavení map dětí/rodičů v rámci manažerů
    manager_ids = set(managers_df["ID_zaměstnance"].tolist())
    for _, row in managers_df.iterrows():
        mid = row["ID_zaměstnance"]
        parent_id = row["Nadřízený_ID"]
        if mid not in children_map:
            children_map[mid] = []
        if pd.notna(parent_id) and int(parent_id) in manager_ids:
            parent_id = int(parent_id)
            parent_map[mid] = parent_id
            if parent_id not in children_map:
                children_map[parent_id] = []
            children_map[parent_id].append(mid)
        else:
            parent_map[mid] = None

    # Kořeny (typicky CEO) – ti bez manažerského nadřízeného
    roots = [mid for mid, pid in parent_map.items() if pid is None]
    for r in roots:
        level_map[r] = 0

    # BFS s pod-levly: pokud má manažer >5 přímých manažerských podřízených,
    # rozdělíme je na dvě patra (L+1 a L+2), aby nebyli v jedné dlouhé řadě.
    from collections import deque
    queue = deque([(r, 0) for r in roots])
    while queue:
        current_id, cur_level = queue.popleft()
        children = children_map.get(current_id, [])
        if not children:
            continue
        # Stabilní pořadí
        children = sorted(children)
        if len(children) <= 5:
            # všichni děti v jednom patře
            for ch in children:
                if ch not in level_map or level_map[ch] > cur_level + 1:
                    level_map[ch] = cur_level + 1
                queue.append((ch, level_map[ch]))
        else:
            # rozdělit do dvou pater: prvních 5 na L+1, zbytek na L+2
            first_row = children[:5]
            second_row = children[5:]
            for ch in first_row:
                if ch not in level_map or level_map[ch] > cur_level + 1:
                    level_map[ch] = cur_level + 1
                queue.append((ch, level_map[ch]))
            for ch in second_row:
                if ch not in level_map or level_map[ch] > cur_level + 2:
                    level_map[ch] = cur_level + 2
                queue.append((ch, level_map[ch]))

    # Pokud některý manažer nemá přiřazený level (bezpečnostní pojistka), dej mu 0
    for mid in managers_df["ID_zaměstnance"]:
        if mid not in level_map:
            level_map[mid] = 0

    # Pro rychlý přístup k řádkům podle ID (jen manažeři)
    df_indexed = managers_df.set_index("ID_zaměstnance")

    # Vytvořit uzly – jen manažeři
    for emp_id, row in df_indexed.iterrows():
        jmeno = row["Jméno_zaměstnance"]
        profese = row["Název_profese"]
        org_jedn = row["Název_organizační_jednotka"]
        profese_id = int(row["ID_profese"])

        # SPOC (velikost týmu): celkový počet přímých podřízených (všech profesí včetně skrytých)
        team_size = len(df[df["Nadřízený_ID"] == emp_id])
        spoc_val = team_size

        # Seznam společností v týmu daného manažera (podle přímých podřízených)
        companies = manager_companies_dict.get(emp_id, [])
        companies_str = ", ".join(companies) if companies else "-"
        is_multicomp = len(companies) > 1

        # Základní velikost a tvar uzlu – pouze manažeři
        # Manažeři – škálování velikosti podle SPOC (minimálně 15)
        size = 15 + spoc_val * 3

        # Barvy dle PNG schématu
        green = "#a6f5b6"  # SPOC >= 5
        red = "#d24545"    # SPOC < 3
        blue = "#b4c4ee"   # ostatní

        if profese_id == 1:
            fill_color = green
        elif spoc_val >= 5:
            fill_color = green
        elif spoc_val < 3:
            fill_color = red
        else:
            fill_color = blue

        border_color = "#f58500" if is_multicomp else "#333333"
        border_width = 6 if is_multicomp else 2

        # Tvar uzlu – obdélník jako v PNG
        shape = "box"

        # Label přímo v boxu: Jméno / OJ / Profese
        label = f"{jmeno}\n{org_jedn}\n{profese}"

        # Text v tooltipu po najetí myší
        if is_multicomp:
            title = (
                f"<b>{jmeno}</b> [MULTICOMPANY]<br>"
                f"Profese: {profese}<br>"
                f"SPOC (přímí podřízení celkem): {spoc_val}<br>"
                f"Společnosti v týmu: {companies_str}"
            )
        else:
            title = (
                f"<b>{jmeno}</b><br>"
                f"Profese: {profese}<br>"
                f"SPOC (přímí podřízení celkem): {spoc_val}"
            )

        net.add_node(
            emp_id,
            label=label,
            title=title,
            size=size,
            shape=shape,
            color={
                "background": fill_color,
                "border": border_color
            },
            borderWidth=border_width,
            level=level_map.get(emp_id, 0)
        )

    # Vytvořit hrany pouze mezi manažery
    for _, row in managers_df.iterrows():
        emp_id = row["ID_zaměstnance"]
        nadrizeny_id = row["Nadřízený_ID"]
        if pd.notna(nadrizeny_id):
            try:
                nadrizeny_id_int = int(nadrizeny_id)
            except ValueError:
                continue
            # Hrana od nadřízeného manažera k podřízenému manažerovi
            if nadrizeny_id_int in df_indexed.index:
                net.add_edge(nadrizeny_id_int, emp_id)

    # Uložit výsledek jako HTML – pouze write_html (žádné show/save_graph)
    net.write_html(output_html_path)
    print(f"Interaktivní síťové schéma uloženo na: {output_html_path}")

# Pokud starý obrázek existuje, smažeme ho (pojistka)
if os.path.exists(output_img):
    os.remove(output_img)

# Vyvoláme vizualizaci
vizualizuj_org_schema(df, spoc_dict, output_img, multicompany_dict)
print(f"Byl vytvořen zcela nový obrázek zde: {output_img}")

# Vytvoření interaktivního síťového schématu na plochu
output_html = os.path.join(desktop_path, "interaktivni_schema.html")
vytvor_interaktivni_sit(df, spoc_dict, manager_companies_dict, output_html)