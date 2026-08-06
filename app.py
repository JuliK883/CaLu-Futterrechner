import streamlit as st
import pandas as pd
from datetime import date
import plotly.express as px
import json
import os
import tempfile
import plotly.io as pio
pio.kaleido.scope.chromium_args = (
    "--headless",
    "--no-sandbox",
    "--disable-gpu",
)

# Versuch, FPDF für den PDF-Export zu laden
try:
    from fpdf import FPDF
    HAS_FPDF = True
except ImportError:
    HAS_FPDF = False

# --- SEITEN-KONFIGURATION ---
# Das MUSS der erste Streamlit-Befehl sein
st.set_page_config(page_title="CaLu Futterrechner", layout="wide")


# ==========================================
# 1. LOGIN-BEREICH
# ==========================================
if 'eingeloggt' not in st.session_state:
    st.session_state['eingeloggt'] = False

# Nur anzeigen, wenn man NICHT eingeloggt ist
if not st.session_state['eingeloggt']:
    st.title("🔒 Login")
    
    try:
        # Lädt die Benutzerverwaltung
        df_logins = pd.read_excel("Futterrechner 12.07.26.xlsx", sheet_name="Benutzerverwaltung")
    except Exception as e:
        st.error(f"Konnte die Zugangsdaten nicht laden. Fehler: {e}")
        st.stop()

   # Eingabefelder für den Nutzer
eingabe_name = st.text_input("Benutzername")
eingabe_passwort = st.text_input("Passwort", type="password")

if st.button("Einloggen"):
    if eingabe_name in df_logins['Benutzername'].values:
        user_row = df_logins[df_logins['Benutzername'] == eingabe_name].iloc[0]
        
        if str(user_row['Passwort Hash']) == eingabe_passwort and str(user_row['Zugang aktiv']).strip().lower() == 'ja':
            st.session_state['eingeloggt'] = True
            st.session_state['aktiver_nutzer'] = eingabe_name
            st.rerun() 
        else:
            st.error("Passwort falsch oder Zugang deaktiviert.")
    else:
        st.error("Benutzername nicht gefunden.")

# Stoppt das Skript hier, wenn der Login noch nicht erfolgreich war
st.stop()
    
# ==========================================
# 2. HAUPT-APP (CaLu Futterrechner)
# ==========================================
st.title("🐾 CaLu Futterrechner 🐾")

# --- DATENBANK FUNKTIONEN ---
def get_db_file():
    nutzer = st.session_state.get('aktiver_nutzer', 'gast')
    nutzer_clean = "".join(c for c in nutzer if c.isalnum())
    return f"calu_profile_{nutzer_clean}.json"

def lade_db():
    db_file = get_db_file()
    if os.path.exists(db_file):
        with open(db_file, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def speichere_db(daten):
    db_file = get_db_file()
    with open(db_file, "w", encoding="utf-8") as f:
        json.dump(daten, f, indent=4)

# Pfad zur Excel-Datei
excel_pfad = "Futterrechner 12.07.26.xlsx"

# EXCEL-DATEN LADEN
@st.cache_data
def lade_daten():
    df_h = pd.read_excel(excel_pfad, sheet_name="Hintergrunddaten", header=2)
    df_l = pd.read_excel(excel_pfad, sheet_name="Lebensmitteldaten", header=1)
    df_l = df_l.fillna(0)
    return df_h, df_l

try:
    df_h, df_l = lade_daten()
    liste_zutaten = df_l.iloc[:, 2].dropna().tolist()
    liste_zutaten = list(dict.fromkeys(liste_zutaten))
    
    auswahl_alter = df_h["Multiplikator Energie nach Alter"].dropna().tolist()
    auswahl_verdaulichkeit = df_h.iloc[:, 5].dropna().tolist()
    naehrstoff_spalten = df_l.columns[3:].tolist()
    
    zutaten_col = df_l.columns[2]
    df_l_clean = df_l[df_l[zutaten_col] != 0] 
    df_l_clean = df_l_clean.drop_duplicates(subset=[zutaten_col], keep='first')
    zutaten_lookup = df_l_clean.set_index(zutaten_col).to_dict('index')

except Exception as e:
    st.error(f"Fehler beim Laden der Excel-Datei: {e}")
    st.stop()

# --- SESSION STATE INITIALISIERUNG ---
if 'hund_name' not in st.session_state: st.session_state.hund_name = ""
if 'geb_datum' not in st.session_state: st.session_state.geb_datum = date.today()
if 'altersgruppe' not in st.session_state: 
    st.session_state.altersgruppe = auswahl_alter[0] if auswahl_alter else "Adult"
if 'energiebedarf' not in st.session_state: st.session_state.energiebedarf = "Normal"
if 'gewicht' not in st.session_state: st.session_state.gewicht = 1.0
if 'zielgewicht' not in st.session_state: st.session_state.zielgewicht = 1.0
if 'verdaulichkeit' not in st.session_state: 
    st.session_state.verdaulichkeit = auswahl_verdaulichkeit[0] if auswahl_verdaulichkeit else "normal"

if 'futter_data' not in st.session_state: st.session_state.futter_data = []
if 'aktiver_plan_name' not in st.session_state: st.session_state.aktiver_plan_name = "Neuer Plan"
if 'selectbox_key' not in st.session_state: st.session_state.selectbox_key = 0
if 'angezeigte_spalten' not in st.session_state: st.session_state.angezeigte_spalten = naehrstoff_spalten.copy()

def sync_gewicht():
    st.session_state.zielgewicht = st.session_state.gewicht

def update_altersgruppe():
    alter_jahre = (date.today() - st.session_state.geb_datum).days / 365.25
    if len(auswahl_alter) > 2:
        if alter_jahre < 1: st.session_state.altersgruppe = auswahl_alter[0]
        elif alter_jahre < 8: st.session_state.altersgruppe = auswahl_alter[1]
        else: st.session_state.altersgruppe = auswahl_alter[2]

def reset_profil():
    st.session_state.hund_name = ""
    st.session_state.geb_datum = date.today()
    if auswahl_alter: st.session_state.altersgruppe = auswahl_alter[0]
    st.session_state.energiebedarf = "Normal"
    st.session_state.gewicht = 1.0  
    st.session_state.zielgewicht = 1.0
    if auswahl_verdaulichkeit: st.session_state.verdaulichkeit = auswahl_verdaulichkeit[0]
    st.session_state.futter_data = []
    st.session_state.aktiver_plan_name = "Neuer Plan"
    st.session_state.angezeigte_spalten = naehrstoff_spalten.copy()
    
    # Checkboxen zurücksetzen
    for col in naehrstoff_spalten:
        st.session_state[f"chk_{col}"] = True

def formatiere_zahl(zahl):
    if zahl == int(zahl): return f"{int(zahl):,}".replace(",", ".")
    return f"{zahl:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def clean_text(text):
    return str(text).encode('latin-1', 'replace').decode('latin-1')

db = lade_db()
hunde_namen = list(db.keys())

# --- TABS ERSTELLEN ---
tab1, tab2, tab3 = st.tabs(["📝 Hundedaten", "📊 Futterplan & Analyse", "🧮 Tools & Umrechner"])

# ==========================================
# TAB 1: HUNDEDATEN
# ==========================================
with tab1:
    st.markdown("### 📂 Profil laden")
    col_load1, col_load2 = st.columns([3, 1])
    with col_load1:
        auswahl_profil = st.selectbox("Gespeichertes Profil auswählen:", ["-- Neues Profil / Keines --"] + hunde_namen, label_visibility="collapsed")
    with col_load2:
        if st.button("Laden", use_container_width=True):
            if auswahl_profil != "-- Neues Profil / Keines --":
                prof = db[auswahl_profil]
                st.session_state.hund_name = auswahl_profil
                st.session_state.geb_datum = date.fromisoformat(prof["geb_datum"])
                st.session_state.altersgruppe = prof["altersgruppe"]
                st.session_state.energiebedarf = prof["energiebedarf"]
                st.session_state.gewicht = prof["gewicht"]
                st.session_state.zielgewicht = prof["zielgewicht"]
                st.session_state.verdaulichkeit = prof["verdaulichkeit"]
                st.session_state.futter_data = [] 
                st.session_state.aktiver_plan_name = "Neuer Plan"
                
                # Nährstoffauswahl laden und Checkboxen syncen
                geladene_spalten = prof.get("angezeigte_spalten", naehrstoff_spalten.copy())
                st.session_state.angezeigte_spalten = geladene_spalten
                for col in naehrstoff_spalten:
                    st.session_state[f"chk_{col}"] = col in geladene_spalten
            else:
                reset_profil()
            st.rerun()

    st.divider()

    st.header("Daten des Hundes")
    col_links, col_rechts = st.columns(2)
    
    with col_links:
        st.text_input("Name des Hundes", key="hund_name")
        st.date_input("Geburtsdatum", key="geb_datum", on_change=update_altersgruppe)
        st.selectbox("Altersgruppe", auswahl_alter, key="altersgruppe")
        st.selectbox("Energiebedarf", ["Normal", "kastriert (- 15%)", "kastriert (- 20%)", "Übergewicht (- 30%)"], key="energiebedarf")
        
    with col_rechts:
        st.number_input("Gewicht (kg)", step=0.1, key="gewicht", on_change=sync_gewicht)
        st.number_input("Zielgewicht (kg)", step=0.1, key="zielgewicht")
        stoffwechselgewicht = round(st.session_state.gewicht ** 0.75, 2) if st.session_state.gewicht > 0 else 0
        st.text_input("Stoffwechselgewicht", value=f"{stoffwechselgewicht} kg", disabled=True)
        st.selectbox("Verdaulichkeit Protein", auswahl_verdaulichkeit, key="verdaulichkeit")

    st.write("")
    if st.button("💾 Aktuelle Hundedaten als Profil speichern", type="primary"):
        name = st.session_state.hund_name
        if not name.strip():
            st.error("Bitte gib dem Hund einen Namen, um das Profil zu speichern.")
        else:
            if name not in db:
                db[name] = {"plaene": {}}
            db[name]["geb_datum"] = st.session_state.geb_datum.isoformat()
            db[name]["altersgruppe"] = st.session_state.altersgruppe
            db[name]["energiebedarf"] = st.session_state.energiebedarf
            db[name]["gewicht"] = st.session_state.gewicht
            db[name]["zielgewicht"] = st.session_state.zielgewicht
            db[name]["verdaulichkeit"] = st.session_state.verdaulichkeit
            db[name]["angezeigte_spalten"] = st.session_state.angezeigte_spalten
            speichere_db(db)
            st.success(f"Profil für '{name}' wurde erfolgreich gespeichert!")

# ==========================================
# TAB 2: FUTTERPLAN & ANALYSE
# ==========================================
with tab2:
    st.header("Futterplan & Bedarfsanalyse")

    col_opt1, col_opt2, col_opt3 = st.columns([2, 1, 1])
    with col_opt1:
        with st.expander("⚙️ Nährstoff-Spalten anpassen"):
            if st.button("🔄 Alle Nährstoffe wieder einblenden", use_container_width=True):
                st.session_state.angezeigte_spalten = naehrstoff_spalten.copy()
                # Status aller Checkboxen hart auf 'True' setzen
                for col in naehrstoff_spalten:
                    st.session_state[f"chk_{col}"] = True
                
                # Auto-Save direkt ausführen
                aktueller_hund = st.session_state.hund_name
                if aktueller_hund and aktueller_hund in db:
                    db[aktueller_hund]["angezeigte_spalten"] = st.session_state.angezeigte_spalten
                    speichere_db(db)
                st.rerun()
                
            st.markdown("Klicke die Nährstoffe an, die angezeigt werden sollen:")
            chk_cols = st.columns(4)
            neue_auswahl = []
            
            for index, col_name in enumerate(naehrstoff_spalten):
                chk_key = f"chk_{col_name}"
                
                # Initialisiere den Key, falls er noch nicht im Session State existiert
                if chk_key not in st.session_state:
                    st.session_state[chk_key] = col_name in st.session_state.angezeigte_spalten
                    
                with chk_cols[index % 4]:
                    if st.checkbox(col_name, key=chk_key):
                        neue_auswahl.append(col_name)
            
            # --- AUTO-SAVE FÜR NÄHRSTOFF-AUSWAHL ---
            if neue_auswahl != st.session_state.angezeigte_spalten:
                st.session_state.angezeigte_spalten = neue_auswahl
                aktueller_hund = st.session_state.hund_name
                if aktueller_hund and aktueller_hund in db:
                    db[aktueller_hund]["angezeigte_spalten"] = neue_auswahl
                    speichere_db(db)
                st.rerun()

    with col_opt2:
        zeige_min = st.toggle("🔽 Mindestbedarf anzeigen", value=True)
        zeige_max = st.toggle("🔼 Maximalbedarf anzeigen", value=True)
        
    with col_opt3:
        if st.button("🔄 Excel neu laden", help="Aktualisiert die Zutaten", use_container_width=True):
            lade_daten.clear()
            st.rerun()

    # --- BERECHNUNG ---
    soll_zeile = {col: 0.0 for col in naehrstoff_spalten}
    min_zeile = {col: 0.0 for col in naehrstoff_spalten}
    max_zeile = {col: 0.0 for col in naehrstoff_spalten}

    if st.session_state.gewicht > 0:
        try:
            alter_mask = df_h["Multiplikator Energie nach Alter"] == st.session_state.altersgruppe
            energie_faktor = float(df_h.loc[alter_mask, "MJ je kg Stoffwechselgewicht"].values[0])
            soll_mj_basis = stoffwechselgewicht * energie_faktor
            
            eb = st.session_state.energiebedarf
            if eb == "kastriert (- 15%)": soll_mj = soll_mj_basis * 0.85
            elif eb == "kastriert (- 20%)": soll_mj = soll_mj_basis * 0.80
            elif eb == "Übergewicht (- 30%)": soll_mj = soll_mj_basis * 0.70
            else: soll_mj = soll_mj_basis
            
            mj_col = next((col for col in naehrstoff_spalten if "MJ" in col), None)
            kcal_col = next((col for col in naehrstoff_spalten if "kcal" in col), None)
            if mj_col: soll_zeile[mj_col] = soll_mj
            if kcal_col: soll_zeile[kcal_col] = soll_mj * 238.8
            
            gewichte = pd.to_numeric(df_h["Gewicht kg"].dropna(), errors='coerce').dropna()
            naechstes_gewicht = gewichte.iloc[(gewichte - st.session_state.gewicht).abs().argsort()[:1]].values[0]
            gewicht_mask = df_h["Gewicht kg"] == naechstes_gewicht
            
            vrp_faktor_soll = float(df_h.loc[gewicht_mask, "Empfehlung je kg KM"].values[0])
            vrp_faktor_min = float(df_h.loc[gewicht_mask, "Mindestbedarf je kg KM"].values[0])
            
            soll_vrp = st.session_state.gewicht * vrp_faktor_soll
            min_vrp = st.session_state.gewicht * vrp_faktor_min
            
            verd_mask = df_h.iloc[:, 5] == st.session_state.verdaulichkeit
            verd_prozent = float(df_h.loc[verd_mask, df_h.columns[6]].values[0])
            if verd_prozent > 1: verd_prozent = verd_prozent / 100.0
                
            soll_rohprotein = soll_vrp / verd_prozent
            min_rohprotein = min_vrp / verd_prozent
            
            vrp_col = next((col for col in naehrstoff_spalten if "vRP" in col or "vRp" in col), None)
            prot_col = next((col for col in naehrstoff_spalten if "Rohprotein" in col), None)
            if vrp_col: soll_zeile[vrp_col] = soll_vrp; min_zeile[vrp_col] = min_vrp
            if prot_col: soll_zeile[prot_col] = soll_rohprotein; min_zeile[prot_col] = min_rohprotein
            
            soll_faktoren = {"Calcium": 80, "Phosphor": 60, "Magnesium": 12, "Natrium": 50, "Kalium": 55, "Chlorid": 75, "Eisen": 1.4, "Zink": 0.9, "Selen": 2.5, "Kupfer": 0.1, "Mangan": 0.07, "Jod": 15, "Vitamin D3": 0.25, "Vitamin E [": 0.5, "Vitamin A [": 22.5, "Vitamin B1 [": 20, "Vitamin B2": 50, "Vit. B3": 200, "Vitamin B6": 20, "Vitamin B12": 0.5, "Folsäure": 4, "Biotin": 2, "EPA": 45, "DHA": 30, "Linolsäure": 150, "Arginin": 56, "Cystein": 51, "Histidin": 32, "Lysin": 68, "Methionin [": 76}
            min_faktoren = {"Rohfett": 1, "Calcium": 62, "Phosphor": 46, "Magnesium": 10, "Natrium": 20, "Mangan": 0.01, "Jod": 13, "EPA": 18, "DHA": 12, "Linolsäure": 50, "Arginin": 56, "Cystein": 31, "Histidin": 32, "Lysin": 56, "Methionin [": 56}
            max_faktoren = {"Kalium": 2000, "Selen": 5, "Mangan": 0.13, "Vitamin D3": 1.25, "Vitamin E [": 11, "Vitamin A [": 4500, "EPA": 90, "DHA": 60}
            
            for naehrstoff, faktor in soll_faktoren.items():
                col_name = next((col for col in naehrstoff_spalten if naehrstoff in col), None)
                if col_name: soll_zeile[col_name] = st.session_state.gewicht * faktor
            for naehrstoff, faktor in min_faktoren.items():
                col_name = next((col for col in naehrstoff_spalten if naehrstoff in col), None)
                if col_name: min_zeile[col_name] = st.session_state.gewicht * faktor
            for naehrstoff, faktor in max_faktoren.items():
                col_name = next((col for col in naehrstoff_spalten if naehrstoff in col), None)
                if col_name: max_zeile[col_name] = st.session_state.gewicht * faktor

            col_d_ie = next((col for col in naehrstoff_spalten if "Vitamin D I.E." in col or "Vitamin D IE" in col), None)
            col_a_ie = next((col for col in naehrstoff_spalten if "Vitamin A IE" in col or "Vitamin A I.E." in col), None)
            col_e_ie = next((col for col in naehrstoff_spalten if "Vitamin E IE" in col or "Vitamin E I.E." in col), None)
            col_d3 = next((col for col in naehrstoff_spalten if "Vitamin D3" in col), None)
            col_a = next((col for col in naehrstoff_spalten if "Vitamin A [" in col), None)
            col_e = next((col for col in naehrstoff_spalten if "Vitamin E [" in col), None)

            if col_d3 and col_d_ie: soll_zeile[col_d_ie] = soll_zeile[col_d3] / 0.025; max_zeile[col_d_ie] = max_zeile[col_d3] / 0.025
            if col_a and col_a_ie: soll_zeile[col_a_ie] = soll_zeile[col_a] / 0.3; max_zeile[col_a_ie] = max_zeile[col_a] / 0.3
            if col_e and col_e_ie: soll_zeile[col_e_ie] = soll_zeile[col_e] / 0.67; max_zeile[col_e_ie] = max_zeile[col_e] / 0.67
        except Exception as e:
            pass 

    min_daten = {"Zutat": "🔽 MINDESTBEDARF", "Menge (g)": 0.0}
    min_daten.update(min_zeile)
    max_daten = {"Zutat": "🔼 MAXIMALBEDARF", "Menge (g)": 0.0}
    max_daten.update(max_zeile)
    soll_daten = {"Zutat": "🎯 EMPFOHLENER BEDARF", "Menge (g)": 0.0}
    soll_daten.update(soll_zeile)

    ergebnis_zeilen_fuer_editor = []
    summen_zeile = {col: 0.0 for col in naehrstoff_spalten}
    aktuelle_gesamtmenge = 0.0

    umrechnungsfaktoren_ist = {
        "Kupfer [": 0.001, "Mangan [": 0.001, "Vitamin B1 [": 1000.0, "Vitamin B2 [": 1000.0,
        "Vit. B3/Niacin [": 1000.0, "Omega-3, gesamt [": 1000.0, "EPA [": 1000.0, "DHA [": 1000.0,
        "Omega-6, gesamt [": 1000.0, "Linolsäure [": 1000.0, "Arginin [": 1000.0, "Cystein [": 1000.0,
        "Histidin [": 1000.0, "Lysin [": 1000.0, "Methionin [": 1000.0
    }

    for item in st.session_state.futter_data:
        zeilen_daten = {"Aktiv": item["Aktiv"], "Zutat": item["Zutat"], "Menge (g)": float(item["Menge"])}
        
        if item["Aktiv"]:
            aktuelle_gesamtmenge += item["Menge"]
            
        zutat_werte = zutaten_lookup.get(item["Zutat"], {})
            
        for col in naehrstoff_spalten:
            if zutat_werte:
                wert_100g = zutat_werte.get(col, 0.0)
                try: 
                    berechneter_wert = (float(wert_100g) / 100) * item["Menge"]
                    for key, faktor in umrechnungsfaktoren_ist.items():
                        if key in col:
                            berechneter_wert *= faktor
                            break
                except: 
                    berechneter_wert = 0.0
            else:
                berechneter_wert = 0.0
                
            zeilen_daten[col] = berechneter_wert
            if item["Aktiv"]:
                summen_zeile[col] += berechneter_wert
                
        ergebnis_zeilen_fuer_editor.append(zeilen_daten)

    summen_daten = {"Zutat": "✅ IST-VERSORGUNG", "Menge (g)": aktuelle_gesamtmenge}
    summen_daten.update(summen_zeile)

    # --- ZUSAMMENFASSUNG TABELLE (ORIGINAL LAYOUT WIEDERHERGESTELLT) ---
    st.markdown("### Bedarf & Ist-Versorgung")
    
    anzeige_zeilen = []
    if zeige_min: anzeige_zeilen.append(min_daten)
    if zeige_max: anzeige_zeilen.append(max_daten)
    anzeige_zeilen.append(soll_daten)
    anzeige_zeilen.append(summen_daten)

    df_summary = pd.DataFrame(anzeige_zeilen)

    if not df_summary.empty:
        spalten_filter = ["Zutat", "Menge (g)"] + st.session_state.angezeigte_spalten
        df_summary_disp = df_summary[[c for c in spalten_filter if c in df_summary.columns]]
        df_summary_disp = df_summary_disp.set_index("Zutat").round(2)

        def summary_hervorheben(row):
            if row.name == "🎯 EMPFOHLENER BEDARF":
                return ['background-color: #d1ecf1; font-weight: bold; color: #0c5460'] * len(row)
            elif row.name == "✅ IST-VERSORGUNG":
                return ['background-color: #d4edda; font-weight: bold; color: #155724'] * len(row)
            elif row.name == "🔽 MINDESTBEDARF":
                return ['background-color: #fff3cd; font-weight: bold; color: #856404'] * len(row)
            elif row.name == "🔼 MAXIMALBEDARF":
                return ['background-color: #ffe8cc; font-weight: bold; color: #d97706'] * len(row)
            return [''] * len(row)

        df_styled = df_summary_disp.style.apply(summary_hervorheben, axis=1).format("{:.2f}")
        st.dataframe(df_styled, use_container_width=True)

    # --- TEXTBOX 1: NUR CALCIUM-PHOSPHOR VERHÄLTNIS ---
    ca_col = next((c for c in naehrstoff_spalten if "Calcium" in c), None)
    p_col = next((c for c in naehrstoff_spalten if "Phosphor" in c), None)
    ratio = 0
    if ca_col and p_col:
        ist_ca = summen_zeile.get(ca_col, 0)
        ist_p = summen_zeile.get(p_col, 0)
        if ist_p > 0:
            ratio = ist_ca / ist_p
            st.markdown(f"<div style='padding:10px; border-radius:5px; background-color:#f8f9fa; border-left: 5px solid #7bdcb5; margin-top: 10px;'>"
                        f"<strong>🦴 Calcium-Phosphor-Verhältnis (Ca:P):</strong> {ratio:.2f} : 1 (Ideal: 1,2 - 1,4 : 1)</div>", unsafe_allow_html=True)

    # --- TEXTBOX 2: NUR URIN-PH-WERT & KAB BERECHNUNG ---
    def hole_wert(suchwort):
        spalte = next((c for c in naehrstoff_spalten if c == suchwort or suchwort in c), None)
        return summen_zeile.get(spalte, 0.0) if spalte else 0.0

    ca_mg = hole_wert("Calcium")
    p_mg = hole_wert("Phosphor")
    mg_mg = hole_wert("Magnesium")
    na_mg = hole_wert("Natrium")
    k_mg = hole_wert("Kalium")
    rp_g = hole_wert("Rohprotein")
    ts_g = hole_wert("TS") 
    
    if ts_g <= 0:
        ts_g = 1.0  

    kab = (50 * ca_mg + 82 * mg_mg + 43 * na_mg + 26 * k_mg - 65 * p_mg) / 1000
    ph_wert = ((kab - (rp_g * 0.625)) / (ts_g / 100) * 0.019) + 6.5

    st.markdown(f"<div style='padding:10px; border-radius:5px; background-color:#f0f8ff; border-left: 5px solid #3498db; margin-top: 10px;'>"
                f"<strong>🧪 Urin-Analytik (geschätzt):</strong> pH-Wert = {ph_wert:.2f} | KAB = {kab:.2f} meq</div>", unsafe_allow_html=True)

    # --- TOXIZITÄTS-WARNUNGEN ---
    tox_stoffe = ["Vitamin A", "Vitamin D", "Kupfer"]
    warnungen_ausgegeben = False
    
    for col, max_v in max_zeile.items():
        if max_v > 0 and summen_zeile.get(col, 0) > max_v:
            for tox in tox_stoffe:
                if tox in col:
                    if not warnungen_ausgegeben:
                        st.markdown("<br>", unsafe_allow_html=True)
                        warnungen_ausgegeben = True
                    st.error(f"🚨 **WARNUNG:** Die toxische Obergrenze für **{col}** wurde überschritten! (Ist: {summen_zeile[col]:.2f} | Max: {max_v:.2f})")

    st.divider()

    # --- INTERAKTIVE ZUTATEN-TABELLE (MIT FIXIERTEN SPALTEN) ---
    st.markdown("### Futterplan bearbeiten")
    if st.session_state.futter_data:
        df_editor = pd.DataFrame(ergebnis_zeilen_fuer_editor)
        df_editor.insert(0, "🗑️ Löschen", False)
        
        display_cols = ["🗑️ Löschen", "Aktiv", "Zutat", "Menge (g)"] + st.session_state.angezeigte_spalten
        df_display = df_editor[display_cols].copy()
        
        edited_df = st.data_editor(
            df_display,
            disabled=["Zutat"] + st.session_state.angezeigte_spalten,
            hide_index=True,
            use_container_width=True,
            key="zutaten_editor",
            column_config={
                "🗑️ Löschen": st.column_config.CheckboxColumn("🗑️ Löschen", pinned=True),
                "Aktiv": st.column_config.CheckboxColumn("Aktiv", pinned=True),
                "Zutat": st.column_config.TextColumn("Zutat", pinned=True),
                "Menge (g)": st.column_config.NumberColumn("Menge (g)", pinned=True)
            }
        )
        
        needs_rerun = False
        to_delete = []
        for i, row in edited_df.iterrows():
            if row["🗑️ Löschen"]:
                to_delete.append(i)
                needs_rerun = True
            else:
                if st.session_state.futter_data[i]["Aktiv"] != row["Aktiv"]:
                    st.session_state.futter_data[i]["Aktiv"] = row["Aktiv"]
                    needs_rerun = True
                if st.session_state.futter_data[i]["Menge"] != row["Menge (g)"]:
                    st.session_state.futter_data[i]["Menge"] = row["Menge (g)"]
                    needs_rerun = True
                    
        if to_delete:
            for i in sorted(to_delete, reverse=True):
                st.session_state.futter_data.pop(i)
                
        if needs_rerun:
            st.rerun()
    else:
        st.info("Füge unten Zutaten hinzu, um deinen Futterplan zu erstellen.")

    st.write("---")

    col_auswahl, col_menge, col_add = st.columns([5, 2, 2])
    with col_auswahl:
        zutat_neu = st.selectbox("Lebensmittel", liste_zutaten, index=None, placeholder="Bitte wählen...", 
                                 label_visibility="collapsed", key=f"zutat_auswahl_{st.session_state.selectbox_key}")
    with col_menge:
        menge_neu = st.number_input("Menge (g)", value=100.0, step=10.0, label_visibility="collapsed")
    with col_add:
        submit = st.button("➕ Hinzufügen", use_container_width=True)
        if submit:
            if zutat_neu is not None:
                st.session_state.futter_data.append({"Zutat": zutat_neu, "Menge": menge_neu, "Aktiv": True})
                st.session_state.selectbox_key += 1 
                st.rerun()
            else:
                st.warning("Bitte wähle zuerst ein Lebensmittel aus der Liste aus!")

    st.write("---")

    # --- FUTTERPLAN SPEICHERN & LADEN ---
    st.markdown("### 💾 Futterplan speichern / laden")
    aktueller_hund = st.session_state.hund_name

    if aktueller_hund in db and aktueller_hund.strip() != "":
        gespeicherte_plaene = list(db[aktueller_hund].get("plaene", {}).keys())
        
        col_p_load1, col_p_load2 = st.columns([3, 1])
        with col_p_load1:
            auswahl_plan = st.selectbox("Gespeicherten Plan laden:", ["-- Bitte wählen --", "✨ + Neuer Futterplan"] + gespeicherte_plaene, label_visibility="collapsed")
        with col_p_load2:
            if st.button("📂 Plan öffnen", use_container_width=True):
                if auswahl_plan == "✨ + Neuer Futterplan":
                    st.session_state.futter_data = []
                    st.session_state.aktiver_plan_name = "Neuer Plan"
                    st.rerun()
                elif auswahl_plan != "-- Bitte wählen --":
                    st.session_state.futter_data = db[aktueller_hund]["plaene"][auswahl_plan]
                    st.session_state.aktiver_plan_name = auswahl_plan
                    st.rerun()
                    
        col_p_save1, col_p_save2 = st.columns([3, 1])
        with col_p_save1:
            neuer_plan_name = st.text_input("Name für diesen Futterplan:", value=st.session_state.aktiver_plan_name, label_visibility="collapsed")
        with col_p_save2:
            if st.button("Plan speichern / aktualisieren", use_container_width=True):
                if neuer_plan_name.strip() == "":
                    st.error("Bitte gib dem Plan einen Namen.")
                else:
                    db[aktueller_hund]["plaene"][neuer_plan_name] = st.session_state.futter_data
                    speichere_db(db)
                    st.session_state.aktiver_plan_name = neuer_plan_name
                    st.success(f"Plan '{neuer_plan_name}' wurde gespeichert / aktualisiert!")
    else:
        st.info("⚠️ Bitte speichere zuerst das Hundeprofil im Tab 'Hundedaten', um Futterpläne ablegen zu können.")


    # ==========================================
    # GRAFISCHE AUSWERTUNG (DIAGRAMME VORBEREITEN)
    # ==========================================
    st.divider()
    st.markdown("### 📈 Grafische Auswertung")
    st.markdown("**Bedarfsdeckung (in % vom Soll)**")
    
    # NEU: Eigene Auswahl für das Diagramm mit Checkboxen & Blacklist (inkl. "lactose")
    with st.expander("📊 Nährstoffe für Diagramm auswählen"):
        chart_blacklist = ["nfe", "rohfaser", "rohasche", "wasser", "trockensubstanz", "ts", "vitamin c", "vitamin k", "laktose", "lactose", "omega-3", "omega-6"]
        
        erlaubte_chart_spalten = []
        for col in st.session_state.angezeigte_spalten:
            if not any(b in col.lower() for b in chart_blacklist):
                erlaubte_chart_spalten.append(col)
                
        if 'chart_spalten' not in st.session_state:
            st.session_state.chart_spalten = erlaubte_chart_spalten.copy()
            
        # Aktualisieren, falls sich oben in der großen Tabelle etwas geändert hat
        st.session_state.chart_spalten = [c for c in st.session_state.chart_spalten if c in erlaubte_chart_spalten]
        
        if not st.session_state.chart_spalten and erlaubte_chart_spalten:
             st.session_state.chart_spalten = erlaubte_chart_spalten.copy()
             
        chk_cols_chart = st.columns(4)
        neue_chart_auswahl = []
        for index, col_name in enumerate(erlaubte_chart_spalten):
            is_checked = col_name in st.session_state.chart_spalten
            with chk_cols_chart[index % 4]:
                if st.checkbox(col_name, value=is_checked, key=f"chart_chk_{col_name}"):
                    neue_chart_auswahl.append(col_name)
                    
        if neue_chart_auswahl != st.session_state.chart_spalten:
            st.session_state.chart_spalten = neue_chart_auswahl
            st.rerun()
    
    chart_daten = []
    
    # Nur Wertebereiche in der Legende
    status_unter = "< 80 %"
    status_optimal = "80 % - 120 %"
    status_ueber = "> 120 %"
    
    for col in st.session_state.chart_spalten:
        soll_wert = soll_zeile.get(col, 0)
        min_wert = min_zeile.get(col, 0)
        max_wert = max_zeile.get(col, 0)
        ist_wert = summen_zeile.get(col, 0)
        
        # LOGIK FÜR FEHLENDE SOLLWERTE (Die 80%-Regel)
        if soll_wert == 0 and min_wert > 0:
            soll_wert = min_wert * 1.25  
            
        if soll_wert > 0:
            prozent = (ist_wert / soll_wert) * 100
            
            if (min_wert > 0 and ist_wert < min_wert) or (ist_wert < soll_wert * 0.8):
                status = status_unter
            elif (max_wert > 0 and ist_wert > max_wert) or (ist_wert > soll_wert * 1.2):
                status = status_ueber
            else:
                status = status_optimal
                
            chart_daten.append({"Nährstoff": col, "Erfüllung (%)": prozent, "Status": status})
            
    fig_bar = None
    if chart_daten:
        df_bar = pd.DataFrame(chart_daten)
        
        df_bar['Nährstoff'] = df_bar['Nährstoff'].str.replace(r' \[g/100g\]', ' [g]', regex=True)
        df_bar['Nährstoff'] = df_bar['Nährstoff'].str.replace(r' \[mg/100g\]', ' [mg]', regex=True)
        df_bar['Nährstoff'] = df_bar['Nährstoff'].str.replace(r' \[µg/100g\]', ' [µg]', regex=True)
        
        color_map = {status_unter: "#c4f0d5", status_optimal: "#5dd893", status_ueber: "#1e8449"}
        
        fig_bar = px.bar(df_bar, x='Nährstoff', y='Erfüllung (%)', color='Status', color_discrete_map=color_map)
        fig_bar.add_hline(y=100, line_dash="dot", line_color="#2c3e50", annotation_text="100% (Soll)")
        
        cleaned_spalten = [c.replace(' [g/100g]', ' [g]').replace(' [mg/100g]', ' [mg]').replace(' [µg/100g]', ' [µg]') for c in st.session_state.chart_spalten]
        
        fig_bar.update_xaxes(
            categoryorder='array', 
            categoryarray=cleaned_spalten,
            tickmode='linear',
            tickfont=dict(size=14)
        )
        
        # Titel der X-Achse (Nährstoff) entfernen: xaxis_title=None
        fig_bar.update_layout(height=500, margin=dict(t=20, b=150, l=20, r=20), plot_bgcolor="white", xaxis_tickangle=-90, xaxis_title=None)
        st.plotly_chart(fig_bar, use_container_width=True)
    else:
        st.info("Bitte wähle Nährstoffe aus, um das Diagramm zu generieren.")

    st.write("---")
    st.markdown("**Rationszusammensetzung (Menge in g)**")
    
    fig_pie = None
    aktive_zutaten = [item for item in st.session_state.futter_data if item["Aktiv"] and float(item["Menge"]) > 0]
    
    col_spacer1, col_pie, col_spacer2 = st.columns([1, 2, 1])
    with col_pie:
        if aktive_zutaten:
            df_pie = pd.DataFrame(aktive_zutaten)
            mint_palette = ['#e8f4f0', '#c4f0d5', '#a2e8bf', '#80e0a9', '#5dd893', '#3cd07d', '#28b968']
            fig_pie = px.pie(df_pie, values='Menge', names='Zutat', hole=0.4, color_discrete_sequence=mint_palette)
            fig_pie.update_layout(margin=dict(t=20, b=20, l=20, r=20))
            st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.info("Keine aktiven Zutaten vorhanden.")


    # --- EXPORT BEREICH (BUNTES PDF MIT AUSWAHL) ---
    st.write("---")
    st.markdown("### 📄 Export")
    
    if HAS_FPDF:
        st.markdown("**Wähle aus, was im PDF angezeigt werden soll:**")
        col_pdf1, col_pdf2, col_pdf3 = st.columns(3)
        with col_pdf1:
            pdf_opt_profil = st.checkbox("Hundedaten (Profil)", value=True)
            pdf_opt_zutaten = st.checkbox("Zusammensetzung der Ration", value=True)
        with col_pdf2:
            pdf_opt_cap = st.checkbox("Calcium-Phosphor-Verhältnis", value=True)
            pdf_opt_urin = st.checkbox("Urin-Analytik (pH & KAB)", value=True)
            pdf_opt_naehrstoffe = st.checkbox("Nährstoff-Ist-Werte", value=False)
        with col_pdf3:
            pdf_opt_grafiken = st.checkbox("Grafische Auswertungen (Diagramme)", value=True)

        if st.button("📄 PDF-Export vorbereiten", type="primary"):
            with st.spinner("PDF wird erstellt... Das kann je nach Diagramm-Größe einen Moment dauern."):
                try:
                    pdf = FPDF()
                    
                    pdf.add_page(orientation='P')
                    
                    pdf.set_font("Arial", 'B', 22)
                    pdf.set_text_color(44, 62, 80)
                    pdf.cell(0, 15, txt=clean_text("CaLu Futterrechner"), ln=True, align="C")
                    
                    pdf.set_font("Arial", '', 14)
                    pdf.set_text_color(127, 140, 141)
                    pdf.cell(0, 10, txt=clean_text(f"Futterplan: {st.session_state.aktiver_plan_name}"), ln=True, align="C")
                    pdf.ln(5)
                    
                    if pdf_opt_profil:
                        pdf.set_fill_color(248, 249, 250)
                        pdf.set_text_color(44, 62, 80)
                        pdf.set_font("Arial", '', 11)
                        geb_str = st.session_state.geb_datum.strftime('%d.%m.%Y')
                        profil_text = f"Hund: {st.session_state.hund_name}   |   Gewicht: {st.session_state.gewicht} kg   |   Geburtsdatum: {geb_str}"
                        pdf.cell(0, 10, txt=clean_text(profil_text), ln=True, align="C", fill=True)
                        pdf.ln(5)

                    if pdf_opt_zutaten:
                        pdf.set_font("Arial", 'B', 14)
                        pdf.set_text_color(44, 62, 80)
                        pdf.cell(0, 10, txt=clean_text("Zusammensetzung der Ration"), ln=True)
                        
                        pdf.set_font("Arial", 'B', 11)
                        pdf.set_fill_color(232, 244, 240)
                        pdf.set_text_color(39, 174, 96)
                        pdf.cell(140, 10, txt=clean_text("Zutat"), border=0, fill=True)
                        pdf.cell(50, 10, txt=clean_text("Menge (g)"), border=0, fill=True, ln=True)
                        
                        pdf.set_font("Arial", '', 11)
                        pdf.set_text_color(44, 62, 80)
                        for item in st.session_state.futter_data:
                            if item["Aktiv"]:
                                pdf.cell(140, 10, txt=clean_text(item["Zutat"]), border='B')
                                pdf.cell(50, 10, txt=clean_text(str(item["Menge"])), border='B', ln=True)
                        pdf.ln(10)

                    if pdf_opt_cap and ratio > 0:
                        pdf.set_fill_color(248, 249, 250)
                        pdf.set_text_color(44, 62, 80)
                        pdf.set_font("Arial", 'B', 12)
                        cap_text = f"Calcium-Phosphor-Verhaeltnis (Ca:P): {ratio:.2f} : 1 (Ideal: 1,2 - 1,4 : 1)"
                        pdf.cell(0, 12, txt=clean_text(cap_text), ln=True, align="L", fill=True)
                        pdf.ln(5)
                        
                    if pdf_opt_urin:
                        # Bläulicher Hintergrund wie in der Streamlit-App
                        pdf.set_fill_color(240, 248, 255)
                        pdf.set_text_color(44, 62, 80)
                        pdf.set_font("Arial", 'B', 12)
                        urin_text = f"Urin-Analytik (geschaetzt): pH-Wert = {ph_wert:.2f} | KAB = {kab:.2f} meq"
                        pdf.cell(0, 12, txt=clean_text(urin_text), ln=True, align="L", fill=True)
                        pdf.ln(5)

                    if pdf_opt_naehrstoffe:
                        pdf.add_page(orientation='P')
                        pdf.set_font("Arial", 'B', 14)
                        pdf.set_text_color(44, 62, 80)
                        pdf.cell(0, 10, txt=clean_text("Erreichte Naehrstoffwerte (Ist-Zustand)"), ln=True)
                        
                        pdf.set_font("Arial", 'B', 10)
                        pdf.set_fill_color(232, 244, 240)
                        pdf.set_text_color(39, 174, 96)
                        pdf.cell(95, 8, txt=clean_text("Naehrstoff"), border=0, fill=True)
                        pdf.cell(95, 8, txt=clean_text("Erreichte Menge"), border=0, fill=True, ln=True)
                        
                        pdf.set_font("Arial", '', 10)
                        pdf.set_text_color(44, 62, 80)
                        for col in st.session_state.angezeigte_spalten:
                            ist_wert = summen_zeile.get(col, 0)
                            pdf.cell(95, 8, txt=clean_text(col), border='B')
                            pdf.cell(95, 8, txt=clean_text(f"{ist_wert:.2f}"), border='B', ln=True)
                        pdf.ln(10)

                    if pdf_opt_grafiken:
                        try:
                            if fig_bar is not None:
                                pdf.add_page(orientation='L')
                                pdf.set_font("Arial", 'B', 16)
                                pdf.set_text_color(44, 62, 80)
                                pdf.cell(0, 10, txt=clean_text("Grafische Auswertungen"), ln=True)
                                pdf.ln(5)
                                
                                pdf.set_font("Arial", 'B', 12)
                                pdf.cell(0, 10, txt=clean_text("Bedarfsdeckung (in % vom Soll)"), ln=True)
                                
                                # Grafiken nach unten schieben
                                pdf.ln(15)
                                
                                tmp_bar = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                                
                                fig_bar.update_layout(margin=dict(t=20, b=250, l=40, r=20))
                                fig_bar.update_xaxes(tickfont=dict(size=16))
                                
                                fig_bar.write_image(tmp_bar.name, engine="kaleido", width=1600, height=800)
                                pdf.image(tmp_bar.name, x=10, w=277)
                                tmp_bar.close()
                                os.unlink(tmp_bar.name)

                            if fig_pie is not None:
                                pdf.add_page(orientation='P')
                                pdf.set_font("Arial", 'B', 16)
                                pdf.set_text_color(44, 62, 80)
                                pdf.cell(0, 10, txt=clean_text("Rationszusammensetzung"), ln=True)
                                
                                pdf.ln(50)
                                
                                tmp_pie = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                                fig_pie.write_image(tmp_pie.name, engine="kaleido", width=800, height=500)
                                pdf.image(tmp_pie.name, x=10, w=190)
                                tmp_pie.close()
                                os.unlink(tmp_pie.name)
                        except Exception as e:
                            pdf.add_page(orientation='P')
                            pdf.set_font("Arial", 'I', 10)
                            pdf.set_text_color(231, 76, 60)
                            pdf.cell(0, 10, txt=clean_text(f"Hinweis: Diagramme konnten nicht geladen werden. Fehler: {e}"), ln=True)

                    pdf_out = pdf.output(dest='S')
                    if isinstance(pdf_out, str):
                        st.session_state['ready_pdf'] = pdf_out.encode('latin-1', 'replace')
                    else:
                        st.session_state['ready_pdf'] = bytes(pdf_out)
                        
                except Exception as e:
                    st.error(f"Es gab ein Problem beim Generieren des PDFs: {e}")

        if 'ready_pdf' in st.session_state:
            st.success("✅ Das PDF ist fertig und bereit zum Download!")
            st.download_button(
                label="📥 Buntes PDF herunterladen", 
                data=st.session_state['ready_pdf'], 
                file_name=f"Futterplan_{st.session_state.hund_name}.pdf", 
                mime='application/pdf'
            )
    else:
        st.warning("⚠️ Um das bunte PDF zu generieren, fehlt das fpdf-Modul.")
        st.caption("Bitte installiere es in deiner Konsole mit: `pip install fpdf`")

# ==========================================
# TAB 3: TOOLS & UMRECHNER
# ==========================================
with tab3:
    st.header("🧮 Einheiten-Umrechner")
    st.markdown("Rechne Gewichte und Internationale Einheiten (IE) für Vitamine bequem um.")
    
    col_val, col_in, col_out = st.columns(3)
    with col_val:
        eingabe_wert = st.number_input("Wert", value=1.0, step=1.0)
    with col_in:
        einheit_von = st.selectbox("Von", ["kg", "g", "mg", "µg", "IE"])
    with col_out:
        einheit_in = st.selectbox("In", ["kg", "g", "mg", "µg", "IE"])

    stoff = None
    if einheit_von == "IE" or einheit_in == "IE":
        stoff = st.selectbox("Für welchen Stoff berechnest du die IE?", ["Vitamin A", "Vitamin D3", "Vitamin E"])

    def konvertiere_zu_mikrogramm(wert, einheit, stoff_art):
        if einheit == "kg": return wert * 1e9
        if einheit == "g": return wert * 1e6
        if einheit == "mg": return wert * 1000
        if einheit == "µg": return wert
        if einheit == "IE":
            if stoff_art == "Vitamin A": return wert * 0.3
            if stoff_art == "Vitamin D3": return wert * 0.025
            if stoff_art == "Vitamin E": return wert * 670.0 
        return 0

    def konvertiere_von_mikrogramm(wert_ug, ziel_einheit, stoff_art):
        if ziel_einheit == "kg": return wert_ug / 1e9
        if ziel_einheit == "g": return wert_ug / 1e6
        if ziel_einheit == "mg": return wert_ug / 1000
        if ziel_einheit == "µg": return wert_ug
        if ziel_einheit == "IE":
            if stoff_art == "Vitamin A": return wert_ug / 0.3
            if stoff_art == "Vitamin D3": return wert_ug / 0.025
            if stoff_art == "Vitamin E": return wert_ug / 670.0
        return 0

    if st.button("🔄 Berechnen", type="primary"):
        ug_basis = konvertiere_zu_mikrogramm(eingabe_wert, einheit_von, stoff)
        ergebnis = konvertiere_von_mikrogramm(ug_basis, einheit_in, stoff)
        
        zusatz = f" (für {stoff})" if stoff else ""
        
        str_eingabe = formatiere_zahl(eingabe_wert)
        str_ergebnis = formatiere_zahl(ergebnis)
        
        st.success(f"**Ergebnis:** {str_eingabe} {einheit_von} = **{str_ergebnis} {einheit_in}** {zusatz}")
