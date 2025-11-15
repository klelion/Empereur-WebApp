import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from pathlib import Path
import shutil

# ======================
# CONFIG
# ======================

TEMPLATE_FILE = "Systeme_Entrainement_Empereur_ULTIME.xlsx"
DATA_FILE = "empereur_data.xlsx"


def get_excel_file(data_only=False):
    """
    Utilise une copie modifiable de l'Excel (DATA_FILE).
    Si elle n'existe pas, on la crée à partir du TEMPLATE_FILE.
    """
    template_path = Path(TEMPLATE_FILE)
    if not template_path.exists():
        st.error(f"Fichier modèle introuvable : {template_path.resolve()}")
        st.stop()

    data_path = Path(DATA_FILE)
    if not data_path.exists():
        shutil.copy(template_path, data_path)

    wb = load_workbook(data_path, data_only=data_only)
    return wb, data_path


# ======================
# UTILITAIRES EXCEL
# ======================

def find_row_by_session(ws, session_number, col_session=1):
    max_row = ws.max_row
    for row in range(2, max_row + 1):
        val = ws.cell(row=row, column=col_session).value
        if val == session_number:
            return row
    for row in range(2, max_row + 2):
        val = ws.cell(row=row, column=col_session).value
        if val is None:
            ws.cell(row=row, column=col_session).value = session_number
            return row


def get_next_lifestyle_day(ws):
    last = 0
    for row in range(2, ws.max_row + 1):
        val = ws.cell(row=row, column=1).value
        if isinstance(val, int) and val > last:
            last = val
    return last + 1 if last > 0 else 1


# ======================
# PAGES
# ======================

def page_lifestyle():
    st.header("📋 Lifestyle – Saisie quotidienne")

    wb, data_path = get_excel_file()
    ws = wb["Lifestyle"]

    jour = get_next_lifestyle_day(ws)
    st.info(f"Jour enregistré : **{jour}** (prochain enregistrement)")

    col1, col2 = st.columns(2)
    with col1:
        sommeil = st.number_input("Sommeil (0-10)", 0.0, 10.0, 7.0, 0.5)
        hydrat = st.number_input("Hydratation (0-10)", 0.0, 10.0, 8.0, 0.5)
        nutri = st.number_input("Nutrition (0-10)", 0.0, 10.0, 7.0, 0.5)
        stress = st.number_input("Stress (0-10, plus = pire)", 0.0, 10.0, 3.0, 0.5)
    with col2:
        conc = st.number_input("Concentration (0-10)", 0.0, 10.0, 7.0, 0.5)
        energie = st.number_input("Énergie (0-10)", 0.0, 10.0, 7.0, 0.5)
        humeur = st.number_input("Humeur (0-10)", 0.0, 10.0, 7.0, 0.5)

    if st.button("💾 Enregistrer Lifestyle"):
        # Trouver la première ligne vide
        row = None
        for r in range(2, ws.max_row + 2):
            if ws.cell(row=r, column=1).value is None:
                row = r
                break
        if row is None:
            row = ws.max_row + 1

        # On convertit une bonne fois pour toutes en float
        s = float(sommeil)
        h = float(hydrat)
        n = float(nutri)
        stv = float(stress)
        c = float(conc)
        e = float(energie)
        hm = float(humeur)

        # Écriture brute
        ws.cell(row=row, column=1).value = jour
        ws.cell(row=row, column=2).value = s
        ws.cell(row=row, column=3).value = h
        ws.cell(row=row, column=4).value = n
        ws.cell(row=row, column=5).value = stv
        ws.cell(row=row, column=6).value = c
        ws.cell(row=row, column=7).value = e
        ws.cell(row=row, column=8).value = hm

        # 🔥 Calcul Readiness en Python (0–100)
        score_pos = (s + h + n + c + e + hm) / 6.0
        score_stress = 10.0 - stv
        readiness10 = 0.7 * score_pos + 0.3 * score_stress   # 0–10
        readiness100 = round(readiness10 * 10)                # 0–100

        # Colonne 9 = Readiness
        ws.cell(row=row, column=9).value = readiness100

        wb.save(data_path)
        st.success(f"Lifestyle jour {jour} enregistré. Readiness = {readiness100}/100")


def page_force():
    st.header("🏋️‍♂️ Séance Force")

    wb, data_path = get_excel_file()
    ws = wb["Données Force"]
    headers = {ws.cell(row=1, column=c).value: c for c in range(1, ws.max_column + 1)}

    session = st.number_input("Numéro de séance", min_value=1, step=1, value=1)
    row = find_row_by_session(ws, int(session))

    st.write("Remplis uniquement les exercices faits. Laisse vide pour ignorer.")

    exos = [
        ("Squat", "Squat (kg)", "Squat (reps)"),
        ("Front Squat", "Front Squat (kg)", "Front Squat (reps)"),
        ("Bench", "Bench (kg)", "Bench (reps)"),
        ("Deadlift", "Deadlift (kg)", "Deadlift (reps)"),
        ("Overhead Press", "OHP (kg)", "OHP (reps)"),
        ("Rowing", "Rowing (kg)", "Rowing (reps)"),
        ("Traction lestée", "Traction Lestée (kg)", "Traction Lestée (reps)")
    ]

    inputs = []
    for label, col_kg, col_rep in exos:
        col1, col2, col3 = st.columns([2, 1, 1])
        with col1:
            st.markdown(f"**{label}**")
        with col2:
            kg = st.text_input(f"{label} (kg)", key=f"{label}_kg")
        with col3:
            reps = st.text_input(f"{label} (reps)", key=f"{label}_reps")
        inputs.append((col_kg, col_rep, kg, reps))

    rpe_moy_txt = st.text_input("RPE moyen de la séance (optionnel)")

    if st.button("💾 Enregistrer Séance Force"):
        for col_kg, col_rep, kg_txt, reps_txt in inputs:
            if kg_txt.strip() == "" or reps_txt.strip() == "":
                continue
            try:
                kg = float(kg_txt)
                reps = int(reps_txt)
            except ValueError:
                continue
            ckg = headers.get(col_kg)
            crep = headers.get(col_rep)
            if ckg and crep:
                ws.cell(row=row, column=ckg).value = kg
                ws.cell(row=row, column=crep).value = reps

        if rpe_moy_txt.strip() != "":
            try:
                rpe_val = float(rpe_moy_txt)
                col_rpe = headers.get("RPE Moyen (à remplir)")
                if col_rpe:
                    ws.cell(row=row, column=col_rpe).value = rpe_val
            except ValueError:
                pass

        wb.save(data_path)
        st.success(f"Séance Force {int(session)} enregistrée.")


def page_calisthenie():
    st.header("🤸 Séance Calisthénie")

    wb, data_path = get_excel_file()
    ws = wb["Données Calisthénie"]
    headers = {ws.cell(row=1, column=c).value: c for c in range(1, ws.max_column + 1)}

    session = st.number_input("Numéro de séance (même que force)", min_value=1, step=1, value=1)
    row = find_row_by_session(ws, int(session))

    champs = [
        ("HSPU (reps)", "HSPU (reps)"),
        ("MU (reps)", "MU (reps)"),
        ("Planche (sec)", "Planche (sec)"),
        ("Traction Lestée (kg)", "Traction Lestée (kg)"),
        ("Box Jump (cm)", "Box Jump (cm)")
    ]

    values = {}
    for label, col_name in champs:
        txt = st.text_input(label, key=f"cali_{label}")
        values[col_name] = txt

    if st.button("💾 Enregistrer Séance Calisthénie"):
        for col_name, txt in values.items():
            if txt.strip() == "":
                continue
            col = headers.get(col_name)
            if not col:
                continue
            try:
                val = float(txt)
                ws.cell(row=row, column=col).value = val
            except ValueError:
                pass

        wb.save(data_path)
        st.success(f"Séance Calisthénie {int(session)} enregistrée.")


def page_rpe_jour():
    st.header("🎯 RPE du jour – Charges cibles")

    wb, data_path = get_excel_file()
    ws = wb["RPE_Jour_Reps"]

    exos_list = [
        "Front Squat","Back Squat","Snatch Grip Deadlift","Bulgarian Split Squat",
        "Hack Squat","Leg Press","Leg Extension","Leg Curl","Calf Raise",
        "Bench Press","Weighted Push-ups","Dips","Handstand Push-up",
        "Military Press","Incline Fly","Lateral Raise","Triceps Extension",
        "Weighted Pull-up","Rowing","Good Morning","Muscle-up",
        "Lat Pulldown","Biceps Curl","Face Pull","Belt Squat",
        "Romanian Deadlift","Hip Thrust","Kickback","Abduction",
        "Box Jump","Pistol Squat","Farmer Walk","Burpees",
        "HSPU","MU","Planche","Traction Lestée","Pompes Diamant"
    ]

    exo = st.selectbox("Exercice", exos_list)
    charge = st.number_input("Charge (kg)", min_value=0.0, step=0.5)
    reps = st.number_input("Reps", min_value=1, step=1)

    if st.button("💾 Enregistrer dans RPE_Jour_Reps"):
        row_to_use = None
        for r in range(2, ws.max_row + 1):
            if ws.cell(row=r, column=1).value == exo:
                if ws.cell(row=r, column=2).value in (None, "") or ws.cell(row=r, column=3).value in (None, ""):
                    row_to_use = r
                    break
        if row_to_use is None:
            row_to_use = ws.max_row + 1
            ws.cell(row=row_to_use, column=1).value = exo

        ws.cell(row=row_to_use, column=2).value = float(charge)
        ws.cell(row=row_to_use, column=3).value = int(reps)

        wb.save(data_path)
        st.success(f"RPE jour enregistré pour {exo}.")

    st.markdown("---")
    st.markdown("### Aperçu des premières lignes RPE_Jour_Reps")
    try:
        df_rpe = pd.read_excel(data_path, sheet_name="RPE_Jour_Reps")
        st.dataframe(df_rpe.head(20))
    except Exception as e:
        st.warning(f"Impossible de lire la feuille RPE_Jour_Reps : {e}")


def page_dashboards():
    st.header("📊 Dashboards – Volume, 1RM, Calisthénie")

    wb, data_path = get_excel_file()

    # ====== LECTURE DONNÉES FORCE ======
    try:
        df_force = pd.read_excel(data_path, sheet_name="Données Force")
    except Exception as e:
        st.error(f"Erreur lecture Données Force : {e}")
        return

    if "Séance" not in df_force.columns:
        st.info("Pas de colonne 'Séance' dans Données Force.")
        return

    df_f = df_force.copy().sort_values("Séance")

    # On s'assure que kg & reps sont numériques
    def to_float(col):
        return pd.to_numeric(df_f.get(col), errors="coerce")

    squat_kg = to_float("Squat (kg)")
    squat_reps = to_float("Squat (reps)")
    bench_kg = to_float("Bench (kg)")
    bench_reps = to_float("Bench (reps)")
    dead_kg = to_float("Deadlift (kg)")
    dead_reps = to_float("Deadlift (reps)")
    fs_kg = to_float("Front Squat (kg)")
    fs_reps = to_float("Front Squat (reps)")
    ohp_kg = to_float("OHP (kg)")
    ohp_reps = to_float("OHP (reps)")
    row_kg = to_float("Rowing (kg)")
    row_reps = to_float("Rowing (reps)")
    pull_kg = to_float("Traction Lestée (kg)")
    pull_reps = to_float("Traction Lestée (reps)")

    # ====== 1RM EPLEY recalculé en Python ======
    def epley(kg, reps):
        return kg * (1 + reps / 30.0)

    df_f["Squat 1RM (py)"] = epley(squat_kg, squat_reps)
    df_f["Bench 1RM (py)"] = epley(bench_kg, bench_reps)
    df_f["Deadlift 1RM (py)"] = epley(dead_kg, dead_reps)

    # ====== VOLUME SESSION (kg * reps) recalculé ======
    vol_cols = [
        squat_kg * squat_reps,
        fs_kg * fs_reps,
        bench_kg * bench_reps,
        dead_kg * dead_reps,
        ohp_kg * ohp_reps,
        row_kg * row_reps,
        pull_kg * pull_reps,
    ]
    df_f["Session Volume (py)"] = sum(vol_cols)

    # ====== GRAPHIQUE VOLUME ======
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Volume par séance (calcul Python)")
        if df_f["Session Volume (py)"].notna().any():
            st.line_chart(df_f.set_index("Séance")["Session Volume (py)"])
        else:
            st.info("Aucun volume calculable (remplis au moins un exo avec kg & reps).")

    # ====== GRAPHIQUE 1RM ======
    with col2:
        st.subheader("1RM estimées Squat / Bench / Deadlift (Epley Python)")
        cols_1rm = ["Squat 1RM (py)", "Bench 1RM (py)", "Deadlift 1RM (py)"]
        if any(df_f[c].notna().any() for c in cols_1rm):
            st.line_chart(df_f.set_index("Séance")[cols_1rm])
        else:
            st.info("Aucun 1RM calculable (remplis kg & reps pour Squat / Bench / Deadlift).")

    st.markdown("---")

    # ====== CALISTHÉNIE ======
    try:
        df_cali = pd.read_excel(data_path, sheet_name="Données Calisthénie")
    except Exception as e:
        st.warning(f"Erreur lecture Données Calisthénie : {e}")
        return

    if "Séance" not in df_cali.columns:
        st.info("Pas de colonne 'Séance' dans Données Calisthénie.")
        return

    df_c = df_cali.copy().sort_values("Séance")

    # On recalcule un volume cali simple si besoin
    hspu = pd.to_numeric(df_c.get("HSPU (reps)"), errors="coerce")
    mu = pd.to_numeric(df_c.get("MU (reps)"), errors="coerce")
    planche = pd.to_numeric(df_c.get("Planche (sec)"), errors="coerce")
    t_lest = pd.to_numeric(df_c.get("Traction Lestée (kg)"), errors="coerce")
    box = pd.to_numeric(df_c.get("Box Jump (cm)"), errors="coerce")

    # pondérations arbitraires mais cohérentes avec ton système
    df_c["Calisth Volume (py)"] = (
        hspu * 10 +
        mu * 15 +
        planche * 1 +
        t_lest * 5 +
        box * 2
    )

    st.subheader("Volume Calisthénie (calcul Python)")
    if df_c["Calisth Volume (py)"].notna().any():
        st.line_chart(df_c.set_index("Séance")["Calisth Volume (py)"])
    else:
        st.info("Aucun volume calisthénie calculable pour l’instant.")


def page_reco_global():
    st.header("🧠 Synthèse & Recommandations globales")

    wb, data_path = get_excel_file(data_only=True)

    try:
        auto = wb["Auto-Séance"]
        pday = wb["Plan Jour Auto"]
        sah_ws = wb["Score Athlète Hybride"]
        life = wb["Lifestyle"]
        fat = wb["Fatigue & Readiness"]

        readiness_vals = [life.cell(row=r, column=9).value for r in range(2, life.max_row + 1)
                          if isinstance(life.cell(row=r, column=9).value, (int, float))]
        readiness_moy = sum(readiness_vals) / len(readiness_vals) if readiness_vals else None

        strain_vals = [fat.cell(row=r, column=6).value for r in range(2, fat.max_row + 1)
                       if isinstance(fat.cell(row=r, column=6).value, (int, float))]
        fatigue_moy = sum(strain_vals) / len(strain_vals) if strain_vals else None

        sah = sah_ws["F2"].value
        reco_auto = auto["C2"].value
        reco_pday = pday["D2"].value

        col1, col2 = st.columns(2)
        with col1:
            st.metric("Readiness moyen", value=round(readiness_moy, 1) if readiness_moy is not None else "N/A")
            st.metric("Fatigue moyenne (Strain)", value=round(fatigue_moy, 1) if fatigue_moy is not None else "N/A")
        with col2:
            st.metric("Score SAH", value=sah if sah is not None else "N/A")

        st.markdown("---")
        st.subheader("Séance recommandée")
        st.write(f"**Auto-Séance** : {reco_auto if reco_auto else 'N/A'}")
        st.write(f"**Plan Jour Auto** : {reco_pday if reco_pday else 'N/A'}")

    except Exception as e:
        st.error(f"Erreur lors de la lecture des recommandations : {e}")


# ======================
# MAIN
# ======================
from io import BytesIO

def page_export_debug():
    st.header("📥 Export & Debug des données Empereur")

    # 1) Aperçu Lifestyle
    st.subheader("Dernières entrées Lifestyle")
    try:
        _, data_path = get_excel_file()
        df_life = pd.read_excel(data_path, sheet_name="Lifestyle")
        st.dataframe(df_life.tail(10))
    except Exception as e:
        st.warning(f"Impossible de lire la feuille Lifestyle : {e}")

    st.markdown("---")

    # 2) Aperçu Données Force
    st.subheader("Dernières entrées Séances Force")
    try:
        df_force = pd.read_excel(data_path, sheet_name="Données Force")
        st.dataframe(df_force.tail(10))
    except Exception as e:
        st.warning(f"Impossible de lire la feuille Données Force : {e}")

    st.markdown("---")

    # 3) Téléchargement du fichier empereur_data.xlsx
    st.subheader("Télécharger le fichier de données complet")

    data_path = Path(DATA_FILE)
    if not data_path.exists():
        st.info("Aucun fichier empereur_data.xlsx trouvé pour l'instant (enregistre au moins une donnée).")
        return

    with open(data_path, "rb") as f:
        binary = f.read()

    st.download_button(
        label="📥 Télécharger empereur_data.xlsx",
        data=binary,
        file_name="empereur_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

PAGES = {
    "Lifestyle": page_lifestyle,
    "Séance Force": page_force,
    "Séance Calisthénie": page_calisthenie,
    "RPE du jour": page_rpe_jour,
    "Dashboards Volume / 1RM / Cali": page_dashboards,
    "PR & SAH": page_pr_sah,
    "Planning (Annuel / Mésocycles)": page_planning,
    "Synthèse & Recos Globales": page_reco_global,
    "Export / Debug": page_export_debug,
}

def main():
    st.set_page_config(page_title="Système Empereur", layout="wide")
    st.sidebar.title("Système d'entraînement de l'Empereur")
    choix = st.sidebar.radio("Navigation", list(PAGES.keys()))
    st.sidebar.markdown("---")
    st.sidebar.write(f"Modèle : `{TEMPLATE_FILE}`")
    st.sidebar.write(f"Données actives : `{DATA_FILE}`")

    PAGES[choix]()


if __name__ == "__main__":
    main()
