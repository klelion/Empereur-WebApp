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
        row = None
        for r in range(2, ws.max_row + 2):
            if ws.cell(row=r, column=1).value is None:
                row = r
                break
        if row is None:
            row = ws.max_row + 1

        ws.cell(row=row, column=1).value = jour
        ws.cell(row=row, column=2).value = float(sommeil)
        ws.cell(row=row, column=3).value = float(hydrat)
        ws.cell(row=row, column=4).value = float(nutri)
        ws.cell(row=row, column=5).value = float(stress)
        ws.cell(row=row, column=6).value = float(conc)
        ws.cell(row=row, column=7).value = float(energie)
        ws.cell(row=row, column=8).value = float(humeur)

        wb.save(data_path)
        st.success("Lifestyle enregistré.")


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
    try:
        df_force = pd.read_excel(data_path, sheet_name="Données Force")
    except Exception as e:
        st.error(f"Erreur lecture Données Force : {e}")
        return

    if "Séance" in df_force.columns:
        df_f = df_force.sort_values("Séance")

        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Volume par séance")
            if "Session Volume (kg·reps)" in df_f.columns:
                st.line_chart(df_f.set_index("Séance")["Session Volume (kg·reps)"])
            else:
                st.info("Pas de colonne 'Session Volume (kg·reps)'.")

        with col2:
            st.subheader("1RM Estimées (Squat / Bench / Deadlift)")
            cols_1rm = [c for c in ["Squat 1RM Est","Bench 1RM Est","Deadlift 1RM Est"] if c in df_f.columns]
            if cols_1rm:
                st.line_chart(df_f.set_index("Séance")[cols_1rm])
            else:
                st.info("Pas de colonnes 1RM trouvées.")
    else:
        st.info("Pas de colonne 'Séance' dans Données Force.")

    st.markdown("---")

    try:
        df_cali = pd.read_excel(data_path, sheet_name="Données Calisthénie")
        st.subheader("Volume Calisthénie")
        if "Séance" in df_cali.columns and "Calisth. Volume (unités)" in df_cali.columns:
            st.line_chart(df_cali.set_index("Séance")["Calisth. Volume (unités)"])
        else:
            st.info("Colonnes 'Séance' ou 'Calisth. Volume (unités)' manquantes.")
    except Exception as e:
        st.warning(f"Erreur lecture Données Calisthénie : {e}")


def page_pr_sah():
    st.header("🏆 PR & Score Athlète Hybride")

    wb, data_path = get_excel_file(data_only=True)

    try:
        df_pr = pd.read_excel(data_path, sheet_name="PR Automatiques")
        st.subheader("PR Automatiques")
        st.dataframe(df_pr)
    except Exception as e:
        st.warning(f"Erreur lecture PR Automatiques : {e}")

    try:
        ws_sah = wb["Score Athlète Hybride"]
        sah = ws_sah["F2"].value
        st.subheader("Score Athlète Hybride (SAH)")
        st.metric("SAH", value=sah if sah is not None else "N/A")
    except Exception as e:
        st.warning(f"Impossible de lire le score SAH : {e}")


def page_planning():
    st.header("📅 Planning – Plan Annuel & Mésocycles")

    wb, data_path = get_excel_file()

    col1, col2 = st.columns(2)
    try:
        df_annuel = pd.read_excel(data_path, sheet_name="Plan Annuel")
        with col1:
            st.subheader("Plan Annuel")
            st.dataframe(df_annuel)
    except Exception as e:
        st.warning(f"Erreur lecture Plan Annuel : {e}")

    try:
        df_meso = pd.read_excel(data_path, sheet_name="Mésocycle-Type")
        with col2:
            st.subheader("Mésocycle-Type")
            st.dataframe(df_meso)
    except Exception as e:
        st.warning(f"Erreur lecture Mésocycle-Type : {e}")

    st.markdown("---")
    try:
        df_auto_meso = pd.read_excel(data_path, sheet_name="Auto-Mesocycles")
        st.subheader("Auto-Mesocycles")
        st.dataframe(df_auto_meso)
    except Exception as e:
        st.warning(f"Erreur lecture Auto-Mesocycles : {e}")


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

PAGES = {
    "Lifestyle": page_lifestyle,
    "Séance Force": page_force,
    "Séance Calisthénie": page_calisthenie,
    "RPE du jour": page_rpe_jour,
    "Dashboards Volume / 1RM / Cali": page_dashboards,
    "PR & SAH": page_pr_sah,
    "Planning (Annuel / Mésocycles)": page_planning,
    "Synthèse & Recos Globales": page_reco_global,
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
