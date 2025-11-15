import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from pathlib import Path
import shutil
import numpy as np
from io import BytesIO

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
# CALCULS PYTHON : CHARGE, FATIGUE, SAH
# ======================

def compute_session_metrics(data_path: Path):
    """
    Recalcule la charge par séance (Force + Cali) et retourne :
    - df_sessions : DataFrame avec colonnes ['Séance', 'Load']
    """
    # Force
    try:
        df_force = pd.read_excel(data_path, sheet_name="Données Force")
    except Exception:
        df_force = None

    # Cali
    try:
        df_cali = pd.read_excel(data_path, sheet_name="Données Calisthénie")
    except Exception:
        df_cali = None

    if df_force is None or "Séance" not in df_force.columns:
        return None

    df_f = df_force.copy()
    df_f["Séance"] = pd.to_numeric(df_f["Séance"], errors="coerce")
    df_f = df_f.dropna(subset=["Séance"]).sort_values("Séance")

    def to_float(df, col):
        return pd.to_numeric(df.get(col), errors="coerce")

    squat_kg = to_float(df_f, "Squat (kg)")
    squat_reps = to_float(df_f, "Squat (reps)")
    fs_kg = to_float(df_f, "Front Squat (kg)")
    fs_reps = to_float(df_f, "Front Squat (reps)")
    bench_kg = to_float(df_f, "Bench (kg)")
    bench_reps = to_float(df_f, "Bench (reps)")
    dead_kg = to_float(df_f, "Deadlift (kg)")
    dead_reps = to_float(df_f, "Deadlift (reps)")
    ohp_kg = to_float(df_f, "OHP (kg)")
    ohp_reps = to_float(df_f, "OHP (reps)")
    row_kg = to_float(df_f, "Rowing (kg)")
    row_reps = to_float(df_f, "Rowing (reps)")
    pull_kg = to_float(df_f, "Traction Lestée (kg)")
    pull_reps = to_float(df_f, "Traction Lestée (reps)")

    vol_force = (
        squat_kg * squat_reps +
        fs_kg * fs_reps +
        bench_kg * bench_reps +
        dead_kg * dead_reps +
        ohp_kg * ohp_reps +
        row_kg * row_reps +
        pull_kg * pull_reps
    )

    df_f["Load"] = vol_force

    if df_cali is not None and "Séance" in df_cali.columns:
        df_c = df_cali.copy()
        df_c["Séance"] = pd.to_numeric(df_c["Séance"], errors="coerce")
        df_c = df_c.dropna(subset=["Séance"])

        hspu = to_float(df_c, "HSPU (reps)")
        mu = to_float(df_c, "MU (reps)")
        planche = to_float(df_c, "Planche (sec)")
        t_lest = to_float(df_c, "Traction Lestée (kg)")
        box = to_float(df_c, "Box Jump (cm)")

        df_c["Calisth Volume (py)"] = (
            hspu * 10 +
            mu * 15 +
            planche * 1 +
            t_lest * 5 +
            box * 2
        )

        df_merged = df_f.merge(
            df_c[["Séance", "Calisth Volume (py)"]],
            on="Séance",
            how="left"
        )
        df_merged["Calisth Volume (py)"] = df_merged["Calisth Volume (py)"].fillna(0)
        df_merged["Load"] = df_merged["Load"].fillna(0) + df_merged["Calisth Volume (py)"]
        df_sessions = df_merged[["Séance", "Load"]]
    else:
        df_f["Load"] = df_f["Load"].fillna(0)
        df_sessions = df_f[["Séance", "Load"]]

    df_sessions = df_sessions[df_sessions["Load"] > 0]

    if df_sessions.empty:
        return None

    return df_sessions.sort_values("Séance")


def compute_fatigue_metrics(data_path: Path, window: int = 7):
    """
    Calcule Monotony et Strain sur les 'window' dernières séances.
    """
    df_sessions = compute_session_metrics(data_path)
    if df_sessions is None or df_sessions.empty:
        return None, None, None

    loads = df_sessions["Load"].values
    if len(loads) >= window:
        loads_window = loads[-window:]
    else:
        loads_window = loads

    mean_load = float(np.mean(loads_window))
    std_load = float(np.std(loads_window)) if len(loads_window) > 1 else 0.0

    if std_load == 0:
        monotony = 0.0
    else:
        monotony = mean_load / std_load

    strain = mean_load * monotony

    return mean_load, monotony, strain


def safe_nanmax(arr):
    arr = np.array(arr, dtype=float)
    if arr.size == 0 or np.isnan(arr).all():
        return 0.0
    return float(np.nanmax(arr))


def compute_sah_score(data_path: Path):
    """
    Calcule un Score Athlète Hybride (SAH) simplifié basé sur :
    - Meilleurs 1RM Squat / Bench / Deadlift (Epley Python)
    - Meilleurs HSPU / MU / Traction lestée / Planche / Box jump
    Retourne (SAH, dict détails)
    """
    try:
        df_force = pd.read_excel(data_path, sheet_name="Données Force")
    except Exception:
        df_force = None

    try:
        df_cali = pd.read_excel(data_path, sheet_name="Données Calisthénie")
    except Exception:
        df_cali = None

    if df_force is None:
        return None, {}

    df_f = df_force.copy()

    def to_float(df, col):
        return pd.to_numeric(df.get(col), errors="coerce")

    squat_kg = to_float(df_f, "Squat (kg)")
    squat_reps = to_float(df_f, "Squat (reps)")
    bench_kg = to_float(df_f, "Bench (kg)")
    bench_reps = to_float(df_f, "Bench (reps)")
    dead_kg = to_float(df_f, "Deadlift (kg)")
    dead_reps = to_float(df_f, "Deadlift (reps)")

    def epley(kg, reps):
        return kg * (1 + reps / 30.0)

    squat_1rm = epley(squat_kg, squat_reps)
    bench_1rm = epley(bench_kg, bench_reps)
    dead_1rm = epley(dead_kg, dead_reps)

    best_squat = safe_nanmax(squat_1rm)
    best_bench = safe_nanmax(bench_1rm)
    best_dead = safe_nanmax(dead_1rm)

    sq_target = 220.0
    bp_target = 160.0
    dl_target = 260.0

    str_squat = min(best_squat / sq_target, 1.2) if sq_target > 0 else 0
    str_bench = min(best_bench / bp_target, 1.2) if bp_target > 0 else 0
    str_dead = min(best_dead / dl_target, 1.2) if dl_target > 0 else 0

    strength_score = np.mean([str_squat, str_bench, str_dead]) * 100.0

    details = {
        "Squat1RM": round(best_squat, 1),
        "Bench1RM": round(best_bench, 1),
        "Dead1RM": round(best_dead, 1),
        "StrengthScore": round(strength_score, 1),
    }

    skill_score = None

    if df_cali is not None:
        df_c = df_cali.copy()
        hspu = to_float(df_c, "HSPU (reps)")
        mu = to_float(df_c, "MU (reps)")
        planche = to_float(df_c, "Planche (sec)")
        t_lest = to_float(df_c, "Traction Lestée (kg)")
        box = to_float(df_c, "Box Jump (cm)")

        best_hspu = safe_nanmax(hspu)
        best_mu = safe_nanmax(mu)
        best_planche = safe_nanmax(planche)
        best_tlest = safe_nanmax(t_lest)
        best_box = safe_nanmax(box)

        details.update({
            "HSPU": best_hspu,
            "MU": best_mu,
            "Planche_sec": best_planche,
            "TractionLestee": best_tlest,
            "BoxJump_cm": best_box,
        })

        hspu_target = 20.0
        mu_target = 10.0
        planche_target = 20.0
        tlest_target = 80.0
        box_target = 120.0

        s_hspu = min(best_hspu / hspu_target, 1.2) if hspu_target > 0 else 0
        s_mu = min(best_mu / mu_target, 1.2) if mu_target > 0 else 0
        s_planche = min(best_planche / planche_target, 1.2) if planche_target > 0 else 0
        s_tlest = min(best_tlest / tlest_target, 1.2) if tlest_target > 0 else 0
        s_box = min(best_box / box_target, 1.2) if box_target > 0 else 0

        skill_score = np.mean([s_hspu, s_mu, s_planche, s_tlest, s_box]) * 100.0
        details["SkillScore"] = round(skill_score, 1)

    if skill_score is not None:
        sah = 0.5 * strength_score + 0.5 * skill_score
    else:
        sah = strength_score

    sah = float(np.clip(sah, 0, 100))
    details["SAH"] = round(sah, 1)
    return sah, details


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

        s = float(sommeil)
        h = float(hydrat)
        n = float(nutri)
        stv = float(stress)
        c = float(conc)
        e = float(energie)
        hm = float(humeur)

        ws.cell(row=row, column=1).value = jour
        ws.cell(row=row, column=2).value = s
        ws.cell(row=row, column=3).value = h
        ws.cell(row=row, column=4).value = n
        ws.cell(row=row, column=5).value = stv
        ws.cell(row=row, column=6).value = c
        ws.cell(row=row, column=7).value = e
        ws.cell(row=row, column=8).value = hm

        # Nommer la colonne 9 "Readiness" si pas déjà fait
        if ws.cell(row=1, column=9).value in (None, ""):
            ws.cell(row=1, column=9).value = "Readiness"

        score_pos = (s + h + n + c + e + hm) / 6.0
        score_stress = 10.0 - stv
        readiness10 = 0.7 * score_pos + 0.3 * score_stress
        readiness100 = round(readiness10 * 10)

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

    try:
        df_force = pd.read_excel(data_path, sheet_name="Données Force")
    except Exception as e:
        st.error(f"Erreur lecture Données Force : {e}")
        return

    if "Séance" not in df_force.columns:
        st.info("Pas de colonne 'Séance' dans Données Force.")
        return

    df_f = df_force.copy().sort_values("Séance")

    def to_float(df, col):
        return pd.to_numeric(df.get(col), errors="coerce")

    squat_kg = to_float(df_f, "Squat (kg)")
    squat_reps = to_float(df_f, "Squat (reps)")
    bench_kg = to_float(df_f, "Bench (kg)")
    bench_reps = to_float(df_f, "Bench (reps)")
    dead_kg = to_float(df_f, "Deadlift (kg)")
    dead_reps = to_float(df_f, "Deadlift (reps)")
    fs_kg = to_float(df_f, "Front Squat (kg)")
    fs_reps = to_float(df_f, "Front Squat (reps)")
    ohp_kg = to_float(df_f, "OHP (kg)")
    ohp_reps = to_float(df_f, "OHP (reps)")
    row_kg = to_float(df_f, "Rowing (kg)")
    row_reps = to_float(df_f, "Rowing (reps)")
    pull_kg = to_float(df_f, "Traction Lestée (kg)")
    pull_reps = to_float(df_f, "Traction Lestée (reps)")

    def epley(kg, reps):
        return kg * (1 + reps / 30.0)

    df_f["Squat 1RM (py)"] = epley(squat_kg, squat_reps)
    df_f["Bench 1RM (py)"] = epley(bench_kg, bench_reps)
    df_f["Deadlift 1RM (py)"] = epley(dead_kg, dead_reps)

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

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Volume par séance (calcul Python)")
        if df_f["Session Volume (py)"].notna().any():
            st.line_chart(df_f.set_index("Séance")["Session Volume (py)"])
        else:
            st.info("Aucun volume calculable (remplis au moins un exo avec kg & reps).")

    with col2:
        st.subheader("1RM estimées Squat / Bench / Deadlift (Epley Python)")
        cols_1rm = ["Squat 1RM (py)", "Bench 1RM (py)", "Deadlift 1RM (py)"]
        if any(df_f[c].notna().any() for c in cols_1rm):
            st.line_chart(df_f.set_index("Séance")[cols_1rm])
        else:
            st.info("Aucun 1RM calculable (remplis kg & reps pour Squat / Bench / Deadlift).")

    st.markdown("---")

    try:
        df_cali = pd.read_excel(data_path, sheet_name="Données Calisthénie")
    except Exception as e:
        st.warning(f"Erreur lecture Données Calisthénie : {e}")
        return

    if "Séance" not in df_cali.columns:
        st.info("Pas de colonne 'Séance' dans Données Calisthénie.")
        return

    df_c = df_cali.copy().sort_values("Séance")

    hspu = pd.to_numeric(df_c.get("HSPU (reps)"), errors="coerce")
    mu = pd.to_numeric(df_c.get("MU (reps)"), errors="coerce")
    planche = pd.to_numeric(df_c.get("Planche (sec)"), errors="coerce")
    t_lest = pd.to_numeric(df_c.get("Traction Lestée (kg)"), errors="coerce")
    box = pd.to_numeric(df_c.get("Box Jump (cm)"), errors="coerce")

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


def page_pr_sah():
    st.header("🏆 PR & Score Athlète Hybride")

    wb, data_path = get_excel_file(data_only=True)

    try:
        df_pr = pd.read_excel(data_path, sheet_name="PR Automatiques")
        st.subheader("PR Automatiques (Excel)")
        st.dataframe(df_pr)
    except Exception as e:
        st.warning(f"Erreur lecture PR Automatiques : {e}")

    sah, details = compute_sah_score(data_path)
    st.subheader("Score Athlète Hybride (calcul Python)")
    if sah is not None:
        col1, col2 = st.columns(2)
        with col1:
            st.metric("SAH", value=round(sah, 1))
        with col2:
            st.write("Détails composantes :")
            st.json(details)
    else:
        st.info("Pas encore assez de données (force / calisthénie) pour calculer un SAH.")


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

    # READINESS MOYEN
    try:
        df_life = pd.read_excel(data_path, sheet_name="Lifestyle")
        if "Readiness" in df_life.columns:
            readiness_col = "Readiness"
        elif df_life.shape[1] >= 9:
            readiness_col = df_life.columns[8]
        else:
            readiness_col = None

        if readiness_col:
            vals = pd.to_numeric(df_life[readiness_col], errors="coerce").dropna()
            readiness_moy = float(vals.mean()) if not vals.empty else None
        else:
            readiness_moy = None
    except Exception:
        readiness_moy = None

    mean_load, monotony, strain = compute_fatigue_metrics(data_path)
    sah, sah_details = compute_sah_score(data_path)

    col1, col2 = st.columns(2)
    with col1:
        st.metric(
            "Readiness moyen",
            value=round(readiness_moy, 1) if readiness_moy is not None else "N/A"
        )
        if mean_load is not None:
            st.metric("Charge moyenne (7 dernières séances)", value=int(mean_load))
        else:
            st.metric("Charge moyenne", value="N/A")

    with col2:
        if monotony is not None:
            st.metric("Monotony", value=round(monotony, 2))
        else:
            st.metric("Monotony", value="N/A")
        if strain is not None:
            st.metric("Strain", value=int(strain))
        else:
            st.metric("Strain", value="N/A")

    st.markdown("---")
    st.subheader("Score Athlète Hybride (Python)")
    if sah is not None:
        st.metric("SAH", value=round(sah, 1))
        with st.expander("Détails SAH"):
            st.json(sah_details)
    else:
        st.info("Pas encore assez de données pour calculer un SAH.")

    st.markdown("---")
    st.subheader("Séance recommandée (logique à définir)")
    st.write("🔜 Prochaine étape : lier Readiness + Strain pour proposer automatiquement Heavy / Volume / Skill / Rest.")


def page_export_debug():
    st.header("📥 Export & Debug des données Empereur")

    wb, data_path = get_excel_file()

    # 1) Lifestyle
    st.subheader("Lifestyle – dernières entrées")
    try:
        df_life = pd.read_excel(data_path, sheet_name="Lifestyle")
        st.dataframe(df_life.tail(10))
    except Exception as e:
        st.warning(f"Impossible de lire la feuille Lifestyle : {e}")

    st.markdown("---")

    # 2) Séances Force
    st.subheader("Séances Force – dernières entrées")
    try:
        df_force = pd.read_excel(data_path, sheet_name="Données Force")
        st.dataframe(df_force.tail(10))
    except Exception as e:
        st.warning(f"Impossible de lire la feuille Données Force : {e}")

    st.markdown("---")

    # 3) Séances Calisthénie
    st.subheader("Séances Calisthénie – dernières entrées")
    try:
        df_cali = pd.read_excel(data_path, sheet_name="Données Calisthénie")
        st.dataframe(df_cali.tail(10))
    except Exception as e:
        st.warning(f"Impossible de lire la feuille Données Calisthénie : {e}")

    st.markdown("---")

    # 4) RPE Jour
    st.subheader("RPE_Jour_Reps – dernières entrées")
    try:
        df_rpe = pd.read_excel(data_path, sheet_name="RPE_Jour_Reps")
        st.dataframe(df_rpe.tail(10))
    except Exception as e:
        st.warning(f"Impossible de lire la feuille RPE_Jour_Reps : {e}")

    st.markdown("---")
    st.subheader("Télécharger le fichier de données complet")

    data_path = Path(DATA_FILE)
    if not data_path.exists():
        st.info("Aucun fichier empereur_data.xlsx trouvé pour l'instant (enregistre d'abord des données).")
    else:
        with open(data_path, "rb") as f:
            binary = f.read()

        st.download_button(
            label="📥 Télécharger empereur_data.xlsx",
            data=binary,
            file_name="empereur_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.markdown("---")
    st.subheader("♻️ Réinitialiser toutes les données")

    st.warning(
        "⚠️ Cette action supprime **toutes** les données actuelles (Lifestyle, Force, Calisthénie, RPE, etc.) "
        "et recrée un fichier vierge à partir du modèle."
    )

    if st.button("🔴 Réinitialiser empereur_data.xlsx"):
        if data_path.exists():
            data_path.unlink()
            st.success(
                "Toutes les données ont été réinitialisées. "
                "La prochaine utilisation de l'app recréera un fichier vierge à partir du modèle."
            )
        else:
            st.info("Aucun fichier de données à supprimer.")




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
