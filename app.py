import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from pathlib import Path
import shutil
import numpy as np

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
# CALCULS PYTHON : CHARGE, FATIGUE, SAH V2, AUTO-SEANCE
# ======================

def _to_float(df, col):
    return pd.to_numeric(df.get(col), errors="coerce")


def compute_session_metrics(data_path: Path):
    """
    Recalcule la charge par séance (Force + Cali) et retourne :
    DataFrame ['Séance', 'Load', 'Load_Force', 'Load_Cali']
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

    squat_kg = _to_float(df_f, "Squat (kg)")
    squat_reps = _to_float(df_f, "Squat (reps)")
    fs_kg = _to_float(df_f, "Front Squat (kg)")
    fs_reps = _to_float(df_f, "Front Squat (reps)")
    bench_kg = _to_float(df_f, "Bench (kg)")
    bench_reps = _to_float(df_f, "Bench (reps)")
    dead_kg = _to_float(df_f, "Deadlift (kg)")
    dead_reps = _to_float(df_f, "Deadlift (reps)")
    ohp_kg = _to_float(df_f, "OHP (kg)")
    ohp_reps = _to_float(df_f, "OHP (reps)")
    row_kg = _to_float(df_f, "Rowing (kg)")
    row_reps = _to_float(df_f, "Rowing (reps)")
    pull_kg = _to_float(df_f, "Traction Lestée (kg)")
    pull_reps = _to_float(df_f, "Traction Lestée (reps)")

    vol_force = (
        squat_kg * squat_reps +
        fs_kg * fs_reps +
        bench_kg * bench_reps +
        dead_kg * dead_reps +
        ohp_kg * ohp_reps +
        row_kg * row_reps +
        pull_kg * pull_reps
    )

    df_f["Load_Force"] = vol_force.fillna(0)

    # Cali
    load_cali_series = None
    if df_cali is not None and "Séance" in df_cali.columns:
        df_c = df_cali.copy()
        df_c["Séance"] = pd.to_numeric(df_c["Séance"], errors="coerce")
        df_c = df_c.dropna(subset=["Séance"])

        hspu = _to_float(df_c, "HSPU (reps)")
        mu = _to_float(df_c, "MU (reps)")
        planche = _to_float(df_c, "Planche (sec)")
        t_lest = _to_float(df_c, "Traction Lestée (kg)")
        box = _to_float(df_c, "Box Jump (cm)")

        df_c["Load_Cali"] = (
            hspu * 10 +
            mu * 15 +
            planche * 1 +
            t_lest * 5 +
            box * 2
        ).fillna(0)

        df_merged = df_f.merge(
            df_c[["Séance", "Load_Cali"]],
            on="Séance",
            how="left"
        )
        df_merged["Load_Cali"] = df_merged["Load_Cali"].fillna(0)
        df_merged["Load"] = df_merged["Load_Force"] + df_merged["Load_Cali"]
        df_sessions = df_merged[["Séance", "Load", "Load_Force", "Load_Cali"]]
    else:
        df_f["Load_Cali"] = 0.0
        df_f["Load"] = df_f["Load_Force"]
        df_sessions = df_f[["Séance", "Load", "Load_Force", "Load_Cali"]]

    df_sessions = df_sessions[df_sessions["Load"] > 0]

    if df_sessions.empty:
        return None

    return df_sessions.sort_values("Séance")


def compute_fatigue_metrics(data_path: Path, window: int = 7):
    """
    Calcule Charge moyenne, Monotony et Strain sur les 'window' dernières séances.
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


def compute_sah_v2(data_path: Path):
    """
    SAH V2 – Score Athlète Hybride 0–100 basé sur :
    - StrengthIndex : Squat / Bench / Deadlift
    - SkillIndex    : HSPU / MU / Planche / Traction lestée
    - PowerIndex    : Box jump + charge calisthénique
    Poids : Strength 40%, Skill 40%, Power 20%.
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

    squat_kg = _to_float(df_f, "Squat (kg)")
    squat_reps = _to_float(df_f, "Squat (reps)")
    bench_kg = _to_float(df_f, "Bench (kg)")
    bench_reps = _to_float(df_f, "Bench (reps)")
    dead_kg = _to_float(df_f, "Deadlift (kg)")
    dead_reps = _to_float(df_f, "Deadlift (reps)")

    def epley(kg, reps):
        return kg * (1 + reps / 30.0)

    squat_1rm = epley(squat_kg, squat_reps)
    bench_1rm = epley(bench_kg, bench_reps)
    dead_1rm = epley(dead_kg, dead_reps)

    best_squat = safe_nanmax(squat_1rm)
    best_bench = safe_nanmax(bench_1rm)
    best_dead = safe_nanmax(dead_1rm)

    # Cibles "empereur hybride" (à affiner à terme)
    sq_target = 220.0
    bp_target = 160.0
    dl_target = 260.0

    str_squat = min(best_squat / sq_target, 1.3) if sq_target > 0 else 0
    str_bench = min(best_bench / bp_target, 1.3) if bp_target > 0 else 0
    str_dead = min(best_dead / dl_target, 1.3) if dl_target > 0 else 0

    strength_index = float(np.mean([str_squat, str_bench, str_dead]) * 100.0)

    details = {
        "Squat1RM": round(best_squat, 1),
        "Bench1RM": round(best_bench, 1),
        "Dead1RM": round(best_dead, 1),
        "StrengthIndex": round(strength_index, 1),
    }

    skill_index = 0.0
    power_index = 0.0

    if df_cali is not None:
        df_c = df_cali.copy()
        hspu = _to_float(df_c, "HSPU (reps)")
        mu = _to_float(df_c, "MU (reps)")
        planche = _to_float(df_c, "Planche (sec)")
        t_lest = _to_float(df_c, "Traction Lestée (kg)")
        box = _to_float(df_c, "Box Jump (cm)")

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

        # Cibles calisthéniques
        hspu_target = 20.0
        mu_target = 10.0
        planche_target = 20.0   # secondes
        tlest_target = 80.0     # +80kg
        box_target = 120.0      # cm

        s_hspu = min(best_hspu / hspu_target, 1.3) if hspu_target > 0 else 0
        s_mu = min(best_mu / mu_target, 1.3) if mu_target > 0 else 0
        s_planche = min(best_planche / planche_target, 1.3) if planche_target > 0 else 0
        s_tlest = min(best_tlest / tlest_target, 1.3) if tlest_target > 0 else 0
        s_box_skill = min(best_box / box_target, 1.3) if box_target > 0 else 0

        # Skill = contrôle + complexité
        skill_index = float(np.mean([s_hspu, s_mu, s_planche, s_tlest]) * 100.0)
        # Power = box jump + traction lestée
        power_index = float(np.mean([s_tlest, s_box_skill]) * 100.0)

        details["SkillIndex"] = round(skill_index, 1)
        details["PowerIndex"] = round(power_index, 1)

    # Pondération SAH V2
    sah_components = []
    weights = []

    sah_components.append(strength_index)
    weights.append(0.4)

    sah_components.append(skill_index)
    weights.append(0.4)

    sah_components.append(power_index)
    weights.append(0.2)

    sah_v2 = float(np.average(sah_components, weights=weights))
    sah_v2 = float(np.clip(sah_v2, 0, 100))

    details["SAH_V2"] = round(sah_v2, 1)
    return sah_v2, details


def classify_skill_level(skill_index: float):
    """
    Classe le niveau Skill en catégories.
    """
    if skill_index is None:
        return "Inconnu"
    if skill_index < 30:
        return "Débutant"
    if skill_index < 60:
        return "Intermédiaire"
    if skill_index < 85:
        return "Avancé"
    return "Élite"


def get_latest_readiness(data_path: Path):
    try:
        df_life = pd.read_excel(data_path, sheet_name="Lifestyle")
    except Exception:
        return None

    if "Readiness" in df_life.columns:
        col = "Readiness"
    elif df_life.shape[1] >= 9:
        col = df_life.columns[8]
    else:
        return None

    vals = pd.to_numeric(df_life[col], errors="coerce").dropna()
    if vals.empty:
        return None
    return float(vals.iloc[-1])


def get_last_session_info(data_path: Path):
    df_sessions = compute_session_metrics(data_path)
    if df_sessions is None or df_sessions.empty:
        return None

    last = df_sessions.iloc[-1]
    load = float(last["Load"])
    load_force = float(last["Load_Force"])
    load_cali = float(last["Load_Cali"])

    if load_force > load_cali * 1.3:
        last_type = "Force"
    elif load_cali > load_force * 1.3:
        last_type = "Calisthénie"
    else:
        last_type = "Mixte"

    return {
        "Séance": int(last["Séance"]),
        "Load": load,
        "Load_Force": load_force,
        "Load_Cali": load_cali,
        "Type": last_type,
    }


def compute_auto_seance_recommendation(
    data_path: Path,
    block_focus: str
):
    """
    Moteur Auto-Séance intelligent.
    Utilise Readiness, Strain, Skill, dernière séance, charge 7j.
    Retourne un dict avec :
    - session_type
    - focus
    - intensity
    - volume_mod
    - rpe_target
    - notes
    - structure_suggestion (liste de points)
    """
    readiness = get_latest_readiness(data_path)
    mean_load, monotony, strain = compute_fatigue_metrics(data_path)
    sah_v2, details = compute_sah_v2(data_path)
    last_info = get_last_session_info(data_path)

    skill_index = details.get("SkillIndex", 0.0)
    strength_index = details.get("StrengthIndex", 0.0)
    skill_level = classify_skill_level(skill_index)

    # Défauts si pas de données
    if readiness is None:
        readiness = 50.0
    if mean_load is None:
        mean_load = 0.0
    if monotony is None:
        monotony = 0.0
    if strain is None:
        strain = 0.0

    # Catégorisation readiness
    if readiness >= 70:
        readiness_zone = "High"
    elif readiness >= 40:
        readiness_zone = "Medium"
    else:
        readiness_zone = "Low"

    # Catégorisation strain
    if strain >= 25000:
        strain_zone = "High"
    elif strain >= 10000:
        strain_zone = "Medium"
    else:
        strain_zone = "Low"

    # Logique de base : type de séance en fonction de readiness / strain / block
    session_type = ""
    focus = ""
    intensity = ""
    volume_mod = ""
    rpe_target = ""
    notes = []
    structure = []

    # Pour simplifier, on mappe block_focus en priorités
    if block_focus == "Force maximale":
        primary = "Force"
    elif block_focus == "Hypertrophie / Volume":
        primary = "Volume"
    elif block_focus == "Skill / Calisthénie":
        primary = "Skill"
    elif block_focus == "Puissance / Explosivité":
        primary = "Power"
    else:  # Déload / Gestion fatigue
        primary = "Deload"

    # Ajustement spécifique si très fatigué
    if readiness_zone == "Low" or strain_zone == "High":
        # Fatigue importante : on privilégie Recovery / Skill propre
        if primary == "Deload":
            session_type = "Recovery / Off"
            focus = "Récupération globale"
            intensity = "Très basse"
            volume_mod = "20–40% du volume habituel"
            rpe_target = "RPE 5–6 max"
            notes.append("Fatigue ou strain élevés : privilégier la récupération active.")
            structure = [
                "20–30 min mobilité totale (hanches, épaules, colonne)",
                "10–20 min marche ou cardio très léger",
                "Travail technique très propre : handstand hold, supports, respiration",
                "Sauna / bain chaud / automassage si possible"
            ]
        else:
            session_type = "Skill / Recovery"
            focus = "Technique + Calisthénie propre + mobilité"
            intensity = "Basse à modérée"
            volume_mod = "40–60% du volume habituel"
            rpe_target = "RPE 6–7"
            notes.append("Readiness bas ou strain élevé : on garde la fréquence mais on baisse l'impact.")
            structure = [
                "Bloc skill : HSPU, MU, planche (progrès techniques, pas de grind)",
                "Volume traction / push modéré, loin de l'échec",
                "Core & gainage (planche, hollow, arch)",
                "Long travail de stretching actif / PNF en fin de séance"
            ]
    else:
        # Readiness OK / Strain gérable → on regarde le focus
        if primary == "Force":
            session_type = "Heavy Strength"
            focus = "Force lourde (bas du corps ou haut, selon rotation)"
            intensity = "Élevée"
            volume_mod = "70–90% du volume habituel"
            rpe_target = "RPE 8–9 sur les principaux mouvements"
            notes.append("Tu peux pousser lourd sur 1–3 lifts principaux.")
            structure = [
                "1–2 mouvements principaux en 3–5 séries lourdes (3–6 reps)",
                "2–3 accessoires lourds ou modérés (6–10 reps)",
                "Un peu de skill en fin si énergie (HSPU / MU)",
                "Finir par un travail léger de mobilité / respiration"
            ]
        elif primary == "Volume":
            session_type = "Hypertrophie / Volume"
            focus = "Accumulation de volume contrôlé"
            intensity = "Modérée"
            volume_mod = "90–110% du volume habituel"
            rpe_target = "RPE 7–8"
            notes.append("Objectif : congestion, volume, mais sans casser le système nerveux.")
            structure = [
                "2 mouvements de base (squat / bench / row / dips...) en 4×8–12",
                "3–4 exercices d'iso / machines (12–20 reps)",
                "Optionnel : finisher métabolique (farmer walk + burpees, etc.)",
                "Étirer les groupes très travaillés"
            ]
        elif primary == "Skill":
            session_type = "Skill Calisthénie"
            focus = "Maîtrise technique + progression sur HSPU / MU / planche"
            intensity = "Modérée"
            volume_mod = "60–80% du volume habituel"
            rpe_target = "RPE 6–8 (jamais à l'échec nerveux sur skill)"
            notes.append(f"Niveau skill actuel : {skill_level}. On consolide la technique.")
            structure = [
                "Bloc 1 : MU (explosifs, bandés si besoin, 3–5 reps par série)",
                "Bloc 2 : HSPU / handstand (négatives, holds, partiels)",
                "Bloc 3 : Planche / front lever (progressions tenues propres)",
                "Finir par du tirage / push plus simple (tractions, dips, pompes)",
                "Mobility shoulders + poignets"
            ]
        elif primary == "Power":
            session_type = "Puissance / Explosivité"
            focus = "Sauts, vitesse, intention explosive"
            intensity = "Élevée mais volume limité"
            volume_mod = "50–70% volume muscu, 100% intensité sur l'explosivité"
            rpe_target = "RPE 7–8 sur explosif, pas d'échec"
            notes.append("Objectif : système nerveux rapide, pas cramé.")
            structure = [
                "Sauts (box jumps, bounds, sauts horizontaux, 3–5 reps par série)",
                "Mouvements olympiques techniques si tu en utilises (high pull, etc.)",
                "Sprints courts ou hill sprints (si contexte adapté)",
                "Un peu de force submax (70–80% 1RM, mouvement rapide)",
                "Finir par mobilité hanches / chevilles"
            ]
        else:  # Deload
            session_type = "Deload intelligent"
            focus = "Réduction de charge, maintien technique"
            intensity = "Basse à modérée"
            volume_mod = "40–60% du volume habituel"
            rpe_target = "RPE 6–7"
            notes.append("Bloc orienté gestion fatigue / décharge.")
            structure = [
                "Même structure de séance qu'habituel mais -40% charge / volume",
                "Travail technique plus propre (tempo contrôlé, pauses)",
                "Beaucoup de mobilité / respiration en fin",
                "Sleep / nutrition prioritaires"
            ]

    # Ajustement léger selon dernière séance
    if last_info is not None:
        if last_info["Type"] == "Force" and "Heavy" in session_type:
            notes.append("Dernière séance déjà très force → surveille tes sensations sur les premiers sets.")
        if last_info["Type"] == "Calisthénie" and "Skill" in session_type:
            notes.append("Tu peux recycler certains patterns de la dernière séance en version plus propre.")

    return {
        "readiness": readiness,
        "mean_load": mean_load,
        "monotony": monotony,
        "strain": strain,
        "sah_v2": sah_v2,
        "strength_index": strength_index,
        "skill_index": skill_index,
        "power_index": details.get("PowerIndex", 0.0),
        "skill_level": skill_level,
        "last_session": last_info,
        "session_type": session_type,
        "focus": focus,
        "intensity": intensity,
        "volume_mod": volume_mod,
        "rpe_target": rpe_target,
        "notes": notes,
        "structure_suggestion": structure,
    }


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

    squat_kg = _to_float(df_f, "Squat (kg)")
    squat_reps = _to_float(df_f, "Squat (reps)")
    bench_kg = _to_float(df_f, "Bench (kg)")
    bench_reps = _to_float(df_f, "Bench (reps)")
    dead_kg = _to_float(df_f, "Deadlift (kg)")
    dead_reps = _to_float(df_f, "Deadlift (reps)")
    fs_kg = _to_float(df_f, "Front Squat (kg)")
    fs_reps = _to_float(df_f, "Front Squat (reps)")
    ohp_kg = _to_float(df_f, "OHP (kg)")
    ohp_reps = _to_float(df_f, "OHP (reps)")
    row_kg = _to_float(df_f, "Rowing (kg)")
    row_reps = _to_float(df_f, "Rowing (reps)")
    pull_kg = _to_float(df_f, "Traction Lestée (kg)")
    pull_reps = _to_float(df_f, "Traction Lestée (reps)")

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

    hspu = _to_float(df_c, "HSPU (reps)")
    mu = _to_float(df_c, "MU (reps)")
    planche = _to_float(df_c, "Planche (sec)")
    t_lest = _to_float(df_c, "Traction Lestée (kg)")
    box = _to_float(df_c, "Box Jump (cm)")

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
    st.header("🏆 PR & Score Athlète Hybride V2")

    wb, data_path = get_excel_file(data_only=True)

    try:
        df_pr = pd.read_excel(data_path, sheet_name="PR Automatiques")
        st.subheader("PR Automatiques (Excel – affichage uniquement)")
        st.dataframe(df_pr)
    except Exception as e:
        st.warning(f"Erreur lecture PR Automatiques : {e}")

    sah_v2, details = compute_sah_v2(data_path)
    st.subheader("Score Athlète Hybride – SAH V2 (calcul Python)")

    if sah_v2 is not None:
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("SAH V2", value=round(sah_v2, 1))
        with col2:
            st.metric("StrengthIndex", value=details.get("StrengthIndex", "N/A"))
        with col3:
            st.metric("SkillIndex", value=details.get("SkillIndex", "N/A"))

        with st.expander("Détails complets SAH V2"):
            st.json(details)
    else:
        st.info("Pas encore assez de données (Force/Cali) pour calculer un SAH V2.")


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
    st.header("🧠 Synthèse & Recommandations globales (100% Python)")

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
    sah_v2, sah_details = compute_sah_v2(data_path)

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(
            "Readiness moyen",
            value=round(readiness_moy, 1) if readiness_moy is not None else "N/A"
        )
    with col2:
        if mean_load is not None:
            st.metric("Charge moyenne (7 dernières séances)", value=int(mean_load))
        else:
            st.metric("Charge moyenne", value="N/A")
    with col3:
        if strain is not None:
            st.metric("Strain (7 dernières séances)", value=int(strain))
        else:
            st.metric("Strain", value="N/A")

    col4, col5 = st.columns(2)
    with col4:
        if monotony is not None:
            st.metric("Monotony", value=round(monotony, 2))
        else:
            st.metric("Monotony", value="N/A")
    with col5:
        if sah_v2 is not None:
            st.metric("SAH V2", value=round(sah_v2, 1))
        else:
            st.metric("SAH V2", value="N/A")

    st.markdown("---")
    st.subheader("Recommandation générale")

    if readiness_moy is None or mean_load is None or strain is None:
        st.info("Pas encore assez de données pour générer une recommandation complète.")
        return

    # Reco simple ici (Auto-Séance détaillée sur la page dédiée)
    if readiness_moy >= 70 and strain < 20000:
        st.write("✅ Tu es dans une bonne zone pour pousser sur des séances lourdes ou de volume.")
    elif readiness_moy < 40 or strain >= 25000:
        st.write("⚠️ Zone de fatigue élevée : privilégie la gestion de la récupération, le skill propre ou le deload.")
    else:
        st.write("🟡 Zone intermédiaire : continue à progresser mais surveille ton sommeil, stress et volumes.")


def page_auto_seance():
    st.header("🤖 Auto-Séance intelligente – Coach Empereur")

    wb, data_path = get_excel_file(data_only=True)

    st.markdown("Cette page te propose un **type de séance du jour** basé sur :")
    st.markdown("- Ta dernière valeur de **Readiness**")
    st.markdown("- La **charge** et le **strain** des 7 dernières séances")
    st.markdown("- Ton **niveau Skill** (calisthénie / puissance)")
    st.markdown("- L’**objectif du bloc** que tu choisis")
    st.markdown("- Le type de ta **dernière séance**")

    block_focus = st.selectbox(
        "Objectif du bloc en cours",
        [
            "Force maximale",
            "Hypertrophie / Volume",
            "Skill / Calisthénie",
            "Puissance / Explosivité",
            "Déload / Gestion fatigue",
        ]
    )

    if st.button("⚡ Générer la séance recommandée"):
        reco = compute_auto_seance_recommendation(data_path, block_focus)

        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Readiness (dernier jour)", value=round(reco["readiness"], 1))
        with col2:
            st.metric("Strain (7 dernières séances)", value=int(reco["strain"]))
        with col3:
            st.metric("SAH V2", value=round(reco["sah_v2"], 1) if reco["sah_v2"] is not None else "N/A")

        st.markdown("---")
        st.subheader("🧬 Profil actuel")
        col4, col5, col6 = st.columns(3)
        with col4:
            st.metric("StrengthIndex", value=round(reco["strength_index"], 1))
        with col5:
            st.metric("SkillIndex", value=round(reco["skill_index"], 1))
        with col6:
            st.metric("PowerIndex", value=round(reco["power_index"], 1))

        st.write(f"**Niveau Skill :** {reco['skill_level']}")

        if reco["last_session"] is not None:
            st.markdown("**Dernière séance enregistrée :**")
            st.json(reco["last_session"])

        st.markdown("---")
        st.subheader("📋 Séance du jour recommandée")

        st.write(f"**Type de séance :** {reco['session_type']}")
        st.write(f"**Focus :** {reco['focus']}")
        st.write(f"**Intensité :** {reco['intensity']}")
        st.write(f"**Volume relatif :** {reco['volume_mod']}")
        st.write(f"**RPE cible :** {reco['rpe_target']}")

        if reco["notes"]:
            st.markdown("**Notes du coach :**")
            for n in reco["notes"]:
                st.write(f"- {n}")

        if reco["structure_suggestion"]:
            st.markdown("**Structure suggérée :**")
            for s in reco["structure_suggestion"]:
                st.write(f"- {s}")


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
    "PR & SAH V2": page_pr_sah,
    "Planning (Annuel / Mésocycles)": page_planning,
    "Synthèse & Recos Globales": page_reco_global,
    "Auto-Séance intelligente": page_auto_seance,
    "Export / Debug": page_export_debug,
}


def main():
    st.set_page_config(page_title="Système Empereur – V2", layout="wide")
    st.sidebar.title("Système d'entraînement de l'Empereur – V2")
    choix = st.sidebar.radio("Navigation", list(PAGES.keys()))
    st.sidebar.markdown("---")
    st.sidebar.write(f"Modèle : `{TEMPLATE_FILE}`")
    st.sidebar.write(f"Données actives : `{DATA_FILE}`")

    PAGES[choix]()


if __name__ == "__main__":
    main()
