import sys
sys.path.append(r"C:\Program Files\DIgSILENT\PowerFactory 2021 SP2\Python\3.9")
import powerfactory as pf
import math
import csv
import os
import random
from datetime import datetime

KV_PER_MV = 1000

app = pf.GetApplication()
if not app:
    raise Exception("Impossible de lancer PowerFactory (pf.GetApplication() = None).")

# ============================================================
# CONFIG : NOMS EXACTS
# ============================================================
PV1_NAME = "Sous-champ 1"
PV2_NAME = "Sous-champ 2"
PCC_NAME = "S/S CAMPUS_6.6 kV"
GRID_NAME = "Kinshasa_30kV"
LDF_NAME = "ComLdf"
QDS_NAME = "ComStatsim"
LINE_NAME = "Makala_Campus_3.6 km"
LOAD_NAME = "Charge Cam1"

# 24h complètes
H_START = 0
H_END = 23

# Nombre de jours synthétiques par scénario
N_DAYS_PER_SCENARIO = 20

# 11 scénarios -> 11 x 20 x 24 = 5280 lignes
SCENARIOS = {
    "S01_normal":                  {"pv_mean": 1.00, "loadP_mean": 1.00, "loadQ_mean": 1.00, "gridV_mean": 1.000},
    "S02_charge_haute":            {"pv_mean": 1.00, "loadP_mean": 1.10, "loadQ_mean": 1.10, "gridV_mean": 0.995},
    "S03_charge_tres_haute":       {"pv_mean": 0.95, "loadP_mean": 1.20, "loadQ_mean": 1.20, "gridV_mean": 0.990},
    "S04_fort_ensoleillement":     {"pv_mean": 1.15, "loadP_mean": 1.00, "loadQ_mean": 1.00, "gridV_mean": 1.005},
    "S05_faible_ensoleillement":   {"pv_mean": 0.75, "loadP_mean": 1.00, "loadQ_mean": 1.00, "gridV_mean": 0.998},
    "S06_charge_haute_pv_faible":  {"pv_mean": 0.80, "loadP_mean": 1.15, "loadQ_mean": 1.15, "gridV_mean": 0.990},
    "S07_charge_faible_pv_fort":   {"pv_mean": 1.20, "loadP_mean": 0.90, "loadQ_mean": 0.90, "gridV_mean": 1.010},
    "S08_reseau_faible":           {"pv_mean": 1.00, "loadP_mean": 1.05, "loadQ_mean": 1.05, "gridV_mean": 0.980},
    "S09_reseau_fort":             {"pv_mean": 1.00, "loadP_mean": 1.00, "loadQ_mean": 1.00, "gridV_mean": 1.020},
    "S10_pointe_soir":             {"pv_mean": 0.90, "loadP_mean": 1.18, "loadQ_mean": 1.18, "gridV_mean": 0.992},
    "S11_mi_saison_mixte":         {"pv_mean": 1.05, "loadP_mean": 0.98, "loadQ_mean": 0.98, "gridV_mean": 1.000},
}

# Variabilité journalière
DAY_RANDOM_SPREAD = {
    "pv": 0.10,      # +/-10%
    "loadP": 0.08,   # +/-8%
    "loadQ": 0.08,   # +/-8%
    "gridV": 0.01    # +/-1%
}

# Onduleur (par sous-champ)
S_INV_MVA_PER_PV = 3.0

# Q-sweep
ALPHA = 0.5
N_SWEEP = 7

# Fonction coût J : poids
wV = 0.45
wL = 0.35
wP = 0.15
wQ = 0.05

# Normalisation
V_BAND = 0.05
LOADING_LIMIT = 80.0
EPS = 1e-9

# Export CSV
OUT_DIR = r"C:\Users\Public\Documents"
ts = datetime.now().strftime("%Y%m%d_%H%M%S")
OUT_CSV = os.path.join(OUT_DIR, f"dataset_multiScenarios_24h_{ts}.csv")

# ============================================================
# OUTILS
# ============================================================
def find_obj_by_loc_name(target_name):
    objs = app.GetCalcRelevantObjects("*.Elm*") + app.GetCalcRelevantObjects("*.Sta*") + app.GetCalcRelevantObjects("*.Typ*")
    for o in objs:
        try:
            if o.loc_name == target_name:
                return o
        except:
            pass

    objs2 = app.GetCalcRelevantObjects("*.*")
    for o in objs2:
        try:
            if o.loc_name == target_name:
                return o
        except:
            pass

    raise Exception(f"Objet introuvable (loc_name exact requis) : {target_name}")

def get_attr_any(obj, keys):
    for k in keys:
        try:
            return obj.GetAttribute(k)
        except:
            pass
    raise Exception(f"Attribut non trouvé sur '{obj.loc_name}'. Essayés: {keys}")

def set_attr_any(obj, keys, value):
    for k in keys:
        try:
            obj.SetAttribute(k, value)
            return k
        except:
            pass
    raise Exception(f"Impossible de définir l'attribut sur '{obj.loc_name}'. Essayés: {keys}")

def try_get_first_existing(obj, keys, default=None):
    for k in keys:
        try:
            return obj.GetAttribute(k), k
        except:
            pass
    return default, None

def linspace_symmetric(q_range, n):
    if n < 2:
        return [0.0]
    if n % 2 == 0:
        n += 1
    step = (2 * q_range) / (n - 1)
    return [-q_range + i * step for i in range(n)]

def qmax_from_SP(S_mva, P_mw):
    val = S_mva * S_mva - P_mw * P_mw
    return math.sqrt(val) if val > 0 else 0.0

def clamp(x, xmin, xmax):
    return max(xmin, min(x, xmax))

def rand_mult(mean, spread):
    return mean * (1.0 + random.uniform(-spread, spread))

def set_if_key(obj, key, value):
    if key is not None:
        obj.SetAttribute(key, value)

# ============================================================
# OBJETS POWERFACTORY
# ============================================================
pv1 = find_obj_by_loc_name(PV1_NAME)
pv2 = find_obj_by_loc_name(PV2_NAME)
pcc = find_obj_by_loc_name(PCC_NAME)
grid = find_obj_by_loc_name(GRID_NAME)
line = find_obj_by_loc_name(LINE_NAME)
load = find_obj_by_loc_name(LOAD_NAME)

ldf = app.GetFromStudyCase(LDF_NAME)
if not ldf:
    raise Exception(f"Commande Load Flow introuvable dans ce Study Case: {LDF_NAME}")

qds = app.GetFromStudyCase(QDS_NAME)
if not qds:
    raise Exception(f"Commande Quasi-Dynamic introuvable dans ce Study Case: {QDS_NAME}")

study = app.GetActiveStudyCase()
if not study:
    raise Exception("Aucun Study Case actif.")

# ============================================================
# LISTE D’ATTRIBUTS
# ============================================================
P_KEYS = ["m:P:bus1", "m:Psum", "m:Pgen", "m:P1"]
Q_KEYS = ["m:Q:bus1", "m:Qsum", "m:Qgen", "m:Q1"]
V_KEYS = ["m:u", "m:U", "m:U1", "m:uk"]
LOADING_KEYS = ["c:loading", "m:loading", "m:Loading", "c:Loading"]
LOSSP_KEYS = ["LossP", "m:LossP", "c:LossP", "m:Ploss", "c:Ploss"]
LOADP_KEYS = ["m:P:bus1", "m:Psum", "m:Plini"]
LOADQ_KEYS = ["m:Q:bus1", "m:Qsum", "m:Qlini"]
QSET_KEYS = ["qsetp", "qset", "qsetp0", "qgini", "qini"]

# ============================================================
# ATTRIBUTS DE VARIATION DES SCÉNARIOS
# Adapte seulement si un attribut n'existe pas chez toi
# ============================================================
LOAD_SCALE_P_KEYS = ["scale0", "slini", "plini_scale", "scale"]
LOAD_SCALE_Q_KEYS = ["scale0", "slini", "qlini_scale", "scale"]

PV_SCALE_KEYS = ["scale0", "pgini_scale", "scale", "pscale"]

GRID_VSET_KEYS = ["usetp", "usetp0", "uini", "Unom", "uknom"]

# ============================================================
# SAUVEGARDE DES VALEURS INITIALES
# ============================================================
load_scale_p_init, load_scale_p_key = try_get_first_existing(load, LOAD_SCALE_P_KEYS, 1.0)
load_scale_q_init, load_scale_q_key = try_get_first_existing(load, LOAD_SCALE_Q_KEYS, 1.0)
pv1_scale_init, pv1_scale_key = try_get_first_existing(pv1, PV_SCALE_KEYS, 1.0)
pv2_scale_init, pv2_scale_key = try_get_first_existing(pv2, PV_SCALE_KEYS, 1.0)
grid_v_init, grid_v_key = try_get_first_existing(grid, GRID_VSET_KEYS, 1.0)

def apply_day_scenario(m_pv, m_loadP, m_loadQ, m_gridV):
    set_if_key(load, load_scale_p_key, load_scale_p_init * m_loadP)
    set_if_key(load, load_scale_q_key, load_scale_q_init * m_loadQ)

    set_if_key(pv1, pv1_scale_key, pv1_scale_init * m_pv)
    set_if_key(pv2, pv2_scale_key, pv2_scale_init * m_pv)

    if grid_v_key is not None and grid_v_init is not None:
        set_if_key(grid, grid_v_key, grid_v_init * m_gridV)

def restore_initial_state():
    set_if_key(load, load_scale_p_key, load_scale_p_init)
    set_if_key(load, load_scale_q_key, load_scale_q_init)
    set_if_key(pv1, pv1_scale_key, pv1_scale_init)
    set_if_key(pv2, pv2_scale_key, pv2_scale_init)

    if grid_v_key is not None and grid_v_init is not None:
        set_if_key(grid, grid_v_key, grid_v_init)

    set_attr_any(pv1, QSET_KEYS, 0.0)
    set_attr_any(pv2, QSET_KEYS, 0.0)

# ============================================================
# EXECUTION
# ============================================================
os.makedirs(OUT_DIR, exist_ok=True)
app.PrintPlain("=== GENERATION DATASET MULTI-JOURS / MULTI-SCENARIOS ===")
app.PrintPlain(f"Fichier CSV : {OUT_CSV}")

n_rows = 0

with open(OUT_CSV, "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f, delimiter=";")
    w.writerow([
        "scenario",
        "day",
        "hour",
        "pv_mult",
        "loadP_mult",
        "loadQ_mult",
        "gridV_mult",
        "P_pv_tot_MW",
        "P_load_MW",
        "Q_load_MVAr",
        "V_pcc_pu",
        "line_loading_pct",
        "P_loss_MW",
        "Qmax_tot_MVAr",
        "Q_try_opt_tot_MVAr",
        "Q_opt_actual_tot_MVAr",
        "J_best"
    ])

    for scen_name, scen_cfg in SCENARIOS.items():
        app.PrintPlain(f"--- Scenario : {scen_name} ---")

        for day in range(1, N_DAYS_PER_SCENARIO + 1):
            pv_mult = clamp(rand_mult(scen_cfg["pv_mean"], DAY_RANDOM_SPREAD["pv"]), 0.50, 1.35)
            loadP_mult = clamp(rand_mult(scen_cfg["loadP_mean"], DAY_RANDOM_SPREAD["loadP"]), 0.70, 1.40)
            loadQ_mult = clamp(rand_mult(scen_cfg["loadQ_mean"], DAY_RANDOM_SPREAD["loadQ"]), 0.70, 1.40)
            gridV_mult = clamp(rand_mult(scen_cfg["gridV_mean"], DAY_RANDOM_SPREAD["gridV"]), 0.97, 1.03)

            apply_day_scenario(pv_mult, loadP_mult, loadQ_mult, gridV_mult)

            ierr = qds.Execute()
            if ierr:
                app.PrintPlain(f"[WARN] ComStatsim non convergé : {scen_name}, jour {day}")
                continue

            for h in range(H_START, H_END + 1):
                # Important : 0h -> 23h
                study.SetStudyTime(h * 3600)

                # Baseline
                set_attr_any(pv1, QSET_KEYS, 0.0)
                set_attr_any(pv2, QSET_KEYS, 0.0)

                ierr = ldf.Execute()
                if ierr:
                    app.PrintPlain(f"[WARN] Baseline non convergé : {scen_name}, jour {day}, heure {h}")
                    continue

                p1 = float(get_attr_any(pv1, P_KEYS))
                p2 = float(get_attr_any(pv2, P_KEYS))
                P_pv_tot = p1 + p2

                P_load = float(get_attr_any(load, LOADP_KEYS))
                Q_load = float(get_attr_any(load, LOADQ_KEYS))
                Ploss_ref = float(get_attr_any(grid, LOSSP_KEYS))

                qmax_tot = qmax_from_SP(S_INV_MVA_PER_PV, p1) + qmax_from_SP(S_INV_MVA_PER_PV, p2)
                q_range = ALPHA * qmax_tot

                if q_range < 1e-6:
                    Q_try_list = [0.0]
                else:
                    Q_try_list = linspace_symmetric(q_range, N_SWEEP)

                best = None

                for qtry_tot in Q_try_list:
                    qtry_each = qtry_tot / 2.0

                    # ici je garde ton principe avec conversion kVAr / MVAr
                    set_attr_any(pv1, QSET_KEYS, qtry_each * KV_PER_MV)
                    set_attr_any(pv2, QSET_KEYS, qtry_each * KV_PER_MV)

                    ierr = ldf.Execute()
                    if ierr:
                        continue

                    q1 = float(get_attr_any(pv1, Q_KEYS))
                    q2 = float(get_attr_any(pv2, Q_KEYS))
                    qact_tot = q1 + q2

                    V = float(get_attr_any(pcc, V_KEYS))
                    L = float(get_attr_any(line, LOADING_KEYS))
                    Pl = float(get_attr_any(grid, LOSSP_KEYS))

                    termV = ((V - 1.0) / V_BAND) ** 2
                    termL = max(0.0, (L - LOADING_LIMIT) / LOADING_LIMIT) ** 2
                    termP = Pl / (Ploss_ref + EPS)
                    termQ = (qact_tot / (qmax_tot + EPS)) ** 2 if qmax_tot > 0 else 0.0

                    J = wV * termV + wL * termL + wP * termP + wQ * termQ

                    if (best is None) or (J < best[0]):
                        best = (J, qtry_tot, qact_tot, V, L, Pl)

                if best is None:
                    app.PrintPlain(f"[WARN] Aucun cas valide : {scen_name}, jour {day}, heure {h}")
                    continue

                Jbest, qtry_opt, qopt_act, Vopt, Lopt, Plopt = best

                w.writerow([
                    scen_name,
                    day,
                    h,
                    round(pv_mult, 6),
                    round(loadP_mult, 6),
                    round(loadQ_mult, 6),
                    round(gridV_mult, 6),
                    round(P_pv_tot, 6),
                    round(P_load, 6),
                    round(Q_load, 6),
                    round(Vopt, 6),
                    round(Lopt, 3),
                    round(Plopt, 6),
                    round(qmax_tot, 6),
                    round(qtry_opt, 6),
                    round(qopt_act, 6),
                    round(Jbest, 8)
                ])
                n_rows += 1

restore_initial_state()

app.PrintPlain("=== FIN ===")
app.PrintPlain(f"Nombre total de lignes générées : {n_rows}")
app.PrintPlain(f"CSV généré : {OUT_CSV}")
app.PrintPlain("Ouvre-le dans Excel (séparateur ;).")