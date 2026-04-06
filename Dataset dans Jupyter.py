#Code qui charge le Dataset dans Jupyter et affiche ses colonnes

df = pd.read_csv(r"C:\Users\Public\Documents\DATASET_TFE_MBUSU_new_essai.csv", sep=";")

print("Nombre de lignes :", len(df))
print("Colonnes :", df.columns)
df.head()







# Code qui renvoi : le chemin d'accès du DATASET, ses nombres des lignes et colonnes :

from pathlib import Path
import pandas as pd

folder = Path(r"C:\Users\Public\Documents")
matches = list(folder.glob("DATASET_TFE_MBUSU_new_essai*.csv"))

if not matches:
    raise FileNotFoundError("Aucun fichier correspondant trouvé dans C:\\Users\\Public\\Documents")

FILE_PATH = matches[0]
print("Fichier utilisé :", FILE_PATH)

df = pd.read_csv(FILE_PATH, sep=";")
print("Nombre de lignes :", len(df))
print("Colonnes :", list(df.columns))








# Code qui utilise le modèle Extrem Learning Machine (ELM) après l'avoir entrainer afin de prédire Q_opt en prenant 6 entrées:
# le modèle est une fonction f(P_pv, P_load, Q_load, V_pcc, Line_loading, P_loss) = Q_opt
# le code débute :

import numpy as np

# Nouvelle observation du réseau
X_new = np.array([[1.3184
, 4.3893
, 2.0917
, 0.99040
, 67.1, 0.2675]])

# Vérification dimensions
print("Entrée brute :", X_new)

# Normalisation
X_new_scaled = sc_X.transform(X_new)
print("Entrée normalisée :", X_new_scaled)

# Prédiction ELM
Q_opt_scaled = elm.predict(X_new_scaled)

# Retour à l'échelle réelle
Q_opt = sc_y.inverse_transform(Q_opt_scaled.reshape(-1, 1)).ravel()

print("Q_opt prédit =", Q_opt[0], "MVAr")








# Code qui donne une prédiction en 24h de Q_opt en fonction des six entrées puis divise chaque Q_opt par deux pour les deux onduleurs :

import pandas as pd
import numpy as np
from pathlib import Path

# ============================================================
# 1) CHEMINS DES 6 FICHIERS QUASI-DYNAMIQUES
#    Remplace si nécessaire par tes vrais chemins locaux
# ============================================================

file_pv = Path(r"C:\Users\Public\Documents\P_pv.csv")               # Sous-champ 2 - Active Power in MW
file_pload = Path(r"C:\Users\Public\Documents\P_load.csv")         # Charge Cam1 - Active Power in MW
file_qload = Path(r"C:\Users\Public\Documents\Q_load.csv")         # Charge Cam1 - Reactive Power in Mvar
file_vpcc = Path(r"C:\Users\Public\Documents\V_pcc.csv")            # S/S CAMPUS_6.6 kV - u, Magnitude in p.u.
file_loading = Path(r"C:\Users\Public\Documents\Line_loading.csv") # Makala_Campus_3.6 km - Loading in %
file_ploss = Path(r"C:\Users\Public\Documents\P_loss.csv")         # Kinshasa_30kV - Losses in MW

# ============================================================
# 2) FONCTION DE LECTURE DES EXPORTS POWERFACTORY
# ============================================================

def read_pf_export(path):
    """
    Lit un export CSV PowerFactory du type :
    ligne 1 : titre
    ligne 2 : en-têtes
    puis données horaires
    Séparateur = ;
    Décimales = ,
    """
    df = pd.read_csv(
        path,
        sep=";",
        decimal=",",
        skiprows=1,   # on saute la première ligne de titre
        engine="python"
    )

    # Nettoyage des noms de colonnes
    df.columns = [str(c).strip().replace('"', '') for c in df.columns]

    # On suppose que la première colonne est le temps et la seconde la valeur
    time_col = df.columns[0]
    value_col = df.columns[1]

    # Conversion temps
    df[time_col] = pd.to_datetime(df[time_col], format="%Y.%m.%d %H.%M.%S", errors="coerce")

    # Conversion numérique
    df[value_col] = pd.to_numeric(df[value_col], errors="coerce")

    # Extraire l'heure
    df["hour"] = df[time_col].dt.hour

    # Garder seulement heure + valeur
    df = df[["hour", value_col]].copy()

    return df, value_col

# ============================================================
# 3) LECTURE DES 6 FICHIERS
# ============================================================

df_pv, col_pv = read_pf_export(file_pv)
df_pload, col_pload = read_pf_export(file_pload)
df_qload, col_qload = read_pf_export(file_qload)
df_vpcc, col_vpcc = read_pf_export(file_vpcc)
df_loading, col_loading = read_pf_export(file_loading)
df_ploss, col_ploss = read_pf_export(file_ploss)

# ============================================================
# 4) RENOMMAGE DES COLONNES SELON LE MODELE ELM
# ============================================================

df_pv = df_pv.rename(columns={col_pv: "P_pv_one_MW"})
df_pload = df_pload.rename(columns={col_pload: "P_load_MW"})
df_qload = df_qload.rename(columns={col_qload: "Q_load_MVAr"})
df_vpcc = df_vpcc.rename(columns={col_vpcc: "V_pcc_pu"})
df_loading = df_loading.rename(columns={col_loading: "line_loading_pct"})
df_ploss = df_ploss.rename(columns={col_ploss: "P_loss_MW"})

# ============================================================
# 5) FUSION DES 6 SERIES SUR L'HEURE
# ============================================================

df_24h = df_pv.merge(df_pload, on="hour") \
              .merge(df_qload, on="hour") \
              .merge(df_vpcc, on="hour") \
              .merge(df_loading, on="hour") \
              .merge(df_ploss, on="hour")

# Si le fichier PV représente un seul sous-champ et que les 2 sont identiques :
df_24h["P_pv_tot_MW"] = 2.0 * df_24h["P_pv_one_MW"]

# ============================================================
# 6) CONSTRUIRE LES ENTREES DU MODELE ELM
# ============================================================

INPUT_COLS = [
    "P_pv_tot_MW",
    "P_load_MW",
    "Q_load_MVAr",
    "V_pcc_pu",
    "line_loading_pct",
    "P_loss_MW"
]

X_new = df_24h[INPUT_COLS].to_numpy(dtype=np.float64)

# ============================================================
# 7) NORMALISATION + PREDICTION ELM
# ============================================================

X_new_scaled = sc_X.transform(X_new)
Q_opt_scaled = elm.predict(X_new_scaled)
Q_opt = sc_y.inverse_transform(Q_opt_scaled.reshape(-1, 1)).ravel()

df_24h["Q_opt_total_MVAr"] = Q_opt
df_24h["Q1_MVAr"] = df_24h["Q_opt_total_MVAr"] / 2.0
df_24h["Q2_MVAr"] = df_24h["Q_opt_total_MVAr"] / 2.0

# Pour PowerFactory en kvar
df_24h["Q1_kvar"] = df_24h["Q1_MVAr"] * 1000.0
df_24h["Q2_kvar"] = df_24h["Q2_MVAr"] * 1000.0

# ============================================================
# 8) RESULTATS
# ============================================================

print("===== Qopt prédits sur 24 h =====")
print(df_24h[[
    "hour",
    "P_pv_tot_MW",
    "P_load_MW",
    "Q_load_MVAr",
    "V_pcc_pu",
    "line_loading_pct",
    "P_loss_MW",
    "Q_opt_total_MVAr",
    "Q1_kvar",
    "Q2_kvar"
]])

# Sauvegarde
out_file = Path(r"C:\Users\Public\Documents\Qopt_24h_ELM.csv")
df_24h.to_csv(out_file, sep=";", index=False, encoding="utf-8-sig")
print("\nFichier sauvegardé :", out_file)