# -*- coding: utf-8 -*-
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score

# =========================
# CONFIG
# =========================
FILE_PATH = Path(r"C:\Users\VotreNom\Desktop\dataset_multiScenarios_24h.csv")
FILE_TYPE = "csv"   # "csv" ou "excel"
CSV_SEP = ";"

INPUT_COLS = [
    "P_pv_tot_MW",
    "P_load_MW",
    "Q_load_MVAr",
    "V_pcc_pu",
    "line_loading_pct",
    "P_loss_MW",
]

TARGET_COL = "Q_opt_actual_tot_MVAr"

TEST_SIZE = 0.2
RANDOM_STATE = 42

# =========================
# LECTURE DU FICHIER
# =========================
if FILE_TYPE == "csv":
    df = pd.read_csv(FILE_PATH, sep=CSV_SEP)
else:
    df = pd.read_excel(FILE_PATH)

df.columns = [str(c).strip() for c in df.columns]

missing = [c for c in INPUT_COLS + [TARGET_COL] if c not in df.columns]
if missing:
    raise ValueError(f"Colonnes manquantes : {missing}")

# Conversion en numérique
for c in INPUT_COLS + [TARGET_COL]:
    df[c] = pd.to_numeric(df[c], errors="coerce")

# Suppression des lignes incomplètes
df = df.dropna(subset=INPUT_COLS + [TARGET_COL]).reset_index(drop=True)

print("Nombre de lignes valides :", len(df))
print("Colonnes utilisées :", INPUT_COLS)
print("Cible :", TARGET_COL)

X = df[INPUT_COLS].to_numpy(dtype=np.float64)
y = df[TARGET_COL].to_numpy(dtype=np.float64)

# =========================
# SPLIT TRAIN / TEST
# =========================
Xtr, Xte, ytr, yte = train_test_split(
    X, y, test_size=TEST_SIZE, random_state=RANDOM_STATE, shuffle=True
)

print("Train :", Xtr.shape, ytr.shape)
print("Test  :", Xte.shape, yte.shape)

# =========================
# NORMALISATION
# =========================
sc_X = StandardScaler()
XtrN = sc_X.fit_transform(Xtr)
XteN = sc_X.transform(Xte)

sc_y = StandardScaler()
ytrN = sc_y.fit_transform(ytr.reshape(-1, 1)).ravel()

# =========================
# MODELE ELM
# =========================
class ExtremeLearningMachine:
    def __init__(self, num_hidden_neurons=50, activation_function="tanh", seed=0):
        self.num_hidden_neurons = int(num_hidden_neurons)
        self.activation_function = activation_function
        self.seed = int(seed)
        self.input_weights = None
        self.biases = None
        self.output_weights = None

    def _activation(self, x):
        if self.activation_function == "sigmoid":
            x = np.clip(x, -500, 500)
            return 1.0 / (1.0 + np.exp(-x))
        elif self.activation_function == "tanh":
            return np.tanh(x)
        elif self.activation_function == "relu":
            return np.maximum(0.0, x)
        else:
            raise ValueError("Activation non supportée")

    def fit(self, X, y):
        rng = np.random.default_rng(self.seed)
        n_features = X.shape[1]

        self.input_weights = rng.standard_normal(
            (self.num_hidden_neurons, n_features)
        )
        self.biases = rng.standard_normal(self.num_hidden_neurons)

        H = self._activation(X @ self.input_weights.T + self.biases[None, :])
        self.output_weights = np.linalg.pinv(H) @ y

    def predict(self, X):
        H = self._activation(X @ self.input_weights.T + self.biases[None, :])
        return H @ self.output_weights

# =========================
# ENTRAINEMENT
# =========================
elm = ExtremeLearningMachine(
    num_hidden_neurons=80,
    activation_function="tanh",
    seed=42
)
elm.fit(XtrN, ytrN)

# =========================
# PREDICTION
# =========================
y_predN = elm.predict(XteN)
y_pred = sc_y.inverse_transform(y_predN.reshape(-1, 1)).ravel()

# =========================
# METRICS
# =========================
mae = mean_absolute_error(yte, y_pred)
rmse = np.sqrt(mean_squared_error(yte, y_pred))
r2 = r2_score(yte, y_pred)

print("\n===== METRICS TEST =====")
print("MAE  :", mae)
print("RMSE :", rmse)
print("R2   :", r2)

# =========================
# TABLEAU RESULTATS
# =========================
df_results = pd.DataFrame({
    "Q_reel": yte,
    "Q_pred": y_pred,
    "Erreur_abs": np.abs(yte - y_pred)
})

print("\n===== APERCU RESULTATS =====")
print(df_results.head(20))

# Sauvegarde
out_path = FILE_PATH.parent / "resultats_elm.csv"
df_results.to_csv(out_path, sep=";", index=False, encoding="utf-8")
print("\nRésultats sauvegardés dans :", out_path)

# =========================
# GRAPHIQUES
# =========================
plt.figure(figsize=(6, 5))
plt.scatter(yte, y_pred, alpha=0.6)
mn = min(np.min(yte), np.min(y_pred))
mx = max(np.max(yte), np.max(y_pred))
plt.plot([mn, mx], [mn, mx], "r--")
plt.xlabel("Valeur réelle")
plt.ylabel("Valeur prédite")
plt.title("ELM - Régression de Q_opt")
plt.grid(True)
plt.tight_layout()
plt.show()

plt.figure(figsize=(10, 4))
nshow = min(200, len(yte))
plt.plot(yte[:nshow], label="Réel")
plt.plot(y_pred[:nshow], label="Prédit")
plt.xlabel("Échantillons test")
plt.ylabel("Q_opt_actual_tot_MVAr")
plt.title("Comparaison Réel / Prédit")
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.show()


CODE Pour Predire avec ELM après entraînement donnant 3 Qopt pour chaque ligne :

import numpy as np

# Plusieurs états du réseau
X_test = np.array([
    [4.8, 3.1, 1.2, 0.98, 67, 0.21],
    [3.5, 4.0, 1.5, 0.96, 72, 0.25],
    [5.2, 2.8, 0.9, 1.01, 60, 0.18]
])

# Normalisation
X_scaled = sc_X.transform(X_test)

# Prédiction
Q_scaled = elm.predict(X_scaled)

# Retour à l’échelle réelle
Q_opt = sc_y.inverse_transform(Q_scaled.reshape(-1,1)).ravel()

print(Q_opt)



Une seule prediction (nom du modèle est elm) :

import numpy as np

# Nouvelle observation du réseau
X_new = np.array([[4.8, 3.1, 1.2, 0.98, 67, 0.21]])

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