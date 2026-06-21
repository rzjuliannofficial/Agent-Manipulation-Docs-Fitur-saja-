import numpy as np
import scipy.stats as stats

# Raw data A01 to A20 (from PDF)
RAW_DATA = np.array([
    [4, 3, 5, 5, 1, 1, 4, 1, 4, 2], # A01
    [4, 3, 1, 1, 3, 1, 1, 1, 1, 2], # A02
    [4, 4, 5, 5, 3, 1, 1, 1, 4, 1], # A03
    [4, 2, 1, 1, 2, 1, 4, 1, 4, 2], # A04
    [4, 3, 5, 5, 3, 2, 4, 1, 1, 2], # A05
    [4, 4, 1, 1, 5, 1, 4, 1, 4, 2], # A06
    [4, 3, 1, 1, 1, 1, 4, 1, 2, 2], # A07
    [3, 3, 5, 5, 5, 1, 3, 1, 3, 3], # A08
    [4, 3, 1, 5, 3, 1, 4, 1, 4, 2], # A09
    [4, 3, 1, 1, 2, 1, 1, 1, 4, 2], # A10
    [4, 2, 1, 1, 2, 1, 4, 1, 1, 2], # A11
    [4, 2, 5, 5, 3, 3, 4, 1, 1, 2], # A12
    [4, 2, 1, 1, 3, 1, 4, 1, 3, 2], # A13
    [4, 5, 5, 5, 1, 2, 4, 1, 5, 2], # A14
    [4, 3, 1, 1, 3, 3, 4, 1, 3, 2], # A15
    [4, 4, 1, 1, 5, 1, 1, 1, 3, 3], # A16
    [4, 5, 1, 1, 1, 3, 4, 2, 3, 2], # A17
    [4, 4, 1, 1, 3, 1, 4, 2, 4, 2], # A18
    [4, 3, 1, 1, 5, 1, 4, 1, 1, 2], # A19
    [4, 4, 1, 1, 3, 4, 4, 1, 5, 2], # A20
])

# Kriteria types (True: Benefit, False: Cost)
IS_BENEFIT = np.array([True, True, True, True, False, False, False, False, False, False])

# SWARA Weights (recalculated from Excel)
# Wj = [0.1096, 0.1260, 0.1456, 0.1386, 0.0996, 0.0748, 0.0823, 0.0648, 0.0906, 0.0680]
W = np.array([
    0.10958435, 0.12602201, 0.14555542, 0.13862421, 0.09962214,
    0.07484759, 0.08233235, 0.06480311, 0.09056558, 0.06804326
])

# SAW, WP, TOPSIS V values from image
SAW_V = np.array([16.80, 13.77, 17.32, 12.00, 15.22, 12.60, 13.85, 13.95, 13.87, 13.35, 
                  12.75, 14.28, 11.75, 16.95, 11.02, 13.27, 13.05, 12.37, 12.75, 11.32])

WP_V = np.array([66.806, 49.833, 69.343, 42.512, 59.856, 43.040, 50.127, 53.380, 50.958, 48.420, 
                 45.563, 54.086, 41.414, 66.550, 39.432, 45.858, 45.897, 43.752, 44.180, 38.994])

TOPSIS_V = np.array([0.728656, 0.49106, 0.765337, 0.460397, 0.65903, 0.458221, 0.518355, 0.594643, 0.558432, 0.497844, 
                     0.474355, 0.548282, 0.443418, 0.7578, 0.336822, 0.466242, 0.470457, 0.488074, 0.442941, 0.329795])

# Calculate ranks (higher value = better rank for SAW, WP, TOPSIS)
def get_rank_descending(arr):
    # rank 1 for highest value
    return len(arr) - np.argsort(np.argsort(arr))

SAW_R = get_rank_descending(SAW_V)
WP_R = get_rank_descending(WP_V)
TOPSIS_R = get_rank_descending(TOPSIS_V)

# --- MAIRCA ---
m = len(RAW_DATA) # 20
max_val = np.max(RAW_DATA, axis=0)
min_val = np.min(RAW_DATA, axis=0)

# Normalization
NORM = np.zeros_like(RAW_DATA, dtype=float)
for j in range(10):
    if IS_BENEFIT[j]:
        NORM[:, j] = (RAW_DATA[:, j] - min_val[j]) / (max_val[j] - min_val[j])
    else:
        NORM[:, j] = (max_val[j] - RAW_DATA[:, j]) / (max_val[j] - min_val[j])

# Preferensi Teoritis
TP = W / m

# Preferensi Nyata
PP = np.zeros_like(RAW_DATA, dtype=float)
for j in range(10):
    PP[:, j] = TP[j] * NORM[:, j]

# Gap
GAP = np.zeros_like(RAW_DATA, dtype=float)
for j in range(10):
    GAP[:, j] = TP[j] - PP[:, j]

# Qi
QI = np.sum(GAP, axis=1)

# Ranks (smaller Qi = better rank)
def get_rank_ascending(arr):
    return np.argsort(np.argsort(arr)) + 1

MAIRCA_R = get_rank_ascending(QI)

print("MAIRCA Qi values and ranks:")
for i in range(20):
    print(f"A{i+1:02d}: Qi={QI[i]:.6f}, Rank={MAIRCA_R[i]}")

# Spearman correlation
coef_saw, _ = stats.spearmanr(MAIRCA_R, SAW_R)
coef_wp, _ = stats.spearmanr(MAIRCA_R, WP_R)
coef_topsis, _ = stats.spearmanr(MAIRCA_R, TOPSIS_R)

print(f"\nSpearman Correlation:\nSAW: {coef_saw:.4f}\nWP: {coef_wp:.4f}\nTOPSIS: {coef_topsis:.4f}")
