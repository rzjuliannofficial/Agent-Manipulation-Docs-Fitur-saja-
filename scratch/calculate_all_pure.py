import sys

# Raw data A01 to A20 (from PDF)
RAW_DATA = [
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
]

IS_BENEFIT = [True, True, True, True, False, False, False, False, False, False]

W = [
    0.10958435, 0.12602201, 0.14555542, 0.13862421, 0.09962214,
    0.07484759, 0.08233235, 0.06480311, 0.09056558, 0.06804326
]

SAW_V = [16.80, 13.77, 17.32, 12.00, 15.22, 12.60, 13.85, 13.95, 13.87, 13.35, 
         12.75, 14.28, 11.75, 16.95, 11.02, 13.27, 13.05, 12.37, 12.75, 11.32]

WP_V = [66.806, 49.833, 69.343, 42.512, 59.856, 43.040, 50.127, 53.380, 50.958, 48.420, 
        45.563, 54.086, 41.414, 66.550, 39.432, 45.858, 45.897, 43.752, 44.180, 38.994]

TOPSIS_V = [0.728656, 0.49106, 0.765337, 0.460397, 0.65903, 0.458221, 0.518355, 0.594643, 0.558432, 0.497844, 
            0.474355, 0.548282, 0.443418, 0.7578, 0.336822, 0.466242, 0.470457, 0.488074, 0.442941, 0.329795]

def get_ranks(values, reverse=True):
    # rank 1 for highest if reverse=True, else rank 1 for lowest
    indexed = list(enumerate(values))
    # sort by value
    sorted_indexed = sorted(indexed, key=lambda x: x[1], reverse=reverse)
    
    ranks = [0] * len(values)
    for rank, (original_idx, val) in enumerate(sorted_indexed, 1):
        ranks[original_idx] = rank
    return ranks

SAW_R = get_ranks(SAW_V, reverse=True)
WP_R = get_ranks(WP_V, reverse=True)
TOPSIS_R = get_ranks(TOPSIS_V, reverse=True)

# MAIRCA calculations
m = len(RAW_DATA) # 20
max_val = [max(RAW_DATA[i][j] for i in range(m)) for j in range(10)]
min_val = [min(RAW_DATA[i][j] for i in range(m)) for j in range(10)]

NORM = []
for i in range(m):
    row = []
    for j in range(10):
        val = RAW_DATA[i][j]
        mx = max_val[j]
        mn = min_val[j]
        if IS_BENEFIT[j]:
            norm_val = (val - mn) / (mx - mn) if mx != mn else 1.0
        else:
            norm_val = (mx - val) / (mx - mn) if mx != mn else 1.0
        row.append(norm_val)
    NORM.append(row)

TP = [w / m for w in W]

PP = []
for i in range(m):
    row = [TP[j] * NORM[i][j] for j in range(10)]
    PP.append(row)

GAP = []
for i in range(m):
    row = [TP[j] - PP[i][j] for j in range(10)]
    GAP.append(row)

QI = [sum(GAP[i]) for i in range(m)]
MAIRCA_R = get_ranks(QI, reverse=False) # Smaller is better

print("MAIRCA Qi values and ranks:")
for i in range(20):
    print(f"A{i+1:02d}: Qi={QI[i]:.6f}, Rank={MAIRCA_R[i]}")

# Spearman correlation function
def spearman_corr(r1, r2):
    n = len(r1)
    sum_d2 = sum((r1[i] - r2[i])**2 for i in range(n))
    return 1 - (6 * sum_d2) / (n * (n**2 - 1))

corr_saw = spearman_corr(MAIRCA_R, SAW_R)
corr_wp = spearman_corr(MAIRCA_R, WP_R)
corr_topsis = spearman_corr(MAIRCA_R, TOPSIS_R)

print(f"\nSpearman Correlation:\nSAW: {corr_saw:.6f}\nWP: {corr_wp:.6f}\nTOPSIS: {corr_topsis:.6f}")

# --- Sensitivity Analysis ---
# Test K3 weight variation: K3 is the highest weight at index 2 (0.14555542)
# We change its relative weight by a multiplier, normalize, and recalculate ranks
multipliers = [-0.20, -0.10, 0.10, 0.20]
print("\nSensitivity Analysis (Weight K3 variation):")
for mult in multipliers:
    # Scale K3 weight
    new_W = list(W)
    new_W[2] = W[2] * (1.0 + mult)
    # Re-normalize weights so sum is 1
    sum_W = sum(new_W)
    new_W = [w / sum_W for w in new_W]
    
    # Recalculate MAIRCA Qi
    new_TP = [w / m for w in new_W]
    new_QI = []
    for i in range(m):
        q = sum((new_TP[j] - (new_TP[j] * NORM[i][j])) for j in range(10))
        new_QI.append(q)
    
    new_R = get_ranks(new_QI, reverse=False)
    
    # Check changes vs base MAIRCA_R
    changes = sum(1 for i in range(m) if new_R[i] != MAIRCA_R[i])
    pct_change = (changes / m) * 100
    top_alt = new_R.index(1) + 1
    print(f"K3 {mult*100:+.0f}%: Top Alt=A{top_alt:02d}, Rank Changes={changes} ({pct_change:.1f}%)")
