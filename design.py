import numpy as np
from scipy.optimize import differential_evolution

def objective(x):
    n_s, n_p, B, n_r, L_p, L_s = x
    penalty = 0.0

    # (1) width
    width = 30*n_s + (30*n_p + 45) + (150*B + 40)
    if width > 1340:
        penalty += (width - 1340)**2

    # (2) length
    length = 12*n_r + 2*40
    if length > 4500:
        penalty += (length - 4500)**2

    # (3) order
    if L_p >= L_s:
        penalty += (L_p - L_s + 1)**2

    # (4) parking
    if B*(2*n_r) < n_p * L_p / 20:
        penalty += (n_p * L_p / 20 - B*(2*n_r))**2

    # (5) proportional
    if n_s < 0.2*n_p:
        penalty += (0.2*n_p - n_s)**2
    if n_s > 0.6*n_p:
        penalty += (n_s - 0.6*n_p)**2

    obj = -n_p * L_p
    return obj + 1e4 * penalty


bounds = [
    (1, 20),      # n_s
    (1, 20),      # n_p
    (1, 20),      # B
    (1, 400),     # n_r
    (2000, 4000), # L_p
    (2000, 4000)  # L_s
]


result = differential_evolution(
    objective,
    bounds,
    maxiter=800,
    popsize=20,
    tol=1e-7,
    polish=True,
    disp=False
)

n_s, n_p, B, n_r, L_p, L_s = result.x
print("\n===== Optimization Result =====")
print(f"n_s = {n_s:.2f}")
print(f"n_p = {n_p:.2f}")
print(f"B   = {B:.2f}")
print(f"n_r = {n_r:.2f}")
print(f"L_p = {L_p:.2f}")
print(f"L_s = {L_s:.2f}")
print(f"Objective (n_p * L_p) ≈ {-result.fun:.2f}")


n_s_i, n_p_i, B_i, n_r_i, L_p_i, L_s_i = np.round(result.x).astype(int)

width_check = 30*n_s_i + (30*n_p_i + 45) + (150*B_i + 40)
length_check = 12*n_r_i + 2*40
parking_LHS = B_i * (2 * n_r_i)
parking_RHS = n_p_i * L_p_i / 20
proportion_low = 0.2 * n_p_i
proportion_high = 0.6 * n_p_i

print("\n===== Rounded Feasible Check =====")
print(f"n_s={n_s_i}, n_p={n_p_i}, B={B_i}, n_r={n_r_i}, L_p={L_p_i}, L_s={L_s_i}")

# 1. Width
print(f"Width: {width_check:.1f} ≤ 1340.0 → {width_check <= 1340}")

# 2. Length
print(f"Length: {length_check:.1f} ≤ 4500.0 → {length_check <= 4500}")

# 3. Parking
print(f"Parking: {parking_LHS:.1f} ≥ {parking_RHS:.1f} → {parking_LHS >= parking_RHS}")

# 4. Proportion constraint
print(f"0.2n_p ≤ n_s ≤ 0.6n_p → {proportion_low:.1f} ≤ {n_s_i} ≤ {proportion_high:.1f} → {proportion_low <= n_s_i <= proportion_high}")

# 5. Objective
print(f"n_p * L_p = {n_p_i * L_p_i:.1f}")