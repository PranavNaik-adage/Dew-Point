import math
import pandas as pd


# ==============================================================================
# HARDCODED INPUT — change these values as needed
# ==============================================================================
PRESSURE_KGCM2_GAUGE = 1.5   # kg/cm² gauge

# MIXTURE = [
#     # ( "COMPONENT NAME",  mole_percent )
#     ("NITROGEN",        4.33),
#     ("CARBON DIOXIDE",  0.005),
#     ("METHANE",         29.76),
#     ("ETHANE",          35.20),
#     ("PROPANE",         12.73),
#     ("ISOBUTANE",       0.10),
#     ("N-BUTANE",        0.17),
#     ("2-METHYLBUTANE",  17.61),
#     ("N-PENTANE",       0.09),
# ]

MIXTURE = [
    ("METHANE",         93.37),
    ("ETHANE",          3.83),
    ("PROPANE",         1.20),
    ("ISOBUTANE",       0.29),
    ("N-BUTANE",        0.30),
    ("2-METHYLBUTANE",  0.16),
    ("N-PENTANE",       0.09),
    ("N-HEXANE",        0.12),
    ("N-HEPTANE",       0.08),
    ("N-OCTANE",        0.04),
    ("NITROGEN",        0.45),
]
# ==============================================================================


def load_database(file_path="database.xlsx"):
    """
    Database layout (0-indexed rows):
      Row 0  : property names line 1   (header)
      Row 1  : property names line 2   (header)
      Row 2  : units line 1            (header)
      Row 3  : units line 2 / symbols  (header)
      Row 4  : *** separator ***
      Row 5  : blank
      Row 6  : blank
      Row 7+ : data rows

    The very first column (col 0) is the component NUMBER,
    the second column (col 1) is the component NAME,
    then the properties follow in the order shown in the header.
    """
    df = pd.read_excel(file_path, header=None, skiprows=7)

    # Assign column names that match the database header symbols exactly
    df.columns = [
        'NUMBER',    # col 0  — component serial number
        'COMPONENT', # col 1  — component name
        'MOLE_WT',   # col 2
        'TFP',       # col 3  — freeze point, K
        'TB',        # col 4  — boiling point, K
        'TC',        # col 5  — critical temp, K
        'PC',        # col 6  — critical pressure, ATM
        'VC',        # col 7  — critical volume, cc/g-mol
        'ZC',        # col 8  — critical compressibility
        'OMEGA',     # col 9  — acentric factor
        'LIQDEN',    # col 10 — liquid density, g/cc
        'TDEN',      # col 11 — ref temp for liq den, K
        'DIPM',      # col 12 — dipole moment, Debyes
        'CP_A',      # col 13
        'CP_B',      # col 14
        'CP_C',      # col 15
        'CP_D',      # col 16
        'LV_B',      # col 17
        'LV_C',      # col 18
        'DELHG',     # col 19 — std heat of formation, kcal/g-mol
        'DELGF',     # col 20 — std energy of formation
        'ANT_A',     # col 21 — Antoine A
        'ANT_B',     # col 22 — Antoine B
        'ANT_C',     # col 23 — Antoine C
        'TMX',       # col 24 — max temp for Antoine, K
        'TMN',       # col 25 — min temp for Antoine, K
        'HAR_A',     # col 26 — Harlacher A
        'HAR_B',     # col 27 — Harlacher B
        'HAR_C',     # col 28 — Harlacher C
        'HAR_D',     # col 29 — Harlacher D
        'HV',        # col 30 — heat of vaporisation, cal/g-mol
    ]

    # Drop rows where COMPONENT is blank/NaN
    df = df.dropna(subset=['COMPONENT'])
    df['COMPONENT'] = df['COMPONENT'].astype(str).str.strip().str.upper()
    return df


def convert_pressure(P_gauge_kgcm2):
    """
    Convert kg/cm² gauge  →  ATM absolute
    1 atm = 1.0332 kg/cm²
    gauge → absolute: add 1.0332 kg/cm²
    absolute kg/cm² → ATM: divide by 1.0332
    """
    P_abs_kgcm2 = P_gauge_kgcm2 + 1.0332
    P_atm = P_abs_kgcm2 / 1.0332
    return P_atm


def wilson_k(Pc_atm, P_atm, omega, Tc_K, T_K):
    """Wilson correlation K-value (same formula as VBA)."""
    return (Pc_atm / P_atm) * math.exp(5.37 * (1.0 + omega) * (1.0 - Tc_K / T_K))


def calculate_dew_bubble(P_atm, component_data, mole_fractions):
    """
    Replicate the VBA loop exactly:
      T_n starts at 0.25 K, step +0.25 K, 2500 iterations  →  max 625 K
      TOL = 1e-300

    For dew point  : sum(y_i / K_i) → 1
    For bubble point: sum(x_i * K_i) → 1
    (x_i == y_i == z_i for a feed — same as VBA)
    """
    TOL = 1e-300
    T_n = 0.25
    step = 0.25
    iterations = 2500

    results = []   # (T, dew_diff, bubble_diff)

    for _ in range(iterations):
        dew_sum = 0.0
        bubble_sum = 0.0

        for z_i, props in zip(mole_fractions, component_data):
            Pc  = props['PC']
            Tc  = props['TC']
            w   = props['OMEGA']

            k = wilson_k(Pc, P_atm, w, Tc, T_n)

            # --- dew sum: y/k  (skip if k < TOL, matching VBA ky logic) ---
            if k >= TOL:
                dew_sum += z_i / k

            # --- bubble sum: x*k  (skip only exact zero, matching VBA kx logic) ---
            if k != 0:
                bubble_sum += z_i * k

        results.append((T_n, abs(dew_sum - 1), abs(bubble_sum - 1)))
        T_n += step

    # Find temperatures where each difference is minimised
    dew_row    = min(results, key=lambda r: r[1])
    bubble_row = min(results, key=lambda r: r[2])

    return dew_row[0], bubble_row[0]


def main():
    # ── 1. Load database ──────────────────────────────────────────────────────
    df = load_database("database.xlsx")

    # ── 2. Convert pressure ───────────────────────────────────────────────────
    P_atm = convert_pressure(PRESSURE_KGCM2_GAUGE)
    print(f"\nPressure : {PRESSURE_KGCM2_GAUGE} kg/cm² gauge  "
          f"→  {P_atm:.4f} ATM absolute")

    # ── 3. Look up each component ─────────────────────────────────────────────
    component_data  = []
    mole_fractions  = []
    raw_percentages = []

    print("\nComponent lookup:")
    for name, pct in MIXTURE:
        key = name.strip().upper()
        row = df[df['COMPONENT'] == key]

        if row.empty:
            # Try partial match to give a helpful hint
            partial = df[df['COMPONENT'].str.contains(key, na=False)]
            if not partial.empty:
                print(f"  ✗ '{name}' not found. Closest matches: "
                      f"{partial['COMPONENT'].tolist()}")
            else:
                print(f"  ✗ '{name}' not found in database.")
            raise SystemExit(1)

        props = row.iloc[0]
        print(f"  ✓ {name:20s}  TC={props['TC']:.1f} K  "
              f"PC={props['PC']:.2f} atm  OMEGA={props['OMEGA']:.4f}")
        component_data.append(props)
        raw_percentages.append(pct)

    # ── 4. Normalise mole fractions ───────────────────────────────────────────
    total = sum(raw_percentages)
    if total > 1.5:                        # entered as percentages
        mole_fractions = [p / 100.0 for p in raw_percentages]
    elif abs(total - 1.0) > 0.01:         # fractions but don't sum to 1
        mole_fractions = [p / total  for p in raw_percentages]
    else:
        mole_fractions = raw_percentages   # already proper fractions

    print(f"\nMole fractions sum = {sum(mole_fractions):.6f}")

    # ── 5. Run calculation ────────────────────────────────────────────────────
    dew_K, bubble_K = calculate_dew_bubble(P_atm, component_data, mole_fractions)

    # ── 6. Print results ──────────────────────────────────────────────────────
    print("\n--- Results ---")
    print(f"Dew    Point : {dew_K:.2f} K  =  {dew_K - 273.15:.2f} °C")
    print(f"Bubble Point : {bubble_K:.2f} K  =  {bubble_K - 273.15:.2f} °C")


if __name__ == "__main__":
    main()