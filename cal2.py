import math
import pandas as pd


def calculate_k(Pc, P, w, Tc, T):
    return (Pc / P) * math.exp(5.37 * (1 + w) * (1 - Tc / T))


def load_database(file_path):
    # Database has 4 stacked header rows + 1 separator + 2 blank rows before data
    # so we skip 7 rows and assign column names manually
    df = pd.read_excel(file_path, header=None, skiprows=7)
    df.columns = [
        'NUMBER', 'COMPONENT', 'MOLE_WT', 'TFP', 'TB', 'TC', 'PC', 'VC', 'ZC', 'OMEGA',
        'LIQDEN', 'TDEN', 'DIPM', 'CP_A', 'CP_B', 'CP_C', 'CP_D', 'LV_B', 'LV_C',
        'DELHG', 'DELGF', 'ANT_A', 'ANT_B', 'ANT_C', 'TMX', 'TMN', 'HAR_A', 'HAR_B',
        'HAR_C', 'HAR_D', 'HV'
    ]
    df = df.dropna(subset=['COMPONENT'])
    return df


def get_user_input(df):
    print("\n--- Dew & Bubble Point Calculator ---\n")
    P = float(input("Enter pressure (Kg/cm2, gauge): "))
    n = int(input("Enter number of components: "))

    components = []
    fractions = []

    print("\nAvailable components (sample - check database.xlsx for full list):")
    print(df['COMPONENT'].head(20).tolist())

    for i in range(n):
        name = input(f"\nComponent {i+1} name (as shown in database): ").strip().upper()
        row = df[df['COMPONENT'].str.upper() == name]

        if row.empty:
            # Try partial match to help the user
            partial = df[df['COMPONENT'].str.upper().str.contains(name, na=False)]
            if not partial.empty:
                print(f"Exact match not found. Did you mean one of these?")
                print(partial['COMPONENT'].tolist())
            else:
                print("Component not found in database.")
            exit()

        frac = float(input("  Mole fraction (or percentage): "))
        components.append(row.iloc[0])
        fractions.append(frac)
    
    # After collecting all fractions, normalize to 0-1 range
    total = sum(fractions)
    if total > 1.5:  # clearly percentages, not fractions
        fractions = [f / 100.0 for f in fractions]
    elif abs(total - 1.0) > 0.01:  # fractions but don't sum to 1, normalize anyway
        fractions = [f / total for f in fractions]

    return P, components, fractions


def extract_properties(components, fractions):
    Pc, Tc, w = [], [], []
    for comp in components:
        Pc.append(comp['PC'])     # ATM
        Tc.append(comp['TC'])     # Kelvin
        w.append(comp['OMEGA'])   # acentric factor
    return fractions, Pc, Tc, w


def calculate(P, y, Pc, Tc, w):
    # Match VBA exactly: T_n starts at 0.25K, step 0.25, 2500 iterations
    T_n = 0.25
    step = 0.25
    iterations = 2500
    TOL = 1e-300  # matches VBA Tol = 1E-300

    best_dew_diff = float('inf')
    best_bubble_diff = float('inf')
    dew_point = None
    bubble_point = None

    for _ in range(iterations):
        k = [calculate_k(Pc[i], P, w[i], Tc[i], T_n) for i in range(len(y))]

        # Dew: skip if k < TOL (matches VBA ky logic)
        dew_sum = sum(y[i] / k[i] for i in range(len(y)) if k[i] >= TOL)

        # Bubble: skip only exact zero (matches VBA kx logic)
        bubble_sum = sum(y[i] * k[i] for i in range(len(y)) if k[i] != 0)

        if abs(dew_sum - 1) < best_dew_diff:
            best_dew_diff = abs(dew_sum - 1)
            dew_point = T_n

        if abs(bubble_sum - 1) < best_bubble_diff:
            best_bubble_diff = abs(bubble_sum - 1)
            bubble_point = T_n

        T_n += step

    return dew_point, bubble_point


def main():
    df = load_database("database.xlsx")
    P, components, fractions = get_user_input(df)
    y, Pc, Tc, w = extract_properties(components, fractions)
    dew, bubble = calculate(P, y, Pc, Tc, w)

    print("\n--- Results ---")
    print(f"Dew Point:    {dew - 273.15:.2f} °C")
    print(f"Bubble Point: {bubble - 273.15:.2f} °C")


if __name__ == "__main__":
    main()