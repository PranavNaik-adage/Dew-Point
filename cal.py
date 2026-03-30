import math
import pandas as pd


def calculate_k(Pc, P, w, Tc, T):
    return (Pc / P) * math.exp(5.37 * (1 + w) * (1 - Tc / T))


def load_database(file_path):
    df = pd.read_excel(file_path)
    
    # Normalize column names (important)
    df.columns = df.columns.str.strip().str.lower()
    
    return df


def get_user_input(df):
    print("\n--- Dew & Bubble Point Calculator ---\n")

    P = float(input("Enter pressure: "))
    n = int(input("Enter number of components: "))

    components = []
    fractions = []

    print("\nAvailable components:")
    print(df.iloc[:, 0].tolist())  # assuming first column = names

    for i in range(n):
        name = input(f"\nComponent {i+1} name: ").strip().lower()

        # find row
        row = df[df.iloc[:, 0].str.lower() == name]

        if row.empty:
            print("❌ Component not found in database!")
            exit()

        frac = float(input("  Mole fraction: "))

        components.append(row.iloc[0])
        fractions.append(frac)

    return P, components, fractions


def extract_properties(components, fractions):
    Pc, Tc, w = [], [], []

    for comp in components:
        # Adjust column names based on your Excel
        Pc.append(comp["pc"])
        Tc.append(comp["tc"])
        w.append(comp["w"])

    x = fractions
    y = fractions

    return x, y, Pc, Tc, w


def calculate(P, x, y, Pc, Tc, w):
    T_n = 0.25          # match VBA exactly
    step = 0.25         # match VBA exactly
    iterations = 2500   # match VBA exactly
    TOL = 1e-300        # match VBA Tol

    best_dew_diff = float("inf")
    best_bubble_diff = float("inf")
    dew_point = None
    bubble_point = None

    for _ in range(iterations):
        k = [(Pc[i] / P) * math.exp(5.37 * (1 + w[i]) * (1 - Tc[i] / T_n))
             for i in range(len(x))]

        # Dew: use TOL threshold like VBA (if k < 1e-300, skip)
        dew_sum = sum(y[i] / k[i] for i in range(len(y)) if k[i] >= TOL)

        # Bubble: use exact zero check like VBA (if k == 0, skip)
        bubble_sum = sum(x[i] * k[i] for i in range(len(x)) if k[i] != 0)

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

    x, y, Pc, Tc, w = extract_properties(components, fractions)

    dew, bubble = calculate(P, x, y, Pc, Tc, w)

    print("\n--- Results ---")
    print(f"Dew Point: {dew - 273.15:.2f} °C")
    print(f"Bubble Point: {bubble - 273.15:.2f} °C")


if __name__ == "__main__":
    main()