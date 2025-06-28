import pandas as pd

TRUCK_CAPACITY = 3500

truck_types = [
    {"Type": "Mini Pickup", "Capacity": 1000, "Efficiency": 12},
    {"Type": "Light Truck A", "Capacity": 1500, "Efficiency": 10},
    {"Type": "Medium A", "Capacity": 2500, "Efficiency": 8},
    {"Type": "Standard A", "Capacity": 3000, "Efficiency": 7},
    {"Type": "Standard B", "Capacity": 3500, "Efficiency": 6.5},
]

city_distance = {
    "RAJKOT": 0,
    "AHMEDABAD": 215,
    "GONDAL": 40,
    "BHAVNAGAR": 170,
    "JAMNAGAR": 90,
    "DWARKA": 230,
    "BHUJ": 270,
    "MUNDRA": 240,
    "BARMER RJ": 370,
    "FARRUKHNAGAR HR": 1100,
    "ADHEWADA": 160,
}


def assign_zone(city: str) -> str:
    if not isinstance(city, str):
        return "Other"
    city = city.upper()
    gujarat_core_cities = [
        "AHMEDABAD", "GONDAL", "RAJKOT", "BHAVNAGAR", "JAMNAGAR",
        "DHROL", "JASDAN", "MORBI", "JETPUR", "VANTHALI",
        "MOTIKHAVDI", "KHAMBHALIA", "KHAMBHALIYA", "KADI",
        "KALOL", "VIJAPUR", "VIRAMGAM", "SANAND", "KHEDA",
        "SAVA", "PRANTIJ",
    ]
    if any(x in city for x in gujarat_core_cities):
        return "Gujarat Core"

    kutch_cities = [
        "BHACHAU", "MUNDRA", "BHATIA", "DWARKA", "OKHA", "BHUJ",
        "DAYAPAR", "ADIPUR", "NALIYA", "MITHAPUR", "LAKHTAR",
    ]
    if any(x in city for x in kutch_cities):
        return "Kutch/Saurashtra"

    saurashtra_cities = [
        "PORBANDAR", "VERAVAL", "KODINAR", "MANAVADAR", "BANTVA",
        "JAM KALYANPUR", "JAM RAVAL", "TALAJA", "UNA",
    ]
    if any(x in city for x in saurashtra_cities):
        return "Saurashtra"

    north_gujarat_cities = ["PALANPUR", "DHANERA", "TALOD", "THEBA", "ADHEWADA"]
    if any(x in city for x in north_gujarat_cities):
        return "North Gujarat"

    if "RJ" in city or any(x in city for x in ["BARMER", "BALESAR", "JODHPUR", "SANCHORE"]):
        return "Rajasthan"

    haryana_cities = [
        "FARRUKHNAGAR", "HANSI", "DHIGAWA", "BHUNA", "BAHAL",
        "NARNAUL", "PANIPAT", "SATNALI", "SIRSA", "TOHANA",
        "UKLANA MANDI", "YAMUNANAGAR",
    ]
    if city.endswith(" HR") or any(x in city for x in haryana_cities):
        return "Haryana"

    if "NEW DELHI" in city:
        return "Delhi"

    return "Other"


def best_fit_decreasing(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["Zone"] = df["Party_city"].apply(assign_zone)

    assignments = []
    truck_counter = 1
    for zone, zone_df in df.groupby("Zone"):
        orders = zone_df.sort_values("Cubic_feet_order_size", ascending=False)
        trucks = []
        for _, row in orders.iterrows():
            vol = row["Cubic_feet_order_size"]
            best_index = None
            min_space = None
            for i, t in enumerate(trucks):
                if t["remaining"] >= vol:
                    space = t["remaining"] - vol
                    if min_space is None or space < min_space:
                        best_index = i
                        min_space = space
            if best_index is None:
                trucks.append({"remaining": TRUCK_CAPACITY, "orders": []})
                best_index = len(trucks) - 1
            truck = trucks[best_index]
            truck["orders"].append(row)
            truck["remaining"] -= vol
        for t in trucks:
            for row in t["orders"]:
                assignments.append(
                    {
                        "Truck_ID": f"T{truck_counter:03}",
                        "Zone": row["Zone"],
                        "Order_No": row["order_no"],
                        "Party_Name": row["Party_name"],
                        "Party_City": row["Party_city"],
                        "Used_Volume": row["Cubic_feet_order_size"],
                    }
                )
            truck_counter += 1
    return pd.DataFrame(assignments)


def summarize(result_df: pd.DataFrame) -> pd.DataFrame:
    summary = (
        result_df.groupby("Zone")
        .agg(
            Number_of_Trucks=("Truck_ID", lambda x: x.nunique()),
            Orders_Count=("Order_No", "count"),
            Total_Volume_cubic_feet=("Used_Volume", "sum"),
        )
        .reset_index()
    )
    return summary


def main(path: str = "PRE_ORDER.xlsx") -> None:
    df = pd.read_excel(path)
    result = best_fit_decreasing(df)
    summary = summarize(result)
    result.to_excel("best_fit_output.xlsx", index=False)
    summary.to_excel("best_fit_summary.xlsx", index=False)
    print("=== BEST FIT SUMMARY ===")
    print(summary)
    print(f"Total trucks: {result['Truck_ID'].nunique()}")
    print(f"Total orders: {len(result)}")


if __name__ == "__main__":
    main()
