# import pandas as pd
# from openpyxl import Workbook
# from openpyxl.utils.dataframe import dataframe_to_rows
#
# # === CONFIGURATION ===
# truck_types = [
#     {"Type": "Mini Pickup", "Capacity": 1000, "Efficiency": 12},
#     {"Type": "Light Truck A", "Capacity": 1500, "Efficiency": 10},
#     {"Type": "Medium A", "Capacity": 2500, "Efficiency": 8},
#     {"Type": "Standard A", "Capacity": 3000, "Efficiency": 7},
#     {"Type": "Standard B", "Capacity": 3500, "Efficiency": 6.5},
# ]
#
# city_distance = {
#     "RAJKOT": 0,
#     "AHMEDABAD": 215,
#     "GONDAL": 40,
#     "BHAVNAGAR": 170,
#     "JAMNAGAR": 90,
#     "SURAT": 370,
#     "DWARKA": 230,
#     "BHUJ": 270,
#     "MUNDRA": 240,
#     "BARMER RJ": 370,
#     "FARRUKHNAGAR HR": 1100,
#     "ADHEWADA": 160
# }
#
# # === LOAD DATA ===
# df = pd.read_excel("PRE_ORDER.xlsx")
#
#
# # === ZONE MAPPING ===
# def assign_zone(city):
#     if pd.isna(city):
#         return "Other"
#     city = city.upper()
#     if any(x in city for x in ["AHMEDABAD", "GONDAL", "RAJKOT", "BHAVNAGAR", "SURENDRANAGAR", "JAMNAGAR"]):
#         return "Gujarat Core"
#     elif any(x in city for x in ["BHACHAU", "MUNDRA", "MANDVI", "BHATIA", "DWARKA", "OKHA", "BHUJ", "DAYAPAR"]):
#         return "Kutch/Saurashtra"
#     elif "RJ" in city or any(x in city for x in ["BARMER", "BALESAR"]):
#         return "Rajasthan"
#     elif "HR" in city or any(x in city for x in ["FARRUKHNAGAR", "HANSI", "DHIGAWA", "BHUNA"]):
#         return "Haryana"
#     else:
#         return "Other"
#
#
# df["Zone"] = df["Party_city"].apply(assign_zone)
#
# # === PER ZONE OPTIMIZATION ===
# assignments = []
# truck_id_counter = 1
#
# for zone, zone_df in df.groupby("Zone"):
#     orders = zone_df.sort_values(by="Cubic_feet_order_size", ascending=False)
#
#     while not orders.empty:
#         truck_orders = []
#         used_volume = 0
#         city_group = []
#
#         for idx, row in orders.iterrows():
#             vol = row["Cubic_feet_order_size"]
#             city = str(row["Party_city"]).upper()
#
#             # Oversized order - split it
#             if vol > 3500:
#                 parts = []
#                 remaining = vol
#                 while remaining > 0:
#                     part_volume = min(remaining, 3500)
#                     parts.append(part_volume)
#                     remaining -= part_volume
#                 for pv in parts:
#                     truck = sorted([t for t in truck_types if t["Capacity"] >= pv], key=lambda x: -x["Efficiency"])[0]
#                     dist = city_distance.get(city, 300)
#                     fuel = round(dist / truck["Efficiency"], 2)
#                     assignments.append({
#                         "Truck_ID": f"TRUCK-{truck_id_counter:03}",
#                         "Truck_Type": truck["Type"],
#                         "Used_Volume": pv,
#                         "Zone": zone,
#                         "Party_Name": row["Party_name"],
#                         "Party_City": city,
#                         "Order_No": row["order_no"],
#                         "Estimated_Distance_km": dist,
#                         "Estimated_Fuel_L": fuel,
#                         "Is_Split": True
#                     })
#                     truck_id_counter += 1
#                 orders = orders.drop(idx)
#                 break  # start next truck loop fresh
#
#             if used_volume + vol <= 3500:
#                 truck_orders.append(row)
#                 used_volume += vol
#                 city_group.append(city)
#                 orders = orders.drop(idx)
#
#         # Assign a truck for this batch
#         if truck_orders:
#             best_truck = \
#             sorted([t for t in truck_types if t["Capacity"] >= used_volume], key=lambda x: -x["Efficiency"])[0]
#             max_dist = max([city_distance.get(c, 300) for c in city_group])
#             fuel = round(max_dist / best_truck["Efficiency"], 2)
#
#             for row in truck_orders:
#                 assignments.append({
#                     "Truck_ID": f"TRUCK-{truck_id_counter:03}",
#                     "Truck_Type": best_truck["Type"],
#                     "Used_Volume": used_volume,
#                     "Zone": zone,
#                     "Party_Name": row["Party_name"],
#                     "Party_City": row["Party_city"],
#                     "Order_No": row["order_no"],
#                     "Estimated_Distance_km": max_dist,
#                     "Estimated_Fuel_L": fuel,
#                     "Is_Split": False
#                 })
#             truck_id_counter += 1
#
# # === CREATE ADDITIONAL TABLES ===
#
# # 1. Zone Cities Table
# zone_cities_data = []
# for zone, zone_df in df.groupby("Zone"):
#     cities = zone_df["Party_city"].dropna().unique()
#     for city in cities:
#         zone_cities_data.append({
#             "Zone": zone,
#             "City": city,
#             "Distance_from_Rajkot_km": city_distance.get(str(city).upper(), 300)
#         })
#
# zone_cities_df = pd.DataFrame(zone_cities_data).sort_values(["Zone", "City"])
#
# # 2. Truck Types Table
# truck_types_df = pd.DataFrame(truck_types)
# truck_types_df["Capacity_cubic_feet"] = truck_types_df["Capacity"]
# truck_types_df["Fuel_Efficiency_km_per_liter"] = truck_types_df["Efficiency"]
# truck_types_df = truck_types_df[["Type", "Capacity_cubic_feet", "Fuel_Efficiency_km_per_liter"]]
#
# # === EXPORT TO EXCEL WITH MULTIPLE SHEETS ===
# result_df = pd.DataFrame(assignments)
#
# # Create Excel file with multiple sheets
# with pd.ExcelWriter("Optimized_By_Zone_Fuel_Enhanced.xlsx", engine='openpyxl') as writer:
#     # Main optimization results
#     result_df.to_excel(writer, sheet_name='Truck_Assignments', index=False)
#
#     # Zone Cities table
#     zone_cities_df.to_excel(writer, sheet_name='Zone_Cities', index=False)
#
#     # Truck Types table
#     truck_types_df.to_excel(writer, sheet_name='Truck_Types', index=False)
#
#     # Summary statistics
#     summary_data = []
#     for zone in result_df['Zone'].unique():
#         zone_data = result_df[result_df['Zone'] == zone]
#         unique_trucks = len(zone_data['Truck_ID'].unique())
#         total_volume = zone_data['Used_Volume'].sum()
#         total_fuel = zone_data['Estimated_Fuel_L'].sum()
#         total_distance = zone_data['Estimated_Distance_km'].sum()
#
#         summary_data.append({
#             'Zone': zone,
#             'Number_of_Trucks': unique_trucks,
#             'Total_Volume_cubic_feet': total_volume,
#             'Total_Fuel_Required_L': round(total_fuel, 2),
#             'Total_Distance_km': total_distance,
#             'Orders_Count': len(zone_data)
#         })
#
#     summary_df = pd.DataFrame(summary_data)
#     summary_df.to_excel(writer, sheet_name='Summary_by_Zone', index=False)
#
# print("Enhanced file exported to: Optimized_By_Zone_Fuel_Enhanced.xlsx")
# print("\nFile contains 4 sheets:")
# print("1. Truck_Assignments - Main optimization results")
# print("2. Zone_Cities - All zones with their cities and distances")
# print("3. Truck_Types - Truck specifications")
# print("4. Summary_by_Zone - Summary statistics by zone")
#
# # Display summary
# print(f"\nOptimization Summary:")
# print(f"Total trucks assigned: {len(result_df['Truck_ID'].unique())}")
# print(f"Total orders processed: {len(result_df)}")
# print(f"Total estimated fuel required: {result_df['Estimated_Fuel_L'].sum():.2f} liters")
# print(f"Zones covered: {', '.join(result_df['Zone'].unique())}")

import pandas as pd

# === CONFIGURATION ===
truck_types = [
    {"Type": "Mini Pickup", "Capacity": 1000, "Efficiency": 12},
    {"Type": "Light Truck A", "Capacity": 1500, "Efficiency": 10},
    {"Type": "Medium A", "Capacity": 2500, "Efficiency": 8},
    {"Type": "Standard A", "Capacity": 3000, "Efficiency": 7},
    {"Type": "Standard B", "Capacity": 3500, "Efficiency": 6.5},
]

# Enhanced city distances based on your actual data
city_distance = {
    "RAJKOT": 0, "AHMEDABAD": 215, "GONDAL": 40, "BHAVNAGAR": 170, "JAMNAGAR": 90,
    "DWARKA": 230, "BHUJ": 270, "MUNDRA": 240, "BARMER RJ": 370, "FARRUKHNAGAR HR": 1100,
    "ADHEWADA": 160, "ADIPUR": 250, "BAHAL HR": 950, "BALESAR SATTA RJ": 320,
    "BANTVA": 120, "BHACHAU": 280, "BHATIA": 260, "BHUNA HR": 1000, "DAYAPAR": 300,
    "DHANERA": 200, "DHIGAWA HR": 1050, "DHROL": 85, "HANSI HR": 1020,
    "JAM KALYANPUR": 130, "JAM RAVAL": 140, "JASDAN": 110, "JETPUR": 150,
    "JODHPUR RJ": 420, "KADI": 190, "KALOL": 180, "KHAMBHALIA": 100,
    "KHAMBHALIYA": 100, "KHEDA": 200, "KODINAR": 200, "LAKHTAR": 280,
    "MANAVADAR": 160, "MITHAPUR": 220, "MORBI": 140, "MOTIKHAVDI": 95,
    "NALIYA": 310, "NARNAUL HR": 980, "NEW DELHI(EAST)": 1150,
    "NEW DELHI(SOUTH WEST)": 1150, "NEW DELHI(WEST)": 1150, "OKHA": 240,
    "PALANPUR": 180, "PANIPAT HR": 1080, "PORBANDAR": 180, "PRANTIJ": 160,
    "SANAND": 200, "SANCHORE RJ": 250, "SATNALI HR": 1000, "SAVA": 190,
    "SIRSA HR": 1100, "TALAJA": 180, "TALOD": 170, "THEBA": 160,
    "TOHANA HR": 1120, "UKLANA MANDI HR": 1050, "UNA": 170, "VANTHALI": 160,
    "VERAVAL": 190, "VIJAPUR": 160, "VIRAMGAM": 170, "YAMUNANAGAR HR": 1130
}

# === LOAD DATA ===
df = pd.read_excel("PRE_ORDER.xlsx")


# === ENHANCED ZONE MAPPING ===
def assign_zone(city):
    if pd.isna(city):
        return "Other"
    city = city.upper()

    # Gujarat Core - main Gujarat cities near Rajkot/Ahmedabad
    gujarat_core_cities = ["AHMEDABAD", "GONDAL", "RAJKOT", "BHAVNAGAR", "JAMNAGAR", "DHROL",
                           "JASDAN", "MORBI", "JETPUR", "VANTHALI", "MOTIKHAVDI", "KHAMBHALIA",
                           "KHAMBHALIYA", "KADI", "KALOL", "VIJAPUR", "VIRAMGAM", "SANAND",
                           "KHEDA", "SAVA", "PRANTIJ"]
    if any(x in city for x in gujarat_core_cities):
        return "Gujarat Core"

    # Kutch/Saurashtra - Kutch region and northern coast
    kutch_cities = ["BHACHAU", "MUNDRA", "BHATIA", "DWARKA", "OKHA", "BHUJ", "DAYAPAR",
                    "ADIPUR", "NALIYA", "MITHAPUR", "LAKHTAR"]
    if any(x in city for x in kutch_cities):
        return "Kutch/Saurashtra"

    # Saurashtra Peninsula - southern Gujarat coast
    saurashtra_cities = ["PORBANDAR", "VERAVAL", "KODINAR", "MANAVADAR", "BANTVA",
                         "JAM KALYANPUR", "JAM RAVAL", "TALAJA", "UNA"]
    if any(x in city for x in saurashtra_cities):
        return "Saurashtra"

    # North Gujarat
    north_gujarat_cities = ["PALANPUR", "DHANERA", "TALOD", "THEBA", "ADHEWADA"]
    if any(x in city for x in north_gujarat_cities):
        return "North Gujarat"

    # Rajasthan
    if "RJ" in city or any(x in city for x in ["BARMER", "BALESAR", "JODHPUR", "SANCHORE"]):
        return "Rajasthan"

    # Haryana - only cities ending with " HR" or specific Haryana cities
    haryana_cities = ["FARRUKHNAGAR", "HANSI", "DHIGAWA", "BHUNA", "BAHAL",
                      "NARNAUL", "PANIPAT", "SATNALI", "SIRSA", "TOHANA",
                      "UKLANA MANDI", "YAMUNANAGAR"]
    if city.endswith(" HR") or any(x in city for x in haryana_cities):
        return "Haryana"

    # Delhi
    if "NEW DELHI" in city:
        return "Delhi"

    return "Other"


df["Zone"] = df["Party_city"].apply(assign_zone)

# === PER ZONE OPTIMIZATION ===
assignments = []
truck_id_counter = 1

for zone, zone_df in df.groupby("Zone"):
    print(f"Optimizing zone: {zone} ({len(zone_df)} orders)")
    orders = zone_df.sort_values(by="Cubic_feet_order_size", ascending=False)

    while not orders.empty:
        truck_orders = []
        used_volume = 0
        city_group = []

        for idx, row in orders.iterrows():
            vol = row["Cubic_feet_order_size"]
            city = str(row["Party_city"]).upper()

            # Oversized order - split it
            if vol > 3500:
                parts = []
                remaining = vol
                while remaining > 0:
                    part_volume = min(remaining, 3500)
                    parts.append(part_volume)
                    remaining -= part_volume
                for pv in parts:
                    truck = sorted([t for t in truck_types if t["Capacity"] >= pv], key=lambda x: -x["Efficiency"])[0]
                    dist = city_distance.get(city, 300)
                    fuel = round(dist / truck["Efficiency"], 2)
                    assignments.append({
                        "Truck_ID": f"TRUCK-{truck_id_counter:03}",
                        "Truck_Type": truck["Type"],
                        "Used_Volume": pv,
                        "Zone": zone,
                        "Party_Name": row["Party_name"],
                        "Party_City": row["Party_city"],
                        "Order_No": row["order_no"],
                        "Estimated_Distance_km": dist,
                        "Estimated_Fuel_L": fuel,
                        "Is_Split": True
                    })
                    truck_id_counter += 1
                orders = orders.drop(idx)
                break  # start next truck loop fresh

            if used_volume + vol <= 3500:
                truck_orders.append(row)
                used_volume += vol
                city_group.append(city)
                orders = orders.drop(idx)

        # Assign a truck for this batch
        if truck_orders:
            best_truck = \
            sorted([t for t in truck_types if t["Capacity"] >= used_volume], key=lambda x: -x["Efficiency"])[0]
            max_dist = max([city_distance.get(c, 300) for c in city_group])
            fuel = round(max_dist / best_truck["Efficiency"], 2)

            for row in truck_orders:
                assignments.append({
                    "Truck_ID": f"TRUCK-{truck_id_counter:03}",
                    "Truck_Type": best_truck["Type"],
                    "Used_Volume": used_volume,
                    "Zone": zone,
                    "Party_Name": row["Party_name"],
                    "Party_City": row["Party_city"],
                    "Order_No": row["order_no"],
                    "Estimated_Distance_km": max_dist,
                    "Estimated_Fuel_L": fuel,
                    "Is_Split": False
                })
            truck_id_counter += 1

# === CREATE ADDITIONAL TABLES ===

# 1. Zone Cities Table
zone_cities_data = []
for zone, zone_df in df.groupby("Zone"):
    cities = zone_df["Party_city"].dropna().unique()
    for city in cities:
        zone_cities_data.append({
            "Zone": zone,
            "City": city,
            "Distance_from_Rajkot_km": city_distance.get(str(city).upper(), 300)
        })

zone_cities_df = pd.DataFrame(zone_cities_data).sort_values(["Zone", "City"])

# 2. Truck Types Table
truck_types_df = pd.DataFrame(truck_types)
truck_types_df["Capacity_cubic_feet"] = truck_types_df["Capacity"]
truck_types_df["Fuel_Efficiency_km_per_liter"] = truck_types_df["Efficiency"]
truck_types_df = truck_types_df[["Type", "Capacity_cubic_feet", "Fuel_Efficiency_km_per_liter"]]

# 3. Summary by Zone
result_df = pd.DataFrame(assignments)
summary_data = []
for zone in result_df['Zone'].unique():
    zone_data = result_df[result_df['Zone'] == zone]
    unique_trucks = len(zone_data['Truck_ID'].unique())
    total_volume = zone_data['Used_Volume'].sum()
    total_fuel = zone_data['Estimated_Fuel_L'].sum()
    total_distance = zone_data['Estimated_Distance_km'].sum()

    summary_data.append({
        'Zone': zone,
        'Number_of_Trucks': unique_trucks,
        'Total_Volume_cubic_feet': total_volume,
        'Total_Fuel_Required_L': round(total_fuel, 2),
        'Average_Distance_km': round(total_distance / len(zone_data)),
        'Orders_Count': len(zone_data)
    })

summary_df = pd.DataFrame(summary_data)

# === EXPORT TO EXCEL WITH MULTIPLE SHEETS ===
with pd.ExcelWriter("Optimized_By_Zone_Fuel_Enhanced.xlsx", engine='openpyxl') as writer:
    # Main optimization results
    result_df.to_excel(writer, sheet_name='Truck_Assignments', index=False)

    # Zone Cities table
    zone_cities_df.to_excel(writer, sheet_name='Zone_Cities', index=False)

    # Truck Types table
    truck_types_df.to_excel(writer, sheet_name='Truck_Types', index=False)

    # Summary statistics
    summary_df.to_excel(writer, sheet_name='Summary_by_Zone', index=False)

print("Enhanced file exported to: Optimized_By_Zone_Fuel_Enhanced.xlsx")
print("\nFile contains 4 sheets:")
print("1. Truck_Assignments - Main optimization results")
print("2. Zone_Cities - All zones with their cities and distances")
print("3. Truck_Types - Truck specifications")
print("4. Summary_by_Zone - Summary statistics by zone")

# Display summary
print("\n=== OPTIMIZATION SUMMARY ===")
print(f"Total trucks assigned: {len(result_df['Truck_ID'].unique())}")
print(f"Total orders processed: {len(result_df)}")
print(f"Total estimated fuel required: {result_df['Estimated_Fuel_L'].sum():.2f} liters")
print(f"Zones covered: {', '.join(sorted(result_df['Zone'].unique()))}")

print("\n=== ZONE BREAKDOWN ===")
for _, row in summary_df.iterrows():
    print(
        f"{row['Zone']}: {row['Number_of_Trucks']} trucks, {row['Orders_Count']} orders, {row['Total_Fuel_Required_L']}L fuel")

print("\n=== EFFICIENCY ANALYSIS ===")
summary_df['Fuel_per_Order'] = summary_df['Total_Fuel_Required_L'] / summary_df['Orders_Count']
most_efficient = summary_df.loc[summary_df['Fuel_per_Order'].idxmin()]
least_efficient = summary_df.loc[summary_df['Fuel_per_Order'].idxmax()]
print(f"Most efficient zone: {most_efficient['Zone']} ({most_efficient['Fuel_per_Order']:.2f}L per order)")
print(f"Least efficient zone: {least_efficient['Zone']} ({least_efficient['Fuel_per_Order']:.2f}L per order)")

print("\n=== TRUCK TYPE USAGE ===")
truck_usage = result_df['Truck_Type'].value_counts()
for truck_type, count in truck_usage.items():
    print(f"{truck_type}: {count} assignments")

print("\nOptimization complete! ✅")
