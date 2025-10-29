import os
import pandas as pd
import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet

# === CONFIGURATION ===
EXCEL_FILE = r"C:\Users\laslj\OneDrive\Documents\Student_Housing_Data.xlsx" #issue number 2 - fixed
OUTPUT_PDF = "reports/housing_financial_report.pdf"  #save in reports folder

os.makedirs('charts', exist_ok=True)
os.makedirs('reports', exist_ok=True)


# === STEP 1: LOAD DATA ===
def load_data(filepath):
    df = pd.read_excel(filepath, engine='openpyxl', sheet_name=0)
    #Drop completely empty rows
    df = df.dropna(how='all')
    #Strip column names
    df.columns = df.columns.str.strip()
    return df


# === STEP 2: CALCULATE METRICS ===
def calculate_metrics(df, target_occupancy):
    #Standardize Lease Status
    df['Lease Status'] = df['Lease Status'].fillna('Unoccupied').str.strip().str.capitalize()

    total_units = len(df)  # counts all rows
    occupied_units = df[df['Lease Status'] == 'Leased'].shape[0]
    current_occupancy = occupied_units / total_units if total_units > 0 else 0

    needed_units = max(0, int((target_occupancy * total_units) - occupied_units))
    needed_to_target = max(0, target_occupancy - current_occupancy)

    return total_units, occupied_units, current_occupancy, needed_to_target, needed_units


# === STEP 3: OCCUPANCY BY COLLEGE ===
def occupancy_by_college(df):
    # Make sure College column is clean
    df['College'] = df['College'].fillna('N/A').str.strip()
    #Make sure College column is clean
    #Normalize college names (clean up spaces, capitalization)
    df['College'] = (
        df['College']
        .astype(str)
        .str.strip()
        .str.title()  #makes colleges consistent
    )

    total_units = len(df)  #total units in the building

    #issue number 1 - fixed
    #Count how many units belong to each college
    college_counts = df.groupby('College').size()

    #Divide by total units to get percentage
    occupancy_percent = (college_counts / total_units * 100).sort_index()

    return occupancy_percent

# === STEP 4: PLOT OCCUPANCY BY COLLEGE ===

def plot_occupancy(college_data):
    #Remove 'N/A' from the Series before plotting - data formatting issue - fixed
    chart_data = college_data[(college_data.index != '') & (college_data > 0)]

    plt.figure(figsize=(8, 5))

    #PDF polish - issue number 3
    #Define a consistent color for each college
    color_map = {
        'Depaul': '#001F4D',  #navy blue
        'Saic': '#8B0000',  # maroon
        'Columbia College': '#00BFFF',  # sky blue
        'Roosevelt': '#FFD700',  # gold
        'N/A': '#C0C0C0'  # gray (for debugging if needed)
    }

    colors = [color_map.get(college, '#7f7f7f') for college in college_data.index]

    bars = plt.bar(college_data.index, college_data.values, color=colors)

    for bar, value in zip(bars, college_data.values):
        plt.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height() + 1,
            f'{value:.1f}%',
            ha='center', va='bottom', fontsize=10, fontweight='bold'
        )

    plt.title("Occupancy by College", fontsize=14, fontweight='bold')
    plt.xlabel("College", fontsize=12)
    plt.ylabel("Occupancy (%)", fontsize=12)
    plt.ylim(0, 110)
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()

    chart_path = "occupancy_by_college.png"
    plt.savefig(chart_path)
    plt.close()

    return chart_path


# === STEP 5: GENERATE PDF REPORT ===
def generate_pdf(total_units, occupied_units, current_occ, needed_to_target, needed_units, chart_path, df,
                 college_data):
    doc = SimpleDocTemplate(OUTPUT_PDF, pagesize=letter)
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph("Student Housing Financial Report", styles['Title']))
    story.append(Spacer(1, 12))
    story.append(Paragraph(f"Target Occupancy: {target_occupancy * 100:.2f}%", styles['Normal']))
    story.append(Paragraph(f"Total Units: {total_units}", styles['Normal']))
    story.append(Paragraph(f"Currently Occupied: {occupied_units}", styles['Normal']))
    story.append(Paragraph(f"Current Occupancy: {current_occ * 100:.2f}%", styles['Normal']))
    story.append(Paragraph(f"Occupancy Increase Needed: {needed_to_target * 100:.2f}%", styles['Normal']))
    story.append(Paragraph(f"Units Needed to Reach Goal: {needed_units}", styles['Normal']))
    story.append(Spacer(1, 12))

    story.append(Paragraph("Occupancy by College:", styles['Heading2']))
    story.append(Image(chart_path, width=400, height=300))

    doc.build(story)

    print(df[['Unit', 'College', 'Lease Status']])
    print("Total units:", total_units)
    print("Occupied units:", occupied_units)
    print("Occupancy by college:")
    for college, occupancy in college_data.items():
        print(f"{college}: {occupancy:.2f}%")
    print(f"✅ Report generated: {OUTPUT_PDF}")


# === MAIN EXECUTION ===
if __name__ == "__main__":
    goal_input = input("Enter your target occupancy (e.g. 87 for 87%): ").strip()

    try:
        target_occupancy = float(goal_input)
        if target_occupancy > 1:
            target_occupancy = target_occupancy / 100
    except ValueError:
        print("Invalid input. Defaulting to 95%.")
        target_occupancy = 0.95

    df = load_data(EXCEL_FILE)

    total_units, occupied_units, current_occ, needed_to_target, needed_units = calculate_metrics(df, target_occupancy)
    college_data = occupancy_by_college(df)
    college_data.index = college_data.index.str.strip().str.title()
    college_data = college_data[college_data.index.notna()]  # remove blank index
    college_data = college_data[college_data.index != 'N/A']  # remove "N/A"df
    chart_path = plot_occupancy(college_data)
    generate_pdf(total_units, occupied_units, current_occ, needed_to_target, needed_units, chart_path, df, college_data)
