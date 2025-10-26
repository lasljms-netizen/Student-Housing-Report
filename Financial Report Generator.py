import os
os.makedirs('charts', exist_ok=True)
os.makedirs('reports', exist_ok=True)

import pandas as pd
import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet

# === CONFIGURATION ===
EXCEL_FILE = "student_housing_data.xlsx"
OUTPUT_PDF = "reports/housing_financial_report.pdf"  # save in reports folder


# === STEP 1: LOAD DATA ===
def load_data(filepath):
    df = pd.read_excel(filepath, engine='openpyxl')
    return df


# === STEP 2: CALCULATE METRICS ===
def calculate_metrics(df, target_occupancy):
    total_units = len(df)
    occupied_units = df[df['Lease Status'].str.lower() == 'occupied'].shape[0]
    current_occupancy = occupied_units / total_units

    # ✅ changed TARGET_OCCUPANCY → target_occupancy
    needed_to_target = max(0, target_occupancy - current_occupancy)

    # ✅ changed TARGET_OCCUPANCY → target_occupancy
    needed_units = int((target_occupancy * total_units) - occupied_units)
    needed_units = max(0, needed_units)

    return total_units, occupied_units, current_occupancy, needed_to_target, needed_units


# === STEP 3: OCCUPANCY BY COLLEGE ===
def occupancy_by_college(df):
    college_counts = df.groupby('College')['Lease Status'].apply(
        lambda x: pd.Series(x).str.lower().eq('occupied').mean()
    ) * 100
    return college_counts


# === STEP 4: PLOT OCCUPANCY BY COLLEGE ===
def plot_occupancy(college_data):
    plt.figure(figsize=(6, 4))
    college_data.plot(kind='bar', color='teal')
    plt.title("Occupancy by College (%)")
    plt.ylabel("Occupancy Rate (%)")
    plt.xlabel("College")
    plt.ylim(0, 100)
    plt.tight_layout()
    chart_path = "charts/occupancy_by_college.png"  # save in charts folder
    plt.savefig(chart_path)
    plt.close()
    return chart_path


# === STEP 5: GENERATE PDF REPORT ===
def generate_pdf(total_units, occupied_units, current_occ, needed_to_target, needed_units, chart_path):
    doc = SimpleDocTemplate(OUTPUT_PDF, pagesize=letter)
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph("Student Housing Financial Report", styles['Title']))
    story.append(Spacer(1, 12))

    story.append(Paragraph(f"Total Units: {total_units}", styles['Normal']))
    story.append(Paragraph(f"Currently Occupied: {occupied_units}", styles['Normal']))
    story.append(Paragraph(f"Current Occupancy: {current_occ * 100:.2f}%", styles['Normal']))
    story.append(Paragraph(f"Occupancy Increase Needed: {needed_to_target * 100:.2f}%", styles['Normal']))
    story.append(Paragraph(f"Units Needed to Reach Goal: {needed_units}", styles['Normal']))
    story.append(Spacer(1, 12))

    story.append(Paragraph("Occupancy by College:", styles['Heading2']))
    story.append(Image(chart_path, width=400, height=300))

    doc.build(story)
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

    # ✅ FIXED: added target_occupancy to the function call
    total_units, occupied_units, current_occ, needed_to_target, needed_units = calculate_metrics(df, target_occupancy)

    college_data = occupancy_by_college(df)
    chart_path = plot_occupancy(college_data)
    generate_pdf(total_units, occupied_units, current_occ, needed_to_target, needed_units, chart_path)