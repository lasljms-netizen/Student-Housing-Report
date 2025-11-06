**Student Housing Financial Report Generator**

A Python-based tool that reads student housing data from an Excel file and automatically generates a professional PDF report showing current occupancy, progress toward goals, and occupancy rates by college (with charts).

**Features**
- Calculates total and occupied units
- Shows current occupancy percentage
- Compares to user-defined target occupancy
- Breaks down occupancy by college (SAIC, Columbia, Roosevelt, DePaul)

**Requirements**
pip install pandas openpyxl matplotlib reportlab

**How to Run**
- Open the project in your IDE of choice(PyCharm was used for this project).
- Make sure your Excel file is in the same folder as the Python script.
- Run the script.
- Enter your target occupancy when prompted (e.g. 95 for 95%).

_Your PDF report will appear in the reports folder._
Example: housing_financial_report.pdf


⚠️ Common Issues ⚠️

All 100% Occupancy: Check for typos or spacing in “Lease Status.”

Missing Colleges: Make sure “College” column matches exactly.

Grey Bars: Color mapping depends on exact college names — check capitalization and spaces.
