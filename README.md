# Timesheet Template Generator for Power BI

A Python script that generates a user-friendly Excel timesheet template optimized for Power BI reporting and analysis.

## 🎯 Features

- **12 Monthly Worksheets**: Separate sheets for June 2025 through May 2026
- **Multiple Rows Per Day**: Support for different projects and activities on the same day
- **Pre-populated Schedule**: Weekdays automatically filled with Part Time Work (08:00-13:00)
- **Weekend Detection**: Saturdays and Sundays automatically marked as "Weekend"
- **Time Formatting**: All time columns formatted as HH:MM for easy data entry
- **Automatic Calculations**: Total work hours calculated automatically
- **Power BI Ready**: Clean, flat structure perfect for Power BI import
- **Projects Reference**: Separate sheet listing all available projects

## 📋 Template Structure

### Columns
- **Date**: Pre-populated for all days
- **Employee Name**: Pre-filled with "Vivian"
- **Project**: Dropdown with project options
- **Task Description**: Free text field
- **Work Start/End**: Time fields for work hours
- **Break Start/End**: Time fields for break periods
- **Gym Start/End**: Time fields for gym sessions
- **Commute Start/End**: Time fields for commute time
- **Total Work Hours**: Automatically calculated

### Daily Structure
- **Weekdays**: 3 rows per day
  - Row 1: Part Time Work (pre-filled 08:00-13:00)
  - Row 2: Self Employment (blank for user input)
  - Row 3: Additional entry (blank for any other activities)
- **Weekends**: 1 row per day (marked as "Weekend")

## 🚀 Quick Start

### Prerequisites
- Python 3.7 or higher
- openpyxl library

### Installation
1. Clone this repository:
```bash
git clone https://github.com/yourusername/timesheet-template-generator.git
cd timesheet-template-generator
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Generate the template:
```bash
python timesheet_generator.py
```

### Usage
1. Open the generated `Timesheet_Template_For_PowerBI.xlsx` file
2. Navigate to any month tab (June 2025 - May 2026)
3. Enter your time data in the appropriate columns
4. Import into Power BI for reporting and analysis

## 🔧 Customization

### Adding Dropdowns (Optional)
To add project dropdowns manually in Excel:
1. Select the Project column (C)
2. Go to **Data** > **Data Validation**
3. Choose **"List"**
4. Enter: `Part Time Work,Self Employment,Vivian Portfolio,Break,Leave,Weekend`

### Modifying Projects
Edit the `projects` list in `timesheet_generator.py` to add or remove project options.

### Changing Time Schedule
Modify the pre-populated times in the script (currently 08:00-13:00 for Part Time Work).

## 📊 Power BI Integration

The template is designed for seamless Power BI integration:

1. **Import Data**: Use "Get Data" > "Excel" in Power BI
2. **Select Sheets**: Choose individual months or combine all
3. **Transform**: Use Power Query to merge data if needed
4. **Create Reports**: Build dashboards and visualizations

### Recommended Power BI Measures
- Total Work Hours by Project
- Daily/Weekly/Monthly Work Patterns
- Project Time Distribution
- Break Time Analysis
- Productivity Trends

## 🛠️ Technical Details

### Dependencies
- **openpyxl**: Excel file manipulation
- **datetime**: Date and time handling
- **calendar**: Calendar calculations

### File Structure
```
timesheet-template-generator/
├── timesheet_generator.py    # Main script
├── requirements.txt          # Python dependencies
├── README.md                # This file
└── Timesheet_Template_For_PowerBI.xlsx  # Generated template
```

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 👤 Author

**Vivian Ferguson**
- Created: June 2025
- Purpose: Personal timesheet management and Power BI reporting

## 🙏 Acknowledgments

- Built with openpyxl for Excel file generation
- Designed for Power BI reporting and analysis
- Optimized for user-friendly data entry

---

**Note**: This template is designed for personal use but can be adapted for team environments by modifying the employee name and project structure. 