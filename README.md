# **Google Sheets Salary Management System**

## **Description**
This project is a fully automated salary management system designed for organizations to manage timesheets and payment reports efficiently. It integrates data from multiple worker spreadsheets and generates detailed reports, ensuring transparency for both workers and employers. 

### **Features**
- Aggregates worker data into a centralized `Data` sheet.
- Automatically calculates total work hours, costs, and payments.
- Generates a monthly summary report for all workers in the `Monthly` sheet.
- Creates personalized monthly reports for each worker in their respective spreadsheets.
- Logs errors and issues for easy troubleshooting.

A **template for the required spreadsheet structure** can be found [here](https://docs.google.com/spreadsheets/d/1iysf9szEguhfCSI2itDNnyUp2ZAhAXVetqwkQ4nIKqE/edit?gid=1557421726#gid=1557421726).

---

## **Project Components**

### **Sheets Overview**
1. **Main Spreadsheet**:
   - `Parameters`: List of workers and their details.
   - `Data`: Centralized database for all work records.
   - `Payments`: Record of payments made to workers.
   - `Monthly`: Overall monthly summary for all workers.

2. **Worker Spreadsheets**:
   - Each worker has their own spreadsheet with:
     - A `Data` sheet for individual work records.
     - A `Summary` sheet for monthly totals (automatically created if missing).

### **Scripts Overview**

#### **1. `collectAndSortWorkerData`**
- **Purpose**:
  - Collects data from individual workers' sheets.
  - Calculates work hours and costs.
  - Consolidates and sorts the data in the `Data` sheet.
- **Trigger**: Runs every 15 minutes.

#### **2. `generateMonthlyReport`**
- **Purpose**:
  - Generates a monthly summary for all workers in the `Monthly` sheet.
  - Includes:
    - `Deserves` (total payments due).
    - `Got` (total payments made).
    - `Difference` (balance).
  - Formats the `Difference` rows in bold for easy identification.
- **Trigger**: Runs daily or weekly.

#### **3. `generateMonthlyReports`**
- **Purpose**:
  - Creates individual monthly reports for each worker in their respective spreadsheets.
  - Aggregates monthly totals and unique workdays.
  - Writes to the "Summary" sheet in the worker’s spreadsheet.
- **Trigger**: Runs daily or weekly.

---

## **Deployment Instructions**

### **Step 1: Set Up the Template**
1. Open the provided **[Google Sheets Template](https://docs.google.com/spreadsheets/d/1iysf9szEguhfCSI2itDNnyUp2ZAhAXVetqwkQ4nIKqE/edit?gid=1557421726#gid=1557421726)**.
2. Make a copy for your organization.
3. Update the `Parameters` sheet with your workers' details, including hourly rates and sheet links.

### **Step 2: Add Scripts**
1. Open the main spreadsheet.
2. Navigate to **Extensions > Apps Script**.
3. Copy and paste the provided scripts (`collectAndSortWorkerData`, `generateMonthlyReport`, and `generateMonthlyReports`) into separate files.

### **Step 3: Configure Triggers**
1. Open the Apps Script Editor.
2. Go to **Triggers** and set up the following:
   - `collectAndSortWorkerData`: Trigger every 15 minutes.
   - `generateMonthlyReport`: Trigger daily or weekly.
   - `generateMonthlyReports`: Trigger daily or weekly.

---

## **How It Works**

1. **Data Collection**:
   - The system collects data from each worker's sheet and consolidates it into the `Data` sheet.
   - Calculates work duration, total costs, and travel expenses.

2. **Monthly Summary**:
   - Generates a comprehensive report in the `Monthly` sheet, showing payments due, made, and remaining for each worker.

3. **Individual Reports**:
   - Updates each worker's "Summary" sheet with their monthly totals, including total costs and unique workdays.

---

## **File Structure**

### **Main Spreadsheet**
- `Parameters`: List of workers, hourly rates, and settings for excluding workers from reports.
- `Data`: Centralized database for all work records.
- `Payments`: Records payments made to workers.
- `Monthly`: Monthly summary report for all workers.

### **Worker Spreadsheet**
- `Data`: Work records for the specific worker.
- `Summary`: Automatically created sheet for monthly summaries.

---

## **Troubleshooting**

### Common Issues
1. **Missing Data**:
   - Ensure all required columns are present in the `Parameters`, `Data`, and `Payments` sheets.
   - Verify that all worker spreadsheets are accessible and contain valid data.

2. **Script Errors**:
   - Check the Apps Script execution logs for detailed error messages.
   - Verify that column names match exactly with those expected by the scripts.

3. **Performance Issues**:
   - Reduce the frequency of triggers for large datasets.
   -
