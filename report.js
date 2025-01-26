// Include the xlsx library
const XLSX = require('xlsx');

// Function to read Excel file and generate dashboard data
function generateDashboardData(excelFilePath) {
    // Read the Excel file
    const workbook = XLSX.readFile(excelFilePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Convert sheet to JSON
    const data = XLSX.utils.sheet_to_json(sheet);

    // Initialize dashboard data
    const dashboardData = {
        totalAuditors: 0,
        activeAudits: 0,
        completedAudits: 0,
        pendingReview: 0,
        recentAudits: [],
        auditorPerformance: {}
    };

    // Process data to populate dashboard
    data.forEach(row => {
        // Update total auditors
        if (!dashboardData.auditorPerformance[row.Auditor]) {
            dashboardData.auditorPerformance[row.Auditor] = { activeAudits: 0, completedAudits: 0 };
            dashboardData.totalAuditors++;
        }

        // Update audit status counts
        if (row.Status === 'In Progress') {
            dashboardData.activeAudits++;
            dashboardData.auditorPerformance[row.Auditor].activeAudits++;
        } else if (row.Status === 'Completed') {
            dashboardData.completedAudits++;
            dashboardData.auditorPerformance[row.Auditor].completedAudits++;
        } else if (row.Status === 'Pending Review') {
            dashboardData.pendingReview++;
        }

        // Add to recent audits
        dashboardData.recentAudits.push({
            auditName: row['Audit Name'],
            auditor: row.Auditor,
            status: row.Status,
            date: row.Date
        });
    });

    return dashboardData;
}

// Example usage
const excelFilePath = 'path/to/your/excel/file.xlsx';
const dashboardData = generateDashboardData(excelFilePath);
console.log(dashboardData);