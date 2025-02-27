const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 5000;

// Middleware
app.use(cors());
app.use(bodyParser.json());

// Excel file path
const EXCEL_FILE = path.join(__dirname, 'feedback_data.xlsx');

// Define the correct column headers
const HEADERS = [
  'Full Name',
  'Email Address',
  'Phone Number',
  'Ticket Number',
  'Shop Location',
  'Other Location',
  'Product Quality',
  'Product Packaging',
  'Customer Service',
  'Comments',
  'Submission Date'
];

// Initialize Excel file if it doesn't exist
const initExcelFile = () => {
  if (!fs.existsSync(EXCEL_FILE)) {
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet([HEADERS]);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Feedback');
    XLSX.writeFile(workbook, EXCEL_FILE);
    console.log('Excel file initialized');
  }
};

// Add feedback in correct column order
const addFeedbackToExcel = (feedbackData) => {
  try {
    const workbook = XLSX.readFile(EXCEL_FILE);
    const worksheet = workbook.Sheets['Feedback'];
    const jsonData = XLSX.utils.sheet_to_json(worksheet);

    // Ensure feedback is saved in correct order
    const row = HEADERS.map(header => feedbackData[mapHeaderToField(header)] || '');
    row[row.length - 1] = new Date().toLocaleString(); // Add submission date

    XLSX.utils.sheet_add_aoa(worksheet, [row], { origin: -1 });
    XLSX.writeFile(workbook, EXCEL_FILE);
    return true;
  } catch (error) {
    console.error('Error saving feedback:', error);
    return false;
  }
};

// Map headers to feedbackData fields
const mapHeaderToField = (header) => {
  const mapping = {
    'Full Name': 'fullName',
    'Email Address': 'email',
    'Phone Number': 'phoneNumber',
    'Ticket Number': 'ticketNumber',
    'Shop Location': 'shopLocation',
    'Other Location': 'otherLocation',
    'Product Quality': 'productQuality',
    'Product Packaging': 'productPackaging',
    'Customer Service': 'customerService',
    'Comments': 'comments',
    'Submission Date': 'submissionDate'
  };
  return mapping[header];
};

// Initialize Excel file on server start
initExcelFile();

// API to save feedback
app.post('/api/feedback', (req, res) => {
  const feedbackData = req.body;

  if (!feedbackData.fullName || !feedbackData.email || !feedbackData.phoneNumber || !feedbackData.ticketNumber || !feedbackData.shopLocation) {
    return res.status(400).json({ success: false, message: 'Missing required fields' });
  }

  if (feedbackData.shopLocation === 'other' && !feedbackData.otherLocation) {
    return res.status(400).json({ success: false, message: 'Other location is required' });
  }

  const success = addFeedbackToExcel(feedbackData);

  if (success) {
    res.status(200).json({ success: true, message: 'Feedback saved successfully' });
  } else {
    res.status(500).json({ success: false, message: 'Failed to save feedback' });
  }
});

// API to get feedback
app.get('/api/feedback', (req, res) => {
  try {
    const workbook = XLSX.readFile(EXCEL_FILE);
    const worksheet = workbook.Sheets['Feedback'];
    const jsonData = XLSX.utils.sheet_to_json(worksheet);
    res.status(200).json({ success: true, data: jsonData });
  } catch (error) {
    console.error('Error reading Excel file:', error);
    res.status(500).json({ success: false, message: 'Failed to fetch feedback data' });
  }
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
