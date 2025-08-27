# Excel Data Processor

A React-based web application that allows users to upload Excel files, add company data, configure column mappings, and download updated Excel files with the new data.

## Features

- 📁 **File Upload**: Drag & drop or click to upload Excel files (.xlsx, .xls)
- 🏢 **Company Data Management**: Add, edit, and remove company information
- ⚙️ **Column Mapping**: Configure which Excel columns correspond to which data fields
- 🔄 **Excel Processing**: Update Excel files with company data while preserving formatting
- 📥 **Download**: Download the processed Excel file
- 🎨 **Modern UI**: Beautiful, responsive interface built with React and Tailwind CSS

## Tech Stack

### Backend

- **Node.js** with Express.js
- **ExcelJS** for Excel file processing
- **Multer** for file upload handling
- **CORS** for cross-origin requests

### Frontend

- **React 18** with functional components and hooks
- **Tailwind CSS** for styling
- **React Dropzone** for file uploads
- **React Hot Toast** for notifications
- **Lucide React** for icons
- **Axios** for API calls

## Installation

1. **Clone or download the project**

2. **Install backend dependencies**

   ```bash
   npm install
   ```

3. **Install frontend dependencies**
   ```bash
   cd client
   npm install
   cd ..
   ```

## Running the Application

### Development Mode (Recommended)

Run both backend and frontend simultaneously:

```bash
npm run dev
```

This will start:

- Backend server on `http://localhost:5000`
- Frontend development server on `http://localhost:3000`

### Production Mode

1. **Build the frontend**

   ```bash
   npm run build
   ```

2. **Start the production server**
   ```bash
   npm start
   ```

The application will be available at `http://localhost:5000`

## Usage Guide

### 1. Upload Excel File

- Drag and drop an Excel file (.xlsx or .xls) into the upload area
- Or click to browse and select a file
- The file should have headers in the first row

### 2. Add Company Data

**Option 1: Import JSON Data (Recommended)**

- Click "Import JSON" button
- Paste your complete company data object (see example below)
- All companies will be imported at once

**Option 2: Add Companies Manually**

- Enter company information in the form:
  - **Company Name** (required)
  - **Mobile Number**
  - **Email ID**
  - **Website**
  - **Address**
- Click "Add Company" to add it to the list
- You can add multiple companies
- Remove individual companies by clicking the trash icon
- Use "Clear All" to remove all companies

### 3. Configure Column Mapping

- Specify the exact column names from your Excel file
- **Company Name Column** is required
- Other columns are optional
- Column names must match exactly (case-sensitive)

### 4. Process and Download

- Click "Process Excel File" to update the Excel file
- Once processing is complete, click "Download Updated Excel"
- The file will be saved with a timestamp

## Example Column Mappings

Based on your existing Excel file, you might use:

- **Company Name Column**: `Company NAME`
- **Mobile Number Column**: `MOBILE NO`
- **Email ID Column**: `EMAIL ID`
- **Website Column**: `WEBSITE`
- **Address Column**: `ADDRESS`

## Example Company Data Format

When using the "Import JSON" feature, use this format:

```json
{
  "AGS FOODS INDIA PRIVATE LIMITED": {
    "MOBILE NO": "9826671476",
    "EMAIL ID": "piyush@agrawalglobal.com",
    "WEBSITE": "",
    "ADDRESS": "Office No. 105, First Floor, Vibrant Business Tower, 9A-9B Manoramaganj, Geeta Bhawan, Indore, Madhya Pradesh – 452001, India"
  },
  "Aico Foods Limited": {
    "MOBILE NO": "",
    "EMAIL ID": "info@aicofoods.com",
    "WEBSITE": "http://aicofoods.com",
    "ADDRESS": "44, Hirabhai Market, Diwan Lalubhai Road, Kankaria Road, Ahmedabad, Gujarat – 380022, India"
  }
}
```

**Note**: Empty strings (`""`) will be ignored during processing, so only fields with actual data will be updated in the Excel file.

## API Endpoints

- `POST /api/process-excel` - Process Excel file with company data
- `GET /api/download/:filename` - Download processed file

## File Structure

```
excel-data/
├── server.js                 # Express server
├── package.json              # Backend dependencies
├── client/                   # React frontend
│   ├── public/
│   ├── src/
│   │   ├── components/       # React components
│   │   │   ├── FileUpload.js
│   │   │   ├── CompanyDataInput.js
│   │   │   └── ColumnMapping.js
│   │   ├── App.js           # Main app component
│   │   ├── index.js         # React entry point
│   │   └── index.css        # Tailwind CSS
│   └── package.json         # Frontend dependencies
├── uploads/                  # Temporary file storage
└── README.md
```

## Features in Detail

### Excel Processing

- Preserves original Excel formatting and styles
- Updates only matching company names
- Maintains font consistency
- Handles large files efficiently

### Data Validation

- Validates file uploads (Excel files only)
- Ensures required fields are filled
- Provides real-time status feedback
- Error handling with user-friendly messages

### User Experience

- Drag & drop file upload
- Real-time form validation
- Loading states and progress indicators
- Toast notifications for user feedback
- Responsive design for all screen sizes

## Troubleshooting

### Common Issues

1. **"Only Excel files are allowed"**

   - Ensure you're uploading a .xlsx or .xls file
   - Check file extension is correct

2. **"Column not found"**

   - Verify column names match exactly (case-sensitive)
   - Check for extra spaces in column names

3. **"No companies updated"**

   - Ensure company names in your data match exactly with Excel
   - Check the company name column mapping

4. **Port already in use**
   - Change the port in `server.js` (line 8)
   - Or kill the process using the port

### Development Tips

- Use browser developer tools to check network requests
- Check server console for detailed error messages
- Verify file permissions for the uploads directory

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## License

This project is open source and available under the [ISC License](LICENSE).
