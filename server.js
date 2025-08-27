const express = require("express");
const cors = require("cors");
const multer = require("multer");
const path = require("path");
const ExcelJS = require("exceljs");
const fs = require("fs");

const app = express();
const PORT = process.env.PORT || 5000;

// Middleware
app.use(cors());
app.use(express.json({ limit: "50mb" }));
app.use(express.urlencoded({ extended: true, limit: "50mb" }));

// Create uploads directory if it doesn't exist
const uploadsDir = path.join(__dirname, "uploads");
if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir);
}

// Configure multer for file uploads - use memory storage for Vercel
const storage = multer.memoryStorage();

const upload = multer({
  storage: storage,
  limits: {
    fileSize: 10 * 1024 * 1024, // 10MB limit for Vercel free tier
  },
  fileFilter: function (req, file, cb) {
    const allowedTypes = [".xlsx", ".xls"];
    const fileExt = path.extname(file.originalname).toLowerCase();
    if (allowedTypes.includes(fileExt)) {
      cb(null, true);
    } else {
      cb(new Error("Only Excel files are allowed!"), false);
    }
  },
});

// Routes
app.get("/api/test", (req, res) => {
  res.json({
    message: "Hello World",
  });
});

app.post("/api/process-excel", upload.single("excelFile"), async (req, res) => {
  try {
    const { companyData, columnMappings } = req.body;
    const excelFile = req.file;

    if (!excelFile) {
      return res.status(400).json({ error: "No Excel file uploaded" });
    }

    if (!companyData || !columnMappings) {
      return res
        .status(400)
        .json({ error: "Company data and column mappings are required" });
    }

    // Parse the JSON data
    const parsedCompanyData = JSON.parse(companyData);
    const parsedColumnMappings = JSON.parse(columnMappings);

    // Process the Excel file from memory buffer
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(excelFile.buffer);
    const sheet = workbook.worksheets[0]; // first sheet

    // Find column indexes by header name
    const headerRow = sheet.getRow(1);
    const colIndex = {};
    headerRow.eachCell((cell, colNumber) => {
      const headerValue = cell.value ? cell.value.toString().trim() : "";
      if (headerValue) {
        colIndex[headerValue] = colNumber;
      }
    });

    // Get the first company data to extract column names
    const firstCompanyName = Object.keys(parsedCompanyData)[0];
    const firstCompanyData = parsedCompanyData[firstCompanyName];

    // Auto-detect column names from the data structure
    const dataColumns = Object.keys(firstCompanyData);
    const companyNameColumn = parsedColumnMappings.companyNameColumn;

    console.log("Found columns in Excel:", Object.keys(colIndex));
    console.log("Data columns from JSON:", dataColumns);
    console.log("Company name column:", companyNameColumn);

    // Get font style from A2 (row 2, col 1) for consistency
    const referenceCell = sheet.getRow(2).getCell(1);
    const referenceFont = referenceCell.font;

    // Loop through each row and update if company matches
    let updatedRows = 0;
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // skip header

      // Check if the company name column exists in the row
      const companyNameColumnName = parsedColumnMappings.companyNameColumn;
      const companyNameColumnIndex = colIndex[companyNameColumnName];
      if (!companyNameColumnIndex) {
        console.log(
          `Company name column "${companyNameColumnName}" not found in available columns:`,
          Object.keys(colIndex)
        );
        return;
      }

      // Safely get the company name cell
      let companyName = null;
      try {
        const companyNameCell = row.getCell(companyNameColumnIndex);
        companyName = companyNameCell ? companyNameCell.value : null;
      } catch (error) {
        console.log(
          `Error getting company name from row ${rowNumber}:`,
          error.message
        );
        return;
      }

      if (companyName && parsedCompanyData[companyName]) {
        const data = parsedCompanyData[companyName];
        console.log(`Processing company: ${companyName}`);
        console.log(`Company data:`, data);

        // Update each data column - only update if the value is not empty
        dataColumns.forEach((columnName) => {
          const columnIndex = colIndex[columnName];
          const dataValue = data[columnName];

          console.log(
            `Checking column: ${columnName} (index: ${columnIndex}), value: ${dataValue}`
          );

          // Only update if the value exists and is not empty and column exists
          if (dataValue && dataValue.trim() !== "" && columnIndex) {
            try {
              const cell = row.getCell(columnIndex);
              cell.value = dataValue;
              console.log(
                `Updated cell at row ${rowNumber}, column ${columnName} with value: ${dataValue}`
              );

              // Apply font from reference cell for consistency
              if (referenceFont) {
                cell.font = referenceFont;
              }
            } catch (error) {
              console.log(
                `Error updating cell in row ${rowNumber}, column ${columnName}:`,
                error.message
              );
            }
          } else {
            console.log(
              `Skipping column ${columnName}: value="${dataValue}", columnIndex=${columnIndex}`
            );
          }
        });
        updatedRows++;
      }
    });

    // Generate the Excel file in memory and send directly
    const buffer = await workbook.xlsx.writeBuffer();

    // Convert buffer to base64 for sending
    const base64Data = buffer.toString("base64");

    res.json({
      success: true,
      message: `Excel file processed successfully! Updated ${updatedRows} rows.`,
      fileData: base64Data,
      filename: `updated_excel_${Date.now()}.xlsx`,
      mimeType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
  } catch (error) {
    console.error("Error processing Excel file:", error);
    res
      .status(500)
      .json({ error: "Error processing Excel file: " + error.message });
  }
});

// Download route removed - files are now sent directly in the response

// Health check route
app.get("/api/health", (req, res) => {
  res.json({
    status: "OK",
    message: "Server is running",
    timestamp: new Date().toISOString(),
    environment: process.env.NODE_ENV || "development",
  });
});

// Serve static files in production
if (process.env.NODE_ENV === "production") {
  app.use(express.static(path.join(__dirname, "client/build")));

  app.get("*", (req, res) => {
    res.sendFile(path.join(__dirname, "client/build", "index.html"));
  });
}

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
