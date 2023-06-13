const mysql = require('mysql');
const ExcelJS = require('exceljs');


// Create a MySQL connection
const connection = mysql.createConnection({

  host: "localhost",
  user: "root",
 //password:"Ravi@036",
  database: "attendance_dev",
});

// Connect to the MySQL database
connection.connect((err) => {
  if (err) {
    console.error('Error connecting to MySQL:', err);
    return;
  }

  console.log('Connected to MySQL database');

  // Define the SQL query to retrieve data from the table
  const query = 'SELECT * FROM ptr_employees'; 

  // Execute the query
  connection.query(query, (err, rows) => {
    if (err) {
      console.error('Error executing query:', err);
      return;
    }

    // Create a new workbook and sheet using ExcelJS
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Sheet 1');

    // Add column headers to the worksheet
    const columnHeaders = Object.keys(rows[0]);
    worksheet.addRow(columnHeaders);

    // Add data rows to the worksheet
    rows.forEach((row) => {
      const rowData = Object.values(row);
      worksheet.addRow(rowData);
    });

    // Save the workbook to an Excel file
    const filename = 'output.xlsx';
    workbook.xlsx.writeFile(filename)
      .then(() => {
        console.log(`Data successfully exported to ${filename}`);
      })
      .catch((err) => {
        console.error('Error exporting data:', err);
      })
      .finally(() => {
        // Close the MySQL connection
        connection.end();
      });
  });
});

