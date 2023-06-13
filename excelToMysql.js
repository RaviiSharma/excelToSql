// const mysql = require('mysql');
// const ExcelJS = require('exceljs');


// // Create a MySQL connection
// const connection = mysql.createConnection({
//     host: "localhost",
//     user: "root",
//    //password:"Ravi@036",
//     database: "attendance_dev",
// });

// // Connect to the MySQL database
// connection.connect((err) => {
//   if (err) {
//     console.error('Error connecting to MySQL:', err);
//     return;
//   }

//   console.log('Connected to MySQL database');

//   // Create a new workbook and load the Excel file
//   const workbook = new ExcelJS.Workbook();
//   workbook.xlsx.readFile('output1.xlsx')
//     .then(() => {
//       // Get the first worksheet from the workbook
//       const worksheet = workbook.getWorksheet(1);

//       // Get the column headers from the first row of the worksheet
//       const columnHeaders = worksheet.getRow(1).values;
      

//       // Prepare the SQL query to insert data into the MySQL table
//       const tableName = 'ptr_employees';
//       //const insertQuery = `INSERT INTO ${tableName} (${columnHeaders.join(',')}) VALUES ?`;
//       const insertQuery = `INSERT INTO ${tableName} (\`${columnHeaders.join('`,`')}\`) VALUES ?`;

//       // Prepare an array to store the data rows
//       const dataRows = [];

//       // Iterate over the worksheet rows (excluding the first row which contains column headers)
//       for (let i = 2; i <= worksheet.rowCount; i++) {
//         const row = worksheet.getRow(i);
//         const rowData = row.values;

//         // Skip empty rows
//         if (rowData.some(cellValue => cellValue !== null)) {
//           dataRows.push(rowData.slice(1)); // Exclude the first column if needed
//         }
//       }

//       // Insert the data into the MySQL table
//       connection.query(insertQuery, [dataRows], (err, result) => {
//         if (err) {
//           console.error('Error inserting data:', err);
//         } else {
//           console.log(`Inserted ${result.affectedRows} rows into ${tableName}`);
//         }

//         // Close the MySQL connection
//         connection.end();
//       });
//     })
//     .catch((err) => {
//       console.error('Error reading Excel file:', err);
//     });
// });



//
const mysql = require('mysql');
const xlsx=require('xlsx');

// Create a MySQL connection
const db = mysql.createConnection({
    host: "localhost",
    user: "root",
  // password:"Ravi@036", //password
    database: "exceldb",
});

let workbook=xlsx.readFile('users.xlsx');
let worksheet = workbook.Sheets[workbook.SheetNames[0]]; //get first sheet
let range =xlsx.utils.decode_range(worksheet["!ref"])

for(let row = range.s.r;row<=range.e.r;row++){
    let data =[]

    for(let col=range.s.c;col<=range.e.c;col++){ //innner loop
    let cell = worksheet[xlsx.utils.encode_cell({r:row,c:col})]
    data.push(cell.v) 
    }
    console.log("data present or not",data)

    let sql="INSERT INTO `users` (`Name`,`Age`,`Country`) VALUES(?,?,?)"

    db.query(sql,data,(err,results,fields)=>{

        if(err){
            return console.error(err.message)
        }
        console.log('User ID:' + results.insertId)
    })
}
db.end(); //to end the connction
