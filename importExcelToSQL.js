/**import excel data and store in sql with the help of knex query builder through postman form-data */

const express = require("express");
const upload = require("express-fileupload");
const path = require("path");

const knex = require("knex");
const xlsx = require("xlsx");

const app = express();
app.use(express.json());
app.use(upload());

app.post("/upload", (req, res) => {
  const file = req.files.file;
  const uploadFilePath = path.join(__dirname, file.name);

  file.mv(uploadFilePath, (err) => {
    if (err) {
      console.log(err);
    } else {
      importExcelDataToSQL(uploadFilePath);

      res.json({
        message: "File upload successful",
      });
    }
  });
});

//connection knex
const db = knex({
  client: "mysql",
  connection: {
    host: "localhost",
    user: "root",
    password: "", //password
    database: "exceldb",
  },
});

async function importExcelDataToSQL(filePath) {
  let workBook = xlsx.readFile(filePath);
  let workSheet = workBook.Sheets[workBook.SheetNames[0]];
  let range = xlsx.utils.decode_range(workSheet["!ref"]);
  //   console.log(range)

  for (let row = range.s.r + 1; row <= range.e.r; row++) {
    let data = [];
    console.log("range", range);

    for (let col = range.s.c; col <= range.e.c; col++) {
      let cell = workSheet[xlsx.utils.encode_cell({ r: row, c: col })];
      console.log("cell.v-", cell);
      data.push(cell.v);
    }

    // console.log("data", data);

    //for employee table check
    db("employee_data")
      .insert({ 
        employee_id: data[0],
        employee_no: data[1],
        employee_name: data[2],
      })
      .catch((err) => {
        console.log("err", err);
      });

    //for user table another check
    //   db("users").insert({
    //     Name: data[0],
    //     Age:data[1],
    //     Country:data[2]
    //  }).catch((err) =>{
    //      console.log("err",err)
    //  })
  }
}

//importExcelDataToSQL();

app.listen(3000, () => {
  console.log("Server started at post 3000");
});

/**++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ */

/**import excel data and store in sql through postman form-data */

// const express = require("express");
// const path = require("path");
// const xlsx = require("xlsx");
// const upload = require("express-fileupload");
// const con = require("./db/conn");

// const app = express();
// app.use(express.json());
// app.use(upload());

// app.post("/upload", (req, res) => {
//   const file = req.files.file;
//   const uploadFilePath = path.join(__dirname, "./upload/", file.name);

//   file.mv(uploadFilePath, (err) => {
//     if (err) {
//       console.log(err);
//     } else {
//       // importFile(uploadFilePath);

//       res.json({
//         message: "File upload successful",
//       });
//     }
//   });
// });

// function importFile(filePath) {
//   let workBook = xlsx.readFile(filePath);
//   let workSheet = workBook.Sheets[workBook.SheetNames[0]];
//   let range = xlsx.utils.decode_range(workSheet["!ref"]);

//   for (let row = range.s.r + 1; row <= range.e.r; row++) {
//     let data = [];

//     for (let col = range.s.c; col <= range.e.c; col++) {
//       let cell = workSheet[xlsx.utils.encode_cell({ r: row, c: col })];
//       data.push(cell.v);
//     }

//     let sql =
//       "INSERT INTO `postal-data`(`officename`, `officeType`, `Deliverystatus`, `circlename`, `Districtname`, `divisionname`, `regionname`, `Taluk`, `statename`, `pincode`) VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";

//     con.query(
//       sql,
//       [
//         data[0],
//         data[2],
//         data[3],
//         data[6],
//         data[8],
//         data[4],
//         data[5],
//         data[7],
//         data[9],
//         data[1],
//       ],
//       (err, results, fields) => {
//         if (err) return console.log(err);
//       }
//     );
//   }
// }

// app.listen(3000, () => {
//   console.log("Server started at post 3000");
// });
