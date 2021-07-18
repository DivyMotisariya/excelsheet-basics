const ExcelJS = require("exceljs");
const mysql = require("mysql");
const connection = mysql.createConnection({
  host: "localhost",
  user: "root",
  password: "",
  database: "sequelize_ims",
});

connection.connect((err, args) => {
  if (err) return console.log(err);
  // console.log(args)
});

let workbook = new ExcelJS.Workbook();

// connection.query("select * from sellers;", (err, rows, fields) => {
//   if (err) return console.log(err);

//   let worksheet = workbook.addWorksheet("Sellers");

//   let columns = [];
//   fields.forEach((column) => {
//     columns.push({ name: column.name, filterButton: true });
//   });

//   let rowArr = [];
//   rows.forEach((row) => {
//     rowArr.push([row.id, row.name.trim(), Number(row.contact_no), row.address]);
//   });

//   let name = fields[0].table;
//   console.log(name);
//   worksheet
//     .addTable({
//       name: name,
//       ref: "A1",
//       headerRow: true,
//       totalsRow: true,
//       style: {
//         theme: "TableStyleDark3",
//         showRowStripes: true,
//       },
//       columns: columns,
//       rows: rowArr,
//     })
//     // .commit();

//   workbook.xlsx
//     .writeFile("./sellers.xlsx", { stream: worksheet })
//     .then(() => {
//       return console.log("Saved");
//     })
//     .catch((err) => console.log(err));
// });

// workbook.xlsx.readFile("./sellers.xlsx").then((wb) => {
//   let sheet = wb.getWorksheet("Sellers")
//   let tables = sheet.getTables();
//   tables.forEach(table => {
//     let tableName = table.table.name;
//     let cols = [];
//     table.table.columns.forEach(column => cols.push(column.name))
//     let SQL = `CREATE TABLE ${tableName}(${cols.map((col) => col = col + ' varchar(100)').join(', ')});`;
//     console.log(SQL);
//   })
// });

// workbook.xlsx
//   .readFile("./sellers.xlsx")
//   .then((wb) => {
//     let sheet = wb.getWorksheet("Sellers");
//     // let tableName = sheet.getTables();
//     // console.log(tableName);

//     // sheet.columns.forEach((col) => console.log(col.values[1]));

//     let rows = sheet.getRows(2, sheet.rowCount - 2);
//     let rowArr = [];
//     // rows.forEach((row) => rowArr.push({ id: row.values[1], name: row.values[2], contact_no: row.values[3], address: row.values[4] }))
//     rows.forEach((row) => {
//       connection.query("INSERT INTO sellers SET ?", {
//         // id: row.values[1],
//         name: row.values[2],
//         contact_no: row.values[3],
//         address: row.values[4],
//       }, (err, res) => {
//         if (err) return console.log(err);
//         console.log(res);
//       });
//     });
//     // console.log(rowArr);

//     // connection.query(`INSERT INTO sellers values ?`, rowArr, (err, result) => {
//     //   if (err) return console.log(err);
//     //   console.log(result);
//     // });
//   })
//   .catch((err) => console.log(err));

// connection.query("SELECT * FROM sellers;", (err, sellers, fields) => {
//   // console.log(sellers);
//   const jsonSellers = JSON.parse(JSON.stringify(sellers));
//   // console.log(jsonSellers);
//   let workbook = new ExcelJS.Workbook();
//   let worksheet = workbook.addWorksheet("Sellers");

//   worksheet.columns = [
//     { header: "Id", key: "id", width: 10 },
//     { header: "Name", key: "name", width: 30 },
//     { header: "Contact", key: "contact_no", width: 30 },
//     { header: "Address", key: "address", width: 50 },
//   ];

//   worksheet.addRows(jsonSellers);

//   workbook.xlsx
//     .writeFile("./sellers.xlsx")
//     .then(() => console.log("Saved"))
//     .catch((err) => console.log(err));
// });

/**
 * Read file to csv
 */

// workbook.csv
//   .readFile("./test.csv")
//   .then((data) => {
//     // console.log(data);
//     data.columns.forEach(col => console.log(col.toString()))
//       // let worksheet = workbook.getWorksheet(1);
//       // console.log(worksheet);
//       // workbook.csv
//       //   .writeFile("./test.csv", { sheetName: worksheet.name })
//       //   .then((value) => console.log(value))
//       //   .catch((err) => console.log(err));
//   })
//   .catch((err) => console.log(err));

/**
 * Read file and write table from mysql db to file
 */

// workbook.xlsx
//   .readFile("./test.xlsx")
//   .then(() => {
//     let worksheet = workbook.getWorksheet(1);
//     if (!worksheet) return console.log("No worksheet with this name");
//     connection.query("select * from customers;", (err, rows, fields) => {
//       if (err) return console.log(err);

//       let columns = [];
//       fields.forEach((column) => {
//         columns.push({ name: column.name, filterButton: true });
//       });

//       let rowArr = [];
//       rows.forEach((row) => {
//         rowArr.push([row.id, row.name.trim(), Number(row.contact_no), row.address]);
//       });

//       let name = fields[0].table;
//       console.log(name);
//       worksheet
//         .addTable({
//           name: name,
//           ref: "A1",
//           headerRow: true,
//           totalsRow: true,
//           style: {
//             theme: "TableStyleDark3",
//             showRowStripes: true,
//           },
//           columns: columns,
//           rows: rowArr,
//         })
//         .commit();

//       workbook.xlsx
//         .writeFile("./test.xlsx", { stream: worksheet })
//         .then(() => {
//           return console.log("Saved");
//         })
//         .catch((err) => console.log(err));
//     });
//   })
//   .catch((err) => console.log(err.message));