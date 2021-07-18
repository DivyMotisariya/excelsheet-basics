const excel = require('excel2mysql')

excel({
  input: './sellers.xlsx',
  model: 'create',
  mysql: {
    host: "localhost",
    user: "root",
    password: "",
    port: 3306,
    database: "temp"
  }
}, (err, sql, result) => {
  err && console.log(err);
  sql && console.log(sql);
  result && console.log(result);
})