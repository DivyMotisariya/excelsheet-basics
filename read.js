const reader = require("xlsx");
const file = reader.readFile("./test.xlsx");

let data = [];

const sheets = file.SheetNames;

for (let i = 0; i < sheets.length; ++i) {
  data.push({ Sheet$: sheets[i].substring(5) });
  const temp = reader.utils.sheet_to_json(file.Sheets[file.SheetNames[i]]);
  temp.forEach((res) => data.push(res));
}

console.log(data);