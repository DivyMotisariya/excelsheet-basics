const reader = require("xlsx");
const file = reader.readFile("./test.xlsx");

let students = [{
    Name: "Divy",
    Age: 22,
    Branch: "Computer",
    Marks: 90,
  },
  {
    Name: "Test",
    Age: 20,
    Branch: "SE",
    Marks: 85,
  },
];

let sheetNo = file.SheetNames.length + 1;
let sheetName = "Sheet" + sheetNo;

const ws = reader.utils.json_to_sheet(students);
reader.utils.book_append_sheet(file, ws, sheetName);
reader.writeFile(file, "./test.xlsx");